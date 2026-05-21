from __future__ import annotations

from anthropic import Anthropic, NotFoundError


SYSTEM_PROMPT = """You are a senior technical writer producing publication-ready Markdown for Microsoft Word.

Transform the draft into a rich, structured report — not a lightly cleaned version of the same text.

Structure (required when content allows):
- Exactly one `#` document title; optional subtitle as normal text under the title.
- `##` for every major section; `###` for subsections. Never skip levels.
- After the title, add a **Executive summary** or **Overview** section (2–4 sentences).
- Before dense sections, add a one-line lead-in paragraph explaining what follows.
- End with **Conclusion** or **Recommendations** when the source material supports it.
- Insert `---` between major `##` sections (not after every paragraph).

Visual richness (use generously):
- **Bold** all severity labels, metrics, CVE IDs, hostnames, and decision outcomes.
- Blockquotes (`>`) for warnings, critical findings, and executive callouts (one per important item).
- Bullet lists for findings, steps, and requirements; numbered lists for procedures and ranked items.
- Markdown tables for any tabular or comparison data (Severity | Finding | Impact | Status).
- Short sub-bullets under main bullets when the source has nested detail.

Quality:
- Fix grammar and tone to formal report style; expand terse bullets into clear sentences where needed.
- Merge duplicate headings; split walls of text into scannable subsections.
- Keep every fact, number, command, path, and code sample from the source.

Output rules:
- Return ONLY Markdown (no preamble, no ```markdown fence around the whole file).
- Do not invent facts. Do not delete technical content.
- Preserve code fences and their contents unchanged.
- No HTML tags.
"""

def _strip_markdown_fences(text: str) -> str:
    """Claude sometimes wraps the whole document in ```markdown ... ```."""
    t = text.strip()
    if t.startswith("```"):
        lines = t.splitlines()
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        t = "\n".join(lines).strip()
    return t


def _safe_int(value) -> int:
    try:
        return int(value or 0)
    except Exception:
        return 0


def list_account_model_ids(*, api_key: str) -> list[str]:
    """Return model ids your API key can use (empty if listing fails)."""
    client = Anthropic(api_key=api_key)
    ids: list[str] = []
    try:
        for item in client.models.list(limit=100):
            mid = getattr(item, "id", None)
            if mid:
                ids.append(str(mid))
    except Exception:
        pass
    return ids


def pick_model_for_account(*, api_key: str, preferred: str) -> str:
    """
    Choose a model id for this API key.
    If preferred is set and valid, use it. Otherwise auto-pick: Haiku → Sonnet → Opus → any.
    """
    preferred = (preferred or "").strip()
    available = list_account_model_ids(api_key=api_key)

    if preferred:
        if not available or preferred in available:
            return preferred

    if not available:
        return preferred  # may be empty; caller handles missing model

    # Rich document restructuring needs Sonnet/Opus; Haiku is last resort.
    for pattern in ("sonnet", "opus", "haiku"):
        for mid in available:
            if pattern in mid.lower():
                return mid
    return available[0]


def _models_to_try(preferred: str, account_models: list[str]) -> list[str]:
    """Ordered model ids to attempt: preferred first, then rest of account models."""
    seen: set[str] = set()
    ordered: list[str] = []

    def add(m: str) -> None:
        m = m.strip()
        if m and m not in seen:
            seen.add(m)
            ordered.append(m)

    add(preferred)
    for m in account_models:
        add(m)
    return ordered


def _create_message(
    *,
    client: Anthropic,
    model: str,
    markdown: str,
    messages: list[dict] | None = None,
):
    if messages is None:
        messages = [
            {
                "role": "user",
                "content": (
                    "Rewrite this into a full professional report in Markdown. "
                    "Add executive summary, clear section hierarchy, tables, callouts, "
                    "and horizontal rules between major sections. Output only Markdown:\n\n"
                    f"{markdown}"
                ),
            }
        ]
    return client.messages.create(
        model=model,
        max_tokens=16384,
        temperature=0.2,
        system=SYSTEM_PROMPT,
        messages=messages,
    )


def _text_from_message(msg) -> str:
    parts: list[str] = []
    for block in msg.content:
        if getattr(block, "type", None) == "text":
            parts.append(block.text)
    return _strip_markdown_fences("".join(parts).strip())


def _format_with_continuation(
    *,
    client: Anthropic,
    model: str,
    markdown: str,
) -> tuple[str, object]:
    """Run Claude; if output hits max_tokens, request one continuation chunk."""
    msg = _create_message(client=client, model=model, markdown=markdown)
    out = _text_from_message(msg) or markdown
    total_usage = getattr(msg, "usage", None)

    if getattr(msg, "stop_reason", None) != "max_tokens":
        return out, msg

    print("  (Document long — continuing formatting...)")
    cont = _create_message(
        client=client,
        model=model,
        markdown=markdown,
        messages=[
            {
                "role": "user",
                "content": (
                    "Rewrite this into a full professional report in Markdown. "
                    "Output only Markdown:\n\n"
                    f"{markdown}"
                ),
            },
            {"role": "assistant", "content": out},
            {
                "role": "user",
                "content": (
                    "Your previous reply was cut off at the token limit. "
                    "Continue EXACTLY where you stopped. Output only the remaining "
                    "Markdown (no repetition of earlier sections)."
                ),
            },
        ],
    )
    tail = _text_from_message(cont)
    if tail:
        out = f"{out.rstrip()}\n\n{tail.lstrip()}"
    return out, cont


def format_markdown_with_claude(*, markdown: str, api_key: str, model: str) -> tuple[str, dict]:
    client = Anthropic(api_key=api_key)
    account_models = list_account_model_ids(api_key=api_key)
    resolved = pick_model_for_account(api_key=api_key, preferred=model)

    if not model.strip() and resolved:
        print(f"Auto-selected model: {resolved!r}")
    elif model.strip() and resolved != model.strip():
        print(f"Using model available on your account: {resolved!r} (configured: {model!r})")

    candidates = _models_to_try(resolved, account_models)
    if not candidates:
        raise RuntimeError(
            "No Claude model available. Set ANTHROPIC_MODEL in .env to a model id from "
            "console.anthropic.com, or check your API key."
        )

    last_error: Exception | None = None

    for attempt_model in candidates:
        try:
            out_md, msg = _format_with_continuation(
                client=client, model=attempt_model, markdown=markdown
            )
            used_model = attempt_model
            break
        except NotFoundError as e:
            last_error = e
            if attempt_model != candidates[-1]:
                print(f"Model not available ({attempt_model!r}), trying next...")
            continue
    else:
        raise RuntimeError(
            "No Claude model worked with your API key (all returned 404).\n"
            f"Last error: {last_error}\n"
            "Set ANTHROPIC_MODEL in .env to a model id from your Anthropic console, "
            "or run: python -c \"from anthropic import Anthropic; import os; "
            "c=Anthropic(api_key=os.environ['ANTHROPIC_API_KEY']); "
            "print([m.id for m in c.models.list()])\""
        ) from last_error

    if not out_md.strip():
        out_md = markdown

    usage = getattr(msg, "usage", None)
    input_tokens = _safe_int(getattr(usage, "input_tokens", 0))
    output_tokens = _safe_int(getattr(usage, "output_tokens", 0))

    return out_md, {
        "provider": "anthropic",
        "model": used_model,
        "input_tokens": input_tokens,
        "output_tokens": output_tokens,
        "estimated_cost_usd": None,
    }
