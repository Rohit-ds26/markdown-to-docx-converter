from __future__ import annotations

import os


def _money(value: float | None) -> str:
    if value is None:
        return "N/A"
    return f"${value:.4f}".rstrip("0").rstrip(".")


def estimate_anthropic_cost_usd(*, model: str | None, input_tokens: int, output_tokens: int) -> float | None:
    """
    Anthropic does not return cost in API responses, so we estimate.

    Prefer env overrides so you can match your exact pricing/plan:
      - ANTHROPIC_INPUT_PER_MILLION_USD
      - ANTHROPIC_OUTPUT_PER_MILLION_USD
    """

    env_in = os.getenv("ANTHROPIC_INPUT_PER_MILLION_USD")
    env_out = os.getenv("ANTHROPIC_OUTPUT_PER_MILLION_USD")
    if env_in and env_out:
        try:
            in_rate = float(env_in)
            out_rate = float(env_out)
            return (input_tokens / 1_000_000) * in_rate + (output_tokens / 1_000_000) * out_rate
        except Exception:
            return None

    # If you don't set env pricing, we still report tokens. Cost stays N/A by default.
    return None


def format_usage_summary(*, output_docx_path: str, usage: dict) -> str:
    input_tokens = int(usage.get("input_tokens", 0) or 0)
    output_tokens = int(usage.get("output_tokens", 0) or 0)
    model = usage.get("model")

    estimated = usage.get("estimated_cost_usd")
    if estimated is None and usage.get("provider") == "anthropic":
        estimated = estimate_anthropic_cost_usd(model=model, input_tokens=input_tokens, output_tokens=output_tokens)

    lines = [
        "-" * 32,
        "Conversion Successful",
        f"Input Tokens: {input_tokens}",
        f"Output Tokens: {output_tokens}",
        f"Estimated Cost: {_money(estimated)}",
        f"DOCX Saved: {output_docx_path}",
        "-" * 32,
    ]
    return "\n".join(lines)

