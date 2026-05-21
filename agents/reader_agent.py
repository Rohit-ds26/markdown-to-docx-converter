from __future__ import annotations

from pathlib import Path


def read_markdown(path: str) -> str:
    p = Path(path).expanduser().resolve()
    if not p.exists():
        raise FileNotFoundError(f"Input file not found: {p}")
    if p.suffix.lower() != ".md":
        raise ValueError(f"Input must be a .md file: {p}")

    content = p.read_text(encoding="utf-8")
    if not content.strip():
        raise ValueError(f"Input Markdown is empty: {p}")
    return content

