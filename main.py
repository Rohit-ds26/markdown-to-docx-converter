"""
Markdown → DOCX converter. Dependencies are verified on startup; missing or
outdated packages are installed automatically via pip when you run this file.

Skip auto-install: python main.py --skip-deps ...
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path


# Canonical dependency list — auto-installed on first run. Keep requirements.txt in sync.
PIP_REQUIREMENTS: tuple[str, ...] = (
    "anthropic>=0.34.0",
    "python-docx>=1.1.2",
    "markdown-it-py>=4.1.0",
    "mdit-py-plugins>=0.4.2",
    "python-dotenv>=1.0.1",
    "pypandoc_binary>=1.13",
)

PROJECT_ROOT = Path(__file__).resolve().parent
REQUIREMENTS_FILE = PROJECT_ROOT / "requirements.txt"
ENV_FILE = PROJECT_ROOT / ".env"
OUTPUT_DIR = PROJECT_ROOT / "output"

# Invalid / deprecated model names → treat as empty (auto-pick from API account).
_INVALID_MODEL_ALIASES: frozenset[str] = frozenset(
    {"claude-3-5-haiku-latest", "claude-3-5-haiku", "claude-3-5-haiku-20241022"}
)


def resolve_anthropic_model(model: str) -> str:
    """Return model id from env/CLI, or '' to auto-pick from the user's Anthropic account."""
    m = (model or "").strip()
    if m in _INVALID_MODEL_ALIASES:
        return ""
    return m


def _parse_version_triple(version: str) -> tuple[int, int, int]:
    parts: list[int] = []
    for token in version.split(".")[:3]:
        num = "".join(ch for ch in token if ch.isdigit())
        parts.append(int(num) if num else 0)
    while len(parts) < 3:
        parts.append(0)
    return (parts[0], parts[1], parts[2])


def _dependencies_ok() -> bool:
    """True when every package used by the app imports (and markdown-it-py >= 4.1)."""
    try:
        import anthropic  # noqa: F401
        import docx  # noqa: F401
        import markdown_it
        import pypandoc  # noqa: F401
        from dotenv import load_dotenv  # noqa: F401
        from mdit_py_plugins.gfm import gfm_plugin  # noqa: F401

        if _parse_version_triple(getattr(markdown_it, "__version__", "0")) < (4, 1, 0):
            return False
        return True
    except ImportError:
        return False


def _pip_install_command() -> list[str]:
    """Prefer requirements.txt next to main.py; fall back to PIP_REQUIREMENTS."""
    if REQUIREMENTS_FILE.is_file():
        return ["-r", str(REQUIREMENTS_FILE)]
    return list(PIP_REQUIREMENTS)


def ensure_dependencies(*, quiet_pip: bool = True, skip: bool = False) -> None:
    """Install PIP_REQUIREMENTS via pip only when imports fail or Markdown-it is too old."""
    if skip:
        return
    if _dependencies_ok():
        return

    print("Installing or updating Python dependencies (one-time)...")
    cmd = [
        sys.executable,
        "-m",
        "pip",
        "install",
        "--disable-pip-version-check",
        *(("-q",) if quiet_pip else ()),
        *PIP_REQUIREMENTS,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        err = (result.stderr or result.stdout or "").strip()
        raise SystemExit(
            "Automatic pip install failed.\n"
            f"Try manually: python -m pip install -r requirements.txt\n\n{err}"
        )

    if not _dependencies_ok():
        raise SystemExit(
            "Dependencies still not usable after pip install.\n"
            "Run: python -m pip install -r requirements.txt\n"
            "Or create a virtual environment (.venv) and install there."
        )
    print("Dependencies ready.")


def reveal_saved_docx(output_path: Path) -> None:
    """
    Make the DOCX obvious on disk. VS Code / Cursor won't show Word files as plain text.

    On Windows we open File Explorer with the file selected. Other OS: print the path only.
    """
    p = Path(output_path).expanduser().resolve()
    if not p.is_file():
        return

    if sys.platform == "win32":
        try:
            # explorer /select,<path> — highlights the file in a folder window
            subprocess.Popen(["explorer.exe", "/select,", str(p)], close_fds=True)
            print("(Opened File Explorer with your DOCX selected. Double‑click it to open in Word.)")
        except Exception as e:
            print(f"Could not open File Explorer ({type(e).__name__}: {e}).")
    print(f"You can always open from disk: {p}")


def open_docx_with_default_app(output_path: Path) -> None:
    """Launch the system's default handler for .docx (often Microsoft Word on Windows)."""
    p = Path(output_path).expanduser().resolve()
    if not p.is_file():
        return
    try:
        if sys.platform == "win32":
            os.startfile(str(p))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.run(["open", str(p)], check=False)
        else:
            subprocess.run(["xdg-open", str(p)], check=False)
    except Exception as e:
        print(f"Could not open DOCX with default app ({type(e).__name__}: {e}).")


def pick_markdown_file() -> Path:
    # Tkinter file picker (no frontend app required).
    print("Opening file picker... (select a .md file)")
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        try:
            selected = filedialog.askopenfilename(
                title="Select a Markdown (.md) file",
                filetypes=[("Markdown files", "*.md"), ("All files", "*.*")],
            )
        finally:
            root.destroy()
    except Exception as e:
        print(f"Could not open file picker ({type(e).__name__}: {e}).")
        selected = input("Paste full path to a .md file and press Enter: ").strip().strip('"')

    if not selected:
        raise SystemExit("No file selected. Cancelled.")

    p = Path(selected).expanduser().resolve()
    if p.suffix.lower() != ".md":
        raise SystemExit(f"Selected file is not a .md file: {p}")
    return p


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Convert Markdown (.md) to DOCX with optional Claude cleanup.")
    p.add_argument(
        "input",
        nargs="?",
        default=None,
        help="Optional path to the input .md file (if omitted, a file picker opens)",
    )
    p.add_argument(
        "--no-ai",
        action="store_true",
        help="Skip AI formatting step and convert raw Markdown to DOCX",
    )
    p.add_argument(
        "--model",
        default=os.getenv("ANTHROPIC_MODEL", ""),
        help="Anthropic model id (default: env ANTHROPIC_MODEL, or auto-pick if unset)",
    )
    p.add_argument(
        "--engine",
        choices=["auto", "pypandoc", "pandoc", "builtin"],
        default=os.getenv("MD_TO_DOCX_ENGINE", "auto"),
        help="DOCX engine: auto (bundled Pandoc if installed), pypandoc, pandoc, builtin",
    )
    p.add_argument(
        "--save",
        choices=["output", "same"],
        default=os.getenv("MD_TO_DOCX_SAVE_MODE", "output"),
        help='Where to save the DOCX: "output" (local output/) or "same" (same folder as .md)',
    )
    p.add_argument(
        "--skip-deps",
        action="store_true",
        help="Do not auto-install/update dependencies (assume already installed).",
    )
    p.add_argument(
        "--no-reveal",
        action="store_true",
        help="After saving, do not open Windows File Explorer on the DOCX file.",
    )
    p.add_argument(
        "--open-docx",
        action="store_true",
        help="After saving, open the DOCX with the default application (e.g. Word).",
    )
    return p.parse_args()


def main() -> int:
    # Resolve --skip-deps before argparse so deps can bootstrap even on first install.
    skip_deps = "--skip-deps" in sys.argv
    ensure_dependencies(skip=skip_deps)

    from dotenv import load_dotenv

    # Load .env next to this script — not cwd. Otherwise running from another folder
    # (common on Windows) skips ANTHROPIC_API_KEY and Claude appears "broken".
    load_dotenv(ENV_FILE)

    args = parse_args()

    # Project modules touch third-party packages; load only after deps are OK.
    from agents.reader_agent import read_markdown
    from agents.formatter_agent import format_markdown_with_claude
    from agents.docx_agent import markdown_to_docx
    from agents.usage_agent import format_usage_summary

    try:
        input_path = Path(args.input).expanduser().resolve() if args.input else pick_markdown_file()
        print(f"Selected: {input_path}")
    except SystemExit:
        raise
    except Exception as e:
        raise SystemExit(f"Failed to select/read input file: {type(e).__name__}: {e}")

    if args.save == "same":
        out_path = (input_path.parent / f"{input_path.stem}.docx").resolve()
    else:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        out_path = (OUTPUT_DIR / f"{input_path.stem}.docx").resolve()

    try:
        md = read_markdown(str(input_path))
    except Exception as e:
        raise SystemExit(f"Failed to read Markdown: {type(e).__name__}: {e}")

    usage = {"input_tokens": 0, "output_tokens": 0, "estimated_cost_usd": None, "provider": "anthropic"}
    processed_md = md

    if not args.no_ai:
        api_key = (os.getenv("ANTHROPIC_API_KEY") or "").strip()
        if not api_key:
            raise SystemExit(
                "ANTHROPIC_API_KEY is not set. Set it in your environment, or run with --no-ai."
            )
        model = resolve_anthropic_model(args.model)
        if model:
            print(f"AI formatting enabled (model: {model})...")
        else:
            print("AI formatting enabled (model: auto — picking from your Anthropic account)...")
        try:
            processed_md, usage = format_markdown_with_claude(
                markdown=md,
                api_key=api_key,
                model=model,
            )
        except Exception as e:
            hint = (
                "\n\nTip: list models your key can use:\n"
                "  python -c \"from anthropic import Anthropic; import os; from dotenv import load_dotenv; "
                "load_dotenv('.env'); c=Anthropic(api_key=os.environ['ANTHROPIC_API_KEY']); "
                "print([m.id for m in c.models.list()])\""
            )
            raise SystemExit(f"Claude formatting failed: {type(e).__name__}: {e}{hint}")
    else:
        print("AI formatting skipped (--no-ai).")

    print(f"Converting to DOCX (requested: {args.engine})...")
    try:
        used_engine = markdown_to_docx(
            markdown=processed_md, output_path=str(out_path), engine=args.engine
        )
        print(f"DOCX engine used: {used_engine}")
    except Exception as e:
        raise SystemExit(f"DOCX conversion failed: {type(e).__name__}: {e}")

    print(format_usage_summary(output_docx_path=str(out_path.absolute()), usage=usage))
    print("\nDOCX files are ZIP/binary — the editor preview may say 'binary'; the file still exists on disk.\n")

    if args.open_docx:
        print("Opening DOCX with default app...")
        open_docx_with_default_app(out_path)
    elif sys.platform == "win32" and not args.no_reveal:
        print("Showing file in Explorer...")
        reveal_saved_docx(out_path)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

