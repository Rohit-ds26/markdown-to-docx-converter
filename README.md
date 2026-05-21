# Markdown → DOCX Converter (Claude optional)

Convert a Markdown file to a polished Word document locally. Optional Claude API step improves grammar and report-style layout before conversion.

---

## Quick start (after `git clone`)

### What you need

- **Python 3.10+** ([python.org](https://www.python.org/downloads/)) — on Windows, check **“Add Python to PATH”** during install
- **Git**
- **Anthropic API key** (only if you want AI formatting) — [console.anthropic.com](https://console.anthropic.com/)

### Steps

**1. Clone the repository**

```powershell
git clone <YOUR_REPO_URL>
cd "Docx Convertor"
```

**2. (Recommended) Create a virtual environment**

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

If PowerShell blocks activation:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

**3. Configure API key (for AI formatting)**

```powershell
copy .env.example .env
```

Edit `.env` and set your key:

```text
ANTHROPIC_API_KEY=sk-ant-...
```

> **Without an API key:** you can still convert files using `python main.py --no-ai` (no Claude step).

**4. Run the app**

```powershell
python main.py
```

- A **file picker** opens → choose your `.md` file  
- The script **installs missing packages automatically** on first run (may take a minute)  
- When finished, the DOCX is saved to **`output\<filename>.docx`**  
- On Windows, **File Explorer** opens with the file selected  

**5. Open the result**

Double-click the `.docx` in the `output` folder (use **Microsoft Word** or similar — not the VS Code text editor).

### One-line summary

| Step | Command / action |
|------|------------------|
| Clone | `git clone <repo>` → `cd "Docx Convertor"` |
| Venv (optional) | `python -m venv .venv` → `.\.venv\Scripts\Activate.ps1` |
| API key | `copy .env.example .env` → paste `ANTHROPIC_API_KEY` |
| Run | `python main.py` → pick `.md` file |
| Output | `output\yourfile.docx` |

### Common commands

```powershell
python main.py                  # file picker + Claude + DOCX
python main.py --no-ai          # convert only, no API
python main.py --open-docx      # open DOCX in Word after conversion
python main.py --save same       # save DOCX next to the .md file
python main.py --skip-deps      # skip automatic pip install
```

---

## How it works

1. Read a `.md` file (file picker or optional path argument)
2. (Optional) Polish Markdown using Claude (Anthropic API)
3. Convert Markdown → `.docx` (premium template + Pandoc when available)
4. Print token usage and estimated cost

## Dependencies (auto-install)

**New users only need:** `python main.py` — no manual `pip install` required.

On startup, `main.py` checks that all libraries are importable:

| Package | Used for |
|---------|----------|
| `anthropic` | Claude API formatting |
| `python-docx` | Word document creation |
| `markdown-it-py` (≥ 4.1) | Markdown parsing |
| `mdit-py-plugins` | Tables / GFM |
| `python-dotenv` | `.env` API key |
| `pypandoc_binary` | High-quality DOCX (bundled Pandoc) |

If anything is missing, it runs:

`python -m pip install -r requirements.txt`

(`requirements.txt` must match `PIP_REQUIREMENTS` in `main.py`.)

Skip auto-install:

```powershell
python main.py --skip-deps
```

## Setup

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

### Claude API key

Create a `.env` file in the project root and paste your key:

```text
ANTHROPIC_API_KEY=YOUR_KEY
```

You can start from `.env.example`.

Use the exact name `ANTHROPIC_API_KEY` (Anthropic’s name, not `CLAUDE_API_KEY`). One line, no spaces around `=`, no quotes unless your shell requires them for `export`. The app loads `.env` from this project folder even if you run `python` from somewhere else.

If you prefer environment variables instead:

```powershell
$env:ANTHROPIC_API_KEY="YOUR_KEY"
```

Optional (override auto model pick):

```powershell
# Leave unset to auto-pick Sonnet → Opus → Haiku (best report structure on Sonnet)
$env:ANTHROPIC_MODEL="your-model-id-from-console.anthropic.com"
```

### Cost estimation (optional)

Anthropic responses include token usage, but not cost. If you want the script to show an estimated USD cost, set:

```powershell
$env:ANTHROPIC_INPUT_PER_MILLION_USD="0.25"
$env:ANTHROPIC_OUTPUT_PER_MILLION_USD="1.25"
```

(Use the rates that match your plan/model.)

## Where is my DOCX? (VS Code / Cursor)

Conversion **already saves** a real file:

- **`output/<name>.docx`** if you use the default `--save output`
- **Next to your `.md` file** if you run `--save same`

The editor may warn that the file **is binary / not shown as text** — that’s normal for Word (`.docx`) files.

How to view it:

1. Open the project’s **`output`** folder or use **File Explorer**, **or**
2. After conversion, Explorer should **select the DOCX automatically** on Windows (**`--no-reveal`** disables this), **or**
3. Open in Word immediately:

```powershell
python main.py --open-docx
```

## Usage

Run with a file picker (no manual path typing):

```powershell
python main.py
```

Convert with Claude formatting (manual path is optional):

```powershell
python main.py "C:\path\to\file.md"
```

Skip AI formatting (plain Markdown → Word; **no** executive summary, callouts, or section rewrite):

```powershell
python main.py "C:\path\to\file.md" --no-ai
```

Use **`python main.py` without `--no-ai`** for rich report-style content.

Choose conversion engine:
- `auto` (default): use Pandoc if installed, else builtin
- `pandoc`: require Pandoc
- `builtin`: pure Python (headings/lists/code blocks/tables best-effort)

```powershell
python main.py "C:\path\to\file.md" --engine auto
```

Choose where the DOCX is saved:
- `--save output` (default): save to local `output/` folder
- `--save same`: save next to the selected `.md` file

```powershell
python main.py --save same
```

## Better DOCX formatting (premium look)

The converter uses a **premium Word template** (`assets/reference.docx`): navy headings, accent underlines, styled tables (blue header row), shaded code blocks, and callout blockquotes.

On first run, the script may install **`pypandoc_binary`** (bundled Pandoc). The terminal should show e.g. `DOCX engine used: pypandoc (premium template + TOC)`.

- **`auto`** (default): bundled Pandoc + template + **table of contents** + numbered sections → post-polish in Word
- Claude rewrites Markdown into an **executive report** (summary, `---` section breaks, callouts, tables) before conversion

By default the app **auto-picks Sonnet** when your API key has it. If `.env` forces an old **Haiku** id, clear `ANTHROPIC_MODEL` or set a **Sonnet** id for the richest text:

```env
ANTHROPIC_MODEL=your-sonnet-model-id
```

## Pandoc (optional system install)

If you install Pandoc separately, `auto` can use it when the bundled copy is unavailable.

Verify it’s available:

```powershell
pandoc --version
```

