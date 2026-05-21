"""
Premium Word styling — shared by Pandoc reference template and built-in converter.
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

# Palette (modern report / SaaS document)
NAVY = RGBColor(0x0F, 0x17, 0x2A)
ACCENT = RGBColor(0x25, 0x63, 0xEB)
ACCENT_DARK = RGBColor(0x1E, 0x40, 0xAF)
MUTED = RGBColor(0x64, 0x74, 0x8B)
BODY = RGBColor(0x33, 0x41, 0x55)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LINK = RGBColor(0x25, 0x63, 0xEB)

FILL_HEADER = "1E40AF"
FILL_ROW_ALT = "F8FAFC"
FILL_CODE = "F1F5F9"
FILL_QUOTE = "EFF6FF"
FILL_ACCENT_BAR = "2563EB"


def reference_docx_path() -> Path:
    return Path(__file__).resolve().parent.parent / "assets" / "reference.docx"


def _set_paragraph_shading(paragraph, fill_hex: str) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    p_pr.append(shd)


def _set_cell_shading(cell, fill_hex: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tc_pr.append(shd)


def _set_paragraph_bottom_border(paragraph, color_hex: str = "2563EB", size: int = 12) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), color_hex)
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def _set_paragraph_left_border(paragraph, color_hex: str = "2563EB", size: int = 24) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), str(size))
    left.set(qn("w:space"), "8")
    left.set(qn("w:color"), color_hex)
    p_bdr.append(left)
    p_pr.append(p_bdr)


def apply_premium_theme(doc: Document) -> None:
    """Page layout + style definitions for built-in conversion."""
    for section in doc.sections:
        section.top_margin = Inches(0.9)
        section.bottom_margin = Inches(0.85)
        section.left_margin = Inches(1.1)
        section.right_margin = Inches(1.1)

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    normal.font.color.rgb = BODY
    pf = normal.paragraph_format
    pf.space_after = Pt(8)
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = 1.22

    specs = (
        (1, 26, NAVY, Pt(18), Pt(10)),
        (2, 16, ACCENT_DARK, Pt(14), Pt(8)),
        (3, 13, MUTED, Pt(10), Pt(6)),
    )
    for level, size, color, before, after in specs:
        h = doc.styles[f"Heading {level}"]
        h.font.name = "Calibri Light"
        h.font.bold = True
        h.font.size = Pt(size)
        h.font.color.rgb = color
        h.paragraph_format.space_before = before
        h.paragraph_format.space_after = after

    for style_name in ("List Bullet", "List Number"):
        if style_name in doc.styles:
            ls = doc.styles[style_name]
            ls.font.name = "Calibri"
            ls.font.size = Pt(11)
            ls.font.color.rgb = BODY
            ls.paragraph_format.space_after = Pt(4)


def style_heading_paragraph(paragraph, level: int, *, is_first_h1: bool = False) -> None:
    """Extra polish on heading paragraphs after content is added."""
    if level == 1:
        if is_first_h1:
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        _set_paragraph_bottom_border(paragraph, color_hex="2563EB", size=16)
    elif level == 2:
        _set_paragraph_bottom_border(paragraph, color_hex="CBD5E1", size=6)


def style_code_block(paragraph) -> None:
    paragraph.paragraph_format.left_indent = Inches(0.2)
    paragraph.paragraph_format.space_before = Pt(10)
    paragraph.paragraph_format.space_after = Pt(10)
    _set_paragraph_shading(paragraph, FILL_CODE)
    _set_paragraph_left_border(paragraph, FILL_ACCENT_BAR, size=28)


def style_blockquote(paragraph) -> None:
    paragraph.paragraph_format.left_indent = Inches(0.35)
    paragraph.paragraph_format.space_before = Pt(8)
    paragraph.paragraph_format.space_after = Pt(8)
    _set_paragraph_shading(paragraph, FILL_QUOTE)
    _set_paragraph_left_border(paragraph, FILL_ACCENT_BAR, size=32)


def style_table(table) -> None:
    table.style = "Table Grid"
    for r_idx, row in enumerate(table.rows):
        for cell in row.cells:
            if r_idx == 0:
                _set_cell_shading(cell, FILL_HEADER)
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.bold = True
                        run.font.color.rgb = WHITE
                        run.font.name = "Calibri"
                        run.font.size = Pt(10.5)
            elif r_idx % 2 == 0:
                _set_cell_shading(cell, FILL_ROW_ALT)
            for p in cell.paragraphs:
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.space_before = Pt(2)


def style_horizontal_rule(paragraph) -> None:
    paragraph.paragraph_format.space_before = Pt(14)
    paragraph.paragraph_format.space_after = Pt(14)
    _set_paragraph_bottom_border(paragraph, color_hex="94A3B8", size=8)


def build_reference_docx(path: Path | None = None) -> Path:
    """
    Pandoc reference document with matching typography (Heading 1–3, Normal, etc.).
    """
    out = path or reference_docx_path()
    out.parent.mkdir(parents=True, exist_ok=True)

    doc = Document()
    apply_premium_theme(doc)

    # Pandoc maps markdown elements to these style names.
    if "Title" in doc.styles:
        t = doc.styles["Title"]
        t.font.name = "Calibri Light"
        t.font.size = Pt(28)
        t.font.bold = True
        t.font.color.rgb = NAVY

    if "First Paragraph" in doc.styles:
        fp = doc.styles["First Paragraph"]
        fp.font.name = "Calibri"
        fp.font.size = Pt(11)
        fp.font.color.rgb = BODY

    if "Block Text" in doc.styles:
        bt = doc.styles["Block Text"]
        bt.font.name = "Calibri"
        bt.font.italic = True
        bt.font.color.rgb = MUTED

    if "Source Code" in doc.styles:
        sc = doc.styles["Source Code"]
        sc.font.name = "Consolas"
        sc.font.size = Pt(9.5)

    # Seed sample content so Word retains style definitions reliably.
    h1 = doc.add_heading("Document Title", level=1)
    style_heading_paragraph(h1, 1, is_first_h1=True)
    doc.add_heading("Section", level=2)
    doc.add_paragraph("Body text sample.")
    doc.save(str(out))
    return out


def ensure_reference_docx() -> Path:
    path = reference_docx_path()
    if not path.is_file():
        build_reference_docx(path)
    return path


def polish_saved_docx(output_path: str) -> None:
    """
    Apply premium borders, table styling, and margins to an existing DOCX
    (Pandoc output often uses plain Word defaults without post-processing).
    """
    path = Path(output_path)
    if not path.is_file():
        return

    doc = Document(str(path))
    apply_premium_theme(doc)

    first_h1 = True
    for paragraph in doc.paragraphs:
        style_name = (paragraph.style.name if paragraph.style else "") or ""
        lower = style_name.lower()

        if style_name.startswith("Heading"):
            try:
                level = int(style_name.replace("Heading ", "").strip() or "1")
            except ValueError:
                level = 1
            is_first = level == 1 and first_h1
            if level == 1:
                first_h1 = False
            style_heading_paragraph(paragraph, level, is_first_h1=is_first)
        elif "quote" in lower or "block text" in lower:
            style_blockquote(paragraph)
        elif "source code" in lower or "verbatim" in lower or "code" in lower:
            style_code_block(paragraph)
            for run in paragraph.runs:
                run.font.name = "Consolas"
                run.font.size = Pt(9.5)

    for table in doc.tables:
        style_table(table)

    doc.save(str(path))
