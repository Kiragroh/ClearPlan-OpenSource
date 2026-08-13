#!/usr/bin/env python3
"""Build the ClearPlan technical-note DOCX from its UTF-8 Markdown source."""

from __future__ import annotations

import re
from datetime import datetime, timezone
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "paper" / "ClearPlan_ZMP_short_communication.md"
OUTPUT = ROOT / "paper" / "ClearPlan_ZMP_short_communication.docx"
IMAGE_PATTERN = re.compile(r"^!\[(?P<caption>.+?)\]\((?P<path>.+?)\)$")
NUMBERED_PATTERN = re.compile(r"^\d+\.\s+")
INLINE_PATTERN = re.compile(
    r"(\*\*[^*]+\*\*|`[^`]+`|\^[^^]+\^|\[[^\]]+\]\([^)]+\))"
)
TABLE_SEPARATOR_PATTERN = re.compile(
    r"^\|\s*:?-{3,}:?\s*(?:\|\s*:?-{3,}:?\s*)+\|$"
)


def set_cell_margins(section):
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)


def set_font(style, name, size, bold=None, color=None):
    style.font.name = name
    style._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    style.font.size = Pt(size)
    if bold is not None:
        style.font.bold = bold
    if color:
        style.font.color.rgb = RGBColor(*color)


def configure_styles(document):
    styles = document.styles
    set_font(styles["Normal"], "Arial", 10.5)
    styles["Normal"].paragraph_format.space_after = Pt(6)
    styles["Normal"].paragraph_format.line_spacing = 1.08

    set_font(styles["Title"], "Arial", 17, bold=True, color=(23, 37, 84))
    styles["Title"].paragraph_format.space_after = Pt(12)

    for name, size, color in [
        ("Heading 1", 14, (23, 37, 84)),
        ("Heading 2", 12, (29, 78, 216)),
        ("Heading 3", 11, (36, 50, 74)),
    ]:
        set_font(styles[name], "Arial", size, bold=True, color=color)
        styles[name].paragraph_format.space_before = Pt(10)
        styles[name].paragraph_format.space_after = Pt(4)
        styles[name].paragraph_format.keep_with_next = True

    if "Figure Caption" not in styles:
        figure_style = styles.add_style("Figure Caption", WD_STYLE_TYPE.PARAGRAPH)
    else:
        figure_style = styles["Figure Caption"]
    set_font(figure_style, "Arial", 9, color=(55, 65, 81))
    figure_style.paragraph_format.space_after = Pt(8)
    figure_style.paragraph_format.keep_with_next = False

    if "Author Line" not in styles:
        author_style = styles.add_style("Author Line", WD_STYLE_TYPE.PARAGRAPH)
    else:
        author_style = styles["Author Line"]
    set_font(author_style, "Arial", 10.5, bold=True, color=(23, 37, 84))
    author_style.paragraph_format.space_after = Pt(6)

    if "Table Text" not in styles:
        table_style = styles.add_style("Table Text", WD_STYLE_TYPE.PARAGRAPH)
    else:
        table_style = styles["Table Text"]
    set_font(table_style, "Arial", 8.5)
    table_style.paragraph_format.space_after = Pt(0)
    table_style.paragraph_format.line_spacing = 1.0

    set_font(styles["List Bullet"], "Arial", 10.5)


def add_page_number(section):
    footer = section.footer
    paragraph = footer.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = "PAGE"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instruction, end])


def add_inline_runs(paragraph, text):
    position = 0
    for match in INLINE_PATTERN.finditer(text):
        if match.start() > position:
            paragraph.add_run(text[position:match.start()])
        token = match.group(0)
        if token.startswith("**"):
            run = paragraph.add_run(token[2:-2])
            run.bold = True
        elif token.startswith("`"):
            run = paragraph.add_run(token[1:-1])
            run.font.name = "Courier New"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "Courier New")
            run.font.size = Pt(9)
        elif token.startswith("^"):
            run = paragraph.add_run(token[1:-1])
            run.font.superscript = True
        else:
            link_match = re.match(r"\[([^\]]+)\]\(([^)]+)\)", token)
            paragraph.add_run(
                "{} ({})".format(link_match.group(1), link_match.group(2))
            )
        position = match.end()
    if position < len(text):
        paragraph.add_run(text[position:])


def add_paragraph(document, text, style=None):
    paragraph = document.add_paragraph(style=style)
    add_inline_runs(paragraph, text)
    return paragraph


def add_figure(document, image_path):
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(4)
    paragraph.paragraph_format.space_after = Pt(3)
    paragraph.paragraph_format.keep_with_next = True
    run = paragraph.add_run()
    run.add_picture(str(image_path), width=Inches(6.45))


def parse_table_row(line):
    return [cell.strip() for cell in line.strip().strip("|").split("|")]


def set_cell_shading(cell, fill):
    properties = cell._tc.get_or_add_tcPr()
    shading = properties.find(qn("w:shd"))
    if shading is None:
        shading = OxmlElement("w:shd")
        properties.append(shading)
    shading.set(qn("w:fill"), fill)


def prevent_row_split(row):
    properties = row._tr.get_or_add_trPr()
    cannot_split = OxmlElement("w:cantSplit")
    properties.append(cannot_split)


def add_table(document, header, body, alignment_tokens):
    table = document.add_table(rows=1, cols=len(header))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    for index, value in enumerate(header):
        cell = table.rows[0].cells[index]
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_shading(cell, "17324D")
        paragraph = cell.paragraphs[0]
        paragraph.style = document.styles["Table Text"]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(value)
        run.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
    prevent_row_split(table.rows[0])

    for row_index, values in enumerate(body):
        cells = table.add_row().cells
        for column_index, value in enumerate(values):
            cell = cells[column_index]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if row_index % 2:
                set_cell_shading(cell, "F1F5F9")
            paragraph = cell.paragraphs[0]
            paragraph.style = document.styles["Table Text"]
            token = alignment_tokens[column_index]
            if token.endswith(":") and token.startswith(":"):
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif token.endswith(":"):
                paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            add_inline_runs(paragraph, value)
        prevent_row_split(table.rows[-1])

    document.add_paragraph().paragraph_format.space_after = Pt(0)


def render_markdown(document, source_path):
    lines = source_path.read_text(encoding="utf-8").splitlines()
    index = 0
    while index < len(lines):
        raw = lines[index].rstrip()
        stripped = raw.strip()
        if not stripped:
            index += 1
            continue

        image_match = IMAGE_PATTERN.match(stripped)
        if image_match:
            image_path = (source_path.parent / image_match.group("path")).resolve()
            if not image_path.is_file():
                raise FileNotFoundError(f"Figure not found: {image_path}")
            add_figure(document, image_path)
            index += 1
            continue

        if (
            stripped.startswith("|")
            and index + 1 < len(lines)
            and TABLE_SEPARATOR_PATTERN.match(lines[index + 1].strip())
        ):
            header = parse_table_row(stripped)
            alignment_tokens = parse_table_row(lines[index + 1].strip())
            if len(header) != len(alignment_tokens):
                raise RuntimeError("Markdown table header and separator do not match.")
            index += 2
            body = []
            while index < len(lines) and lines[index].strip().startswith("|"):
                values = parse_table_row(lines[index].strip())
                if len(values) != len(header):
                    raise RuntimeError("Markdown table row has the wrong column count.")
                body.append(values)
                index += 1
            add_table(document, header, body, alignment_tokens)
            continue

        if stripped.startswith("# "):
            paragraph = document.add_paragraph(style="Title")
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            add_inline_runs(paragraph, stripped[2:])
            index += 1
            continue
        if stripped.startswith("## "):
            add_paragraph(document, stripped[3:], "Heading 1")
            index += 1
            continue
        if stripped.startswith("### "):
            add_paragraph(document, stripped[4:], "Heading 2")
            index += 1
            continue
        if stripped.startswith("- "):
            add_paragraph(document, stripped[2:], "List Bullet")
            index += 1
            continue
        if stripped.startswith("**Figure "):
            add_paragraph(document, stripped, "Figure Caption")
            index += 1
            continue
        if NUMBERED_PATTERN.match(stripped):
            add_paragraph(document, stripped)
            index += 1
            continue
        if stripped.startswith("Maximilian Grohmann^"):
            add_paragraph(document, stripped, "Author Line")
            index += 1
            continue

        paragraph_lines = [stripped]
        index += 1
        while index < len(lines):
            candidate = lines[index].strip()
            if (not candidate or candidate.startswith("#") or candidate.startswith("- ") or
                    NUMBERED_PATTERN.match(candidate) or
                    IMAGE_PATTERN.match(candidate) or
                    candidate.startswith("|") or
                    candidate.startswith("**Figure ")):
                break
            paragraph_lines.append(candidate)
            index += 1
        add_paragraph(document, " ".join(paragraph_lines))


def validate_source(path):
    text = path.read_text(encoding="utf-8")
    abstract_match = re.search(
        r"## Abstract\s+(.+?)\s+Keywords:\s*(.+?)\s+## 1\.",
        text,
        flags=re.DOTALL,
    )
    if not abstract_match:
        raise RuntimeError("Could not locate the abstract and keywords.")
    abstract_words = re.findall(r"\b[\w’'-]+\b", abstract_match.group(1))
    if len(abstract_words) > 250:
        raise RuntimeError(
            f"ZMP abstract limit exceeded: {len(abstract_words)} words."
        )
    keywords = [
        value.strip()
        for value in abstract_match.group(2).split(";")
        if value.strip()
    ]
    if not 1 <= len(keywords) <= 6:
        raise RuntimeError(f"ZMP requires 1–6 keywords, found {len(keywords)}.")

    forbidden = [
        "To be completed",
        "must be completed",
        "Figure_1_ClearPlan_GUI.png",
        "Figure_2_Constraint_sources.png",
    ]
    present = [value for value in forbidden if value in text]
    if present:
        raise RuntimeError(f"Manuscript contains placeholders or stale content: {present}")

    required = [
        "Maximilian Grohmann",
        "Maria Jäckel",
        "Manuel Todorovic",
        "Cordula Petersen",
        "Andrea Baehr",
        "maximilian.grohmann@medizin.uni-leipzig.de",
        "0000-0002-6909-811X",
        "reports exclusively analytically generated synthetic scenario artifacts",
        "125 C# tests",
        "20 tests",
        "releases/tag/v3.1.0",
    ]
    missing = [value for value in required if value not in text]
    if missing:
        raise RuntimeError(f"Manuscript source is missing required content: {missing}")
    print(
        f"Validated manuscript source: abstract={len(abstract_words)} words, "
        f"keywords={len(keywords)}."
    )


def validate_docx(path):
    document = Document(str(path))
    text_parts = [paragraph.text for paragraph in document.paragraphs]
    for table in document.tables:
        for row in table.rows:
            text_parts.extend(cell.text for cell in row.cells)
    text = "\n".join(text_parts)
    required = [
        "Maximilian Grohmann",
        "Maria Jäckel",
        "University Medical Center Hamburg-Eppendorf",
        "Excel workbook",
        "RefDB JSON",
        "179-181 T30 UZ",
        "125 C# tests",
        "v3.1.0",
        "did not receive any specific grant",
        "declare no competing interests",
    ]
    missing = [value for value in required if value not in text]
    if missing:
        raise RuntimeError(f"DOCX is missing expected Unicode/content: {missing}")
    if "\ufffd" in text:
        raise RuntimeError("DOCX contains the Unicode replacement character.")
    if len(document.inline_shapes) != 3:
        raise RuntimeError(
            f"DOCX must contain three figures, found {len(document.inline_shapes)}."
        )
    if len(document.tables) != 4:
        raise RuntimeError(
            f"DOCX must contain four editable tables, found {len(document.tables)}."
        )
    if len(document.sections) != 1:
        raise RuntimeError("The submission DOCX must remain single-section/single-column.")
    print(
        f"Reopened DOCX: figures={len(document.inline_shapes)}, "
        f"tables={len(document.tables)}, sections={len(document.sections)}."
    )


def main():
    document = Document()
    section = document.sections[0]
    set_cell_margins(section)
    section.start_type = WD_SECTION.CONTINUOUS
    add_page_number(section)
    configure_styles(document)

    properties = document.core_properties
    properties.title = "ClearPlan technical note"
    properties.subject = "Simulator-backed verification of radiotherapy plan-review software"
    properties.author = "Maximilian Grohmann et al."
    properties.keywords = (
        "radiotherapy, plan review, quality assurance, software testing, "
        "treatment planning system, simulation"
    )
    fixed_time = datetime(2026, 7, 30, 15, 0, 0, tzinfo=timezone.utc)
    properties.created = fixed_time
    properties.modified = fixed_time

    validate_source(SOURCE)
    render_markdown(document, SOURCE)
    document.save(str(OUTPUT))
    validate_docx(OUTPUT)
    print(f"Wrote and reopened {OUTPUT}")


if __name__ == "__main__":
    main()
