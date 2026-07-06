from pathlib import Path
from zipfile import ZipFile

from docx import Document

from wc_word_report_tool import WordFormatter


def _read_zip_xml(docx_path: Path, inner_path: str) -> str:
    with ZipFile(docx_path) as zf:
        return zf.read(inner_path).decode("utf-8", errors="ignore")


def test_heading_1_uses_builtin_heading_style(tmp_path: Path):
    doc = Document()
    paragraph = WordFormatter.Heading_1(doc, "一级标题")
    out = tmp_path / "heading.docx"
    doc.save(out)

    assert paragraph.style.name == "Heading 1"
    document_xml = _read_zip_xml(out, "word/document.xml")
    assert 'w:pStyle w:val="Heading1"' in document_xml


def test_restart_page_numbering_unlinks_footer_and_sets_start(tmp_path: Path):
    doc = Document()
    section = WordFormatter.insert_section_with_page_numbering(
        doc,
        start_page_number=1,
        prefix="第",
        suffix="页",
    )
    out = tmp_path / "page.docx"
    doc.save(out)

    assert section.footer.is_linked_to_previous is False
    assert section.header.is_linked_to_previous is False

    document_xml = _read_zip_xml(out, "word/document.xml")
    assert 'w:pgNumType w:start="1"' in document_xml

    footer_xml = _read_zip_xml(out, "word/footer1.xml")
    assert "PAGE" in footer_xml


def test_add_toc_with_custom_style_mapping(tmp_path: Path):
    doc = Document()
    WordFormatter.set_paragraph_style(doc, "MyHeading1", font_name="黑体", font_size=14, bold=True)
    WordFormatter.add_toc(
        doc,
        levels=(1, 3),
        use_outline_levels=True,
        custom_style_levels={"MyHeading1": 1},
        toc_level_styles={1: {"font_name": "黑体", "font_size": 14, "bold": True}},
    )
    WordFormatter.add_custom_heading(doc, "自定义标题", style_name="MyHeading1", level=1)

    out = tmp_path / "toc.docx"
    doc.save(out)

    document_xml = _read_zip_xml(out, "word/document.xml")
    settings_xml = _read_zip_xml(out, "word/settings.xml")
    styles_xml = _read_zip_xml(out, "word/styles.xml")

    assert 'TOC \\o "1-3" \\h \\z \\u \\t "MyHeading1,1"' in document_xml
    assert "updateFields" in settings_xml
    assert 'styleId="TOC1"' in styles_xml or 'styleId="TOC 1"' in styles_xml
