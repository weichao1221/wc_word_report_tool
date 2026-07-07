"""测试 wc_word_report_tool 的新 API 与向后兼容性。"""
from pathlib import Path
from zipfile import ZipFile

import pytest
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

from wc_word_report_tool import WordFormatter
from wc_word_report_tool.word import FONT_SIZE_MAP, _to_pt


def _read_zip_xml(docx_path: Path, inner_path: str) -> str:
    with ZipFile(docx_path) as zf:
        return zf.read(inner_path).decode("utf-8", errors="ignore")


# ====================================================================
# 字号工具
# ====================================================================

def test_to_pt_supports_int_and_float():
    assert _to_pt(14) == 14.0
    assert _to_pt(14.5) == 14.5


def test_to_pt_supports_chinese_font_size():
    assert _to_pt("三号") == 16
    assert _to_pt("小五") == 9
    assert _to_pt("初号") == 42


def test_to_pt_supports_numeric_string():
    assert _to_pt("16") == 16.0


def test_to_pt_raises_on_unknown_string():
    with pytest.raises(ValueError):
        _to_pt("unknown")


def test_font_size_map_contains_common_sizes():
    assert FONT_SIZE_MAP["三号"] == 16
    assert FONT_SIZE_MAP["小五"] == 9
    assert FONT_SIZE_MAP["四号"] == 14


# ====================================================================
# 文档级设置
# ====================================================================

def test_set_default_font_sets_normal_style(tmp_path):
    doc = Document()
    WordFormatter.set_default_font(doc, font_size="三号",
                                    cn_font="仿宋_GB2312", en_font="Times New Roman")
    out = tmp_path / "default_font.docx"
    doc.save(out)

    styles_xml = _read_zip_xml(out, "word/styles.xml")
    assert 'w:eastAsia="仿宋_GB2312"' in styles_xml
    assert 'w:ascii="Times New Roman"' in styles_xml
    assert 'w:sz w:val="32"' in styles_xml  # 16pt = 32 half-points


def test_set_language_sets_zh_cn(tmp_path):
    doc = Document()
    WordFormatter.set_language(doc, "zh-CN")
    out = tmp_path / "lang.docx"
    doc.save(out)

    settings_xml = _read_zip_xml(out, "word/settings.xml")
    assert 'w:val="zh-CN"' in settings_xml
    assert 'w:eastAsia="zh-CN"' in settings_xml


def test_set_page_margins(tmp_path):
    doc = Document()
    WordFormatter.set_page_margins(doc, top=2.5, left=2.6)
    out = tmp_path / "margins.docx"
    doc.save(out)

    document_xml = _read_zip_xml(out, "word/document.xml")
    assert 'w:top="1417"' in document_xml  # 2.5cm
    assert 'w:left="1474"' in document_xml  # 2.6cm


# ====================================================================
# 段落与标题（新 API）
# ====================================================================

def test_body_creates_paragraph_with_indent(tmp_path):
    doc = Document()
    par = WordFormatter.body(doc, "正文内容")
    assert par.alignment == WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    assert par.paragraph_format.first_line_indent is not None
    assert par.paragraph_format.line_spacing == 1.5
    out = tmp_path / "body.docx"
    doc.save(out)
    document_xml = _read_zip_xml(out, "word/document.xml")
    assert 'w:eastAsia="仿宋_GB2312"' in document_xml


def test_body_with_no_indent(tmp_path):
    doc = Document()
    par = WordFormatter.body(doc, "无缩进", indent=False)
    assert par.paragraph_format.first_line_indent is None


def test_heading1_uses_builtin_heading_style(tmp_path):
    doc = Document()
    paragraph = WordFormatter.heading1(doc, "一级标题")
    out = tmp_path / "h1.docx"
    doc.save(out)

    assert paragraph.style.name == "Heading 1"
    document_xml = _read_zip_xml(out, "word/document.xml")
    assert 'w:pStyle w:val="Heading1"' in document_xml


def test_heading2_is_bold(tmp_path):
    doc = Document()
    par = WordFormatter.heading2(doc, "二级")
    out = tmp_path / "h2.docx"
    doc.save(out)
    assert par.runs[0].font.bold is True


def test_heading3_is_not_bold():
    doc = Document()
    par = WordFormatter.heading3(doc, "三级")
    assert par.runs[0].font.bold is False


def test_cover_text_is_centered_and_bold():
    doc = Document()
    par = WordFormatter.cover_text(doc, "项目名", font_size=22, bold=True)
    assert par.alignment == WD_PARAGRAPH_ALIGNMENT.CENTER
    assert par.runs[0].font.bold is True
    assert par.runs[0].font.size.pt == 22


def test_blank_lines_creates_n_paragraphs():
    doc = Document()
    initial = len(doc.paragraphs)
    WordFormatter.blank_lines(doc, 3)
    assert len(doc.paragraphs) == initial + 3


def test_right_text_is_right_aligned():
    doc = Document()
    par = WordFormatter.right_text(doc, "公司名")
    assert par.alignment == WD_PARAGRAPH_ALIGNMENT.RIGHT


def test_created_time_uses_today():
    import datetime
    doc = Document()
    par = WordFormatter.created_time(doc)
    today = datetime.datetime.now().strftime("%Y年%m月%d日")
    assert today in par.text


# ====================================================================
# 表格（新增能力）
# ====================================================================

def test_set_table_borders(tmp_path):
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    WordFormatter.set_table_borders(table, color="000000", size=4)
    out = tmp_path / "borders.docx"
    doc.save(out)

    document_xml = _read_zip_xml(out, "word/document.xml")
    assert "tblBorders" in document_xml
    assert 'w:color="000000"' in document_xml


def test_set_cell_supports_bold_and_alignment(tmp_path):
    doc = Document()
    table = doc.add_table(rows=1, cols=1)
    WordFormatter.set_cell(
        table.rows[0].cells[0], "表头",
        bold=True, alignment="居中",
    )
    out = tmp_path / "cell.docx"
    doc.save(out)

    cell = table.rows[0].cells[0]
    assert cell.paragraphs[0].alignment == WD_PARAGRAPH_ALIGNMENT.CENTER
    assert cell.paragraphs[0].runs[0].font.bold is True


def test_add_table_creates_table_with_headers(tmp_path):
    doc = Document()
    headers = ["序号", "项目", "金额"]
    rows = [["1", "建安费", "1200"], ["2", "设备费", "850"]]
    table = WordFormatter.add_table(doc, headers, rows, col_widths=[2, 5, 3])

    out = tmp_path / "table.docx"
    doc.save(out)

    assert len(table.rows) == 3  # 1 header + 2 data
    assert len(table.columns) == 3

    # 表头应该加粗
    header_cell = table.rows[0].cells[0]
    assert header_cell.paragraphs[0].runs[0].font.bold is True

    # 数据不应加粗
    data_cell = table.rows[1].cells[0]
    assert data_cell.paragraphs[0].runs[0].font.bold is False

    # 边框已添加
    document_xml = _read_zip_xml(out, "word/document.xml")
    assert "tblBorders" in document_xml


def test_set_cell_clears_existing_runs():
    """重复调用 set_cell 不应叠加 runs。"""
    doc = Document()
    table = doc.add_table(rows=1, cols=1)
    cell = table.rows[0].cells[0]
    WordFormatter.set_cell(cell, "第一次")
    WordFormatter.set_cell(cell, "第二次")
    assert len(cell.paragraphs[0].runs) == 1
    assert cell.paragraphs[0].text == "第二次"


# ====================================================================
# 页眉页脚（新增能力）
# ====================================================================

def test_set_header(tmp_path):
    doc = Document()
    doc.add_paragraph("正文")
    from docx.enum.section import WD_SECTION_START
    doc.add_section(WD_SECTION_START.NEW_PAGE)
    doc.add_paragraph("第二页")

    sec = doc.sections[1]
    WordFormatter.set_header(sec, "项目名称 报告名称", alignment="居中")
    out = tmp_path / "header.docx"
    doc.save(out)

    # header xml 文件名可能为 header1.xml / header2.xml 等，遍历查找
    with ZipFile(out) as zf:
        header_names = [n for n in zf.namelist() if n.startswith("word/header")]
        header_xml = ""
        for name in header_names:
            content = zf.read(name).decode("utf-8", errors="ignore")
            if "项目名称" in content:
                header_xml = content
                break
    assert "项目名称 报告名称" in header_xml
    assert "w:jc" in header_xml  # 居中


def test_set_footer_uses_kaiti_font(tmp_path):
    doc = Document()
    doc.add_paragraph("正文")
    from docx.enum.section import WD_SECTION_START
    doc.add_section(WD_SECTION_START.NEW_PAGE)
    doc.add_paragraph("第二页")

    sec = doc.sections[1]
    WordFormatter.set_footer(sec, "公司名称", alignment="居右")
    out = tmp_path / "footer.docx"
    doc.save(out)

    # footer xml 文件名可能为 footer1.xml / footer2.xml 等
    with ZipFile(out) as zf:
        footer_names = [n for n in zf.namelist() if n.startswith("word/footer")]
        footer_xml = ""
        for name in footer_names:
            content = zf.read(name).decode("utf-8", errors="ignore")
            if "公司名称" in content:
                footer_xml = content
                break
    assert "公司名称" in footer_xml
    assert 'w:eastAsia="楷体"' in footer_xml
    assert 'w:sz w:val="18"' in footer_xml  # 9pt = 18 half-points


def test_set_header_does_not_stack_runs():
    """重复调用不应叠加。"""
    doc = Document()
    sec = doc.sections[0]
    WordFormatter.set_header(sec, "第一次")
    WordFormatter.set_header(sec, "第二次")
    assert len(sec.header.paragraphs[0].runs) == 1


def test_clear_footer_removes_content():
    doc = Document()
    sec = doc.sections[0]
    WordFormatter.set_footer(sec, "原有页脚")
    WordFormatter.clear_footer(sec)
    assert all(not p.runs for p in sec.footer.paragraphs)


def test_clear_header_removes_content():
    doc = Document()
    sec = doc.sections[0]
    WordFormatter.set_header(sec, "原有页眉")
    WordFormatter.clear_header(sec)
    assert all(not p.runs for p in sec.header.paragraphs)


# ====================================================================
# 跨节页码（新增能力）
# ====================================================================

def test_set_page_number_from_section_clears_previous_sections(tmp_path):
    """前置节无页码，正文节重启页码，后续节链接。"""
    doc = Document()
    doc.add_paragraph("前置内容")

    from docx.enum.section import WD_SECTION_START
    doc.add_section(WD_SECTION_START.NEW_PAGE)
    doc.add_paragraph("正文第一页")

    doc.add_section(WD_SECTION_START.NEW_PAGE)
    doc.add_paragraph("正文第二页")

    WordFormatter.set_page_number_from_section(
        doc, start_section_idx=1, start=1, prefix="第 ", suffix=" 页",
    )
    out = tmp_path / "page_number_from_section.docx"
    doc.save(out)

    # 节 0 页脚应为空（无 runs）
    sec0_footer_runs = sum(
        len(p.runs) for p in doc.sections[0].footer.paragraphs
    )
    assert sec0_footer_runs == 0

    # 节 1 页脚应有 PAGE 域
    with ZipFile(out) as zf:
        footer_names = [n for n in zf.namelist() if n.startswith("word/footer")]
        footer_xml = ""
        for name in footer_names:
            content = zf.read(name).decode("utf-8", errors="ignore")
            if "PAGE" in content:
                footer_xml = content
                break
    assert "PAGE" in footer_xml
    assert "第 " in footer_xml

    # 节 2 应链接到前一节
    assert doc.sections[2].footer.is_linked_to_previous is True


def test_set_page_number_from_section_invalid_index_raises():
    doc = Document()
    with pytest.raises(ValueError, match="超出范围"):
        WordFormatter.set_page_number_from_section(doc, start_section_idx=99)


# ====================================================================
# 严格对齐校验
# ====================================================================

def test_resolve_alignment_strict_raises_on_unknown():
    with pytest.raises(ValueError, match="未知对齐方式"):
        WordFormatter.resolve_alignment("invalid_align", strict=True)


def test_resolve_alignment_non_strict_falls_back_to_left():
    result = WordFormatter.resolve_alignment("invalid_align", strict=False)
    assert result == WD_PARAGRAPH_ALIGNMENT.LEFT


def test_resolve_alignment_supports_chinese():
    assert WordFormatter.resolve_alignment("居中") == WD_PARAGRAPH_ALIGNMENT.CENTER
    assert WordFormatter.resolve_alignment("居左") == WD_PARAGRAPH_ALIGNMENT.LEFT
    assert WordFormatter.resolve_alignment("居右") == WD_PARAGRAPH_ALIGNMENT.RIGHT
    assert WordFormatter.resolve_alignment("左对齐") == WD_PARAGRAPH_ALIGNMENT.LEFT
    assert WordFormatter.resolve_alignment("右对齐") == WD_PARAGRAPH_ALIGNMENT.RIGHT
    assert WordFormatter.resolve_alignment("两端对齐") == WD_PARAGRAPH_ALIGNMENT.JUSTIFY


def test_resolve_alignment_supports_english():
    assert WordFormatter.resolve_alignment("center") == WD_PARAGRAPH_ALIGNMENT.CENTER
    assert WordFormatter.resolve_alignment("LEFT") == WD_PARAGRAPH_ALIGNMENT.LEFT


def test_resolve_alignment_supports_enum():
    enum_val = WD_PARAGRAPH_ALIGNMENT.CENTER
    assert WordFormatter.resolve_alignment(enum_val) == enum_val


# ====================================================================
# 向后兼容（旧 API 仍可用，但发 DeprecationWarning）
# ====================================================================

def test_old_api_heading_1_still_works(tmp_path, recwarn):
    doc = Document()
    paragraph = WordFormatter.Heading_1(doc, "一级标题")
    out = tmp_path / "old_api.docx"
    doc.save(out)

    assert paragraph.style.name == "Heading 1"
    # 应该有 DeprecationWarning
    assert any(issubclass(w.category, DeprecationWarning) for w in recwarn)


def test_old_api_normal_doc_still_works(tmp_path, recwarn):
    doc = Document()
    par = WordFormatter.Normal_doc(doc, "正文")
    out = tmp_path / "old_normal.docx"
    doc.save(out)

    assert par.paragraph_format.line_spacing == 1.5
    assert any(issubclass(w.category, DeprecationWarning) for w in recwarn)


def test_old_api_fengmian_doc1_still_works(recwarn):
    doc = Document()
    par = WordFormatter.fengmian_doc1(doc, "项目名")
    assert par.alignment == WD_PARAGRAPH_ALIGNMENT.CENTER
    assert par.runs[0].font.bold is True
    assert any(issubclass(w.category, DeprecationWarning) for w in recwarn)


def test_old_api_set_all_layout_still_works(tmp_path, recwarn):
    doc = Document()
    WordFormatter.set_all_layout(doc, top=2.5)
    out = tmp_path / "old_layout.docx"
    doc.save(out)

    document_xml = _read_zip_xml(out, "word/document.xml")
    assert 'w:top="1417"' in document_xml  # 2.5cm
    assert any(issubclass(w.category, DeprecationWarning) for w in recwarn)


def test_old_api_set_cell_format_falls_back_to_set_cell(recwarn):
    doc = Document()
    table = doc.add_table(rows=1, cols=1)
    WordFormatter.set_cell_format(table.rows[0].cells[0], "内容")
    assert table.rows[0].cells[0].paragraphs[0].text == "内容"
    assert any(issubclass(w.category, DeprecationWarning) for w in recwarn)


# ====================================================================
# 原有测试（保持兼容）
# ====================================================================

def test_restart_page_numbering_unlinks_footer_and_sets_start(tmp_path):
    doc = Document()
    section = WordFormatter.insert_section_with_page_numbering(
        doc, start_page_number=1, prefix="第", suffix="页",
    )
    out = tmp_path / "page.docx"
    doc.save(out)

    assert section.footer.is_linked_to_previous is False
    assert section.header.is_linked_to_previous is False

    document_xml = _read_zip_xml(out, "word/document.xml")
    assert 'w:pgNumType w:start="1"' in document_xml

    footer_xml = _read_zip_xml(out, "word/footer1.xml")
    assert "PAGE" in footer_xml


def test_add_toc_with_custom_style_mapping(tmp_path):
    doc = Document()
    WordFormatter.set_paragraph_style(doc, "MyHeading1", font_name="黑体",
                                       font_size=14, bold=True)
    WordFormatter.add_toc(
        doc, levels=(1, 3), use_outline_levels=True,
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


def test_normal_doc_supports_custom_font_name_and_size(tmp_path):
    doc = Document()
    WordFormatter.Normal_doc(doc, "正文", font_name="黑体", font_size=18)
    out = tmp_path / "font.docx"
    doc.save(out)

    document_xml = _read_zip_xml(out, "word/document.xml")
    assert 'w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="黑体"' in document_xml
    assert 'w:sz w:val="36"' in document_xml


def test_set_document_layout_supports_numeric_alignment(tmp_path):
    doc = Document()
    WordFormatter.set_document_layout(doc, left=2.5, right=2.6, horizontal_alignment=2)
    out = tmp_path / "layout.docx"
    doc.save(out)

    document_xml = _read_zip_xml(out, "word/document.xml")
    assert 'w:left="1417"' in document_xml
    assert 'w:right="1474"' in document_xml
    assert 'w:jc w:val="center"' in document_xml
