"""Direct OOXML package integration tests."""

import io
import struct
import zlib
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from lxml import etree

from wc_word_report_tool import WordFormatter
from wc_word_report_tool.ooxml import (
    WD_CELL_VERTICAL_ALIGNMENT,
    WD_PARAGRAPH_ALIGNMENT,
    WD_SECTION_START,
    WD_TABLE_DIRECTION,
    Cm,
    Document,
    OxmlElement,
    Pt,
    RGBColor,
    qn,
)
from wc_word_report_tool.ooxml.xml import insert_in_order

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _part_xml(path, name):
    with ZipFile(path) as archive:
        return archive.read(name).decode("utf-8")


def _document_xml(path):
    return _part_xml(path, "word/document.xml")


def _child_names(element):
    return [etree.QName(child).localname for child in element]


def _png_bytes() -> bytes:
    """Smallest valid 1x1 PNG so picture tests need no binary fixture."""

    def chunk(tag: bytes, data: bytes) -> bytes:
        body = tag + data
        return (
            struct.pack(">I", len(data))
            + body
            + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)
        )

    header = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", header)
        + chunk(b"IDAT", zlib.compress(b"\x00\xff\x00\x00"))
        + chunk(b"IEND", b"")
    )


def test_word_formatter_can_create_and_open_documents(tmp_path):
    formatter = WordFormatter()
    formatter.body("空白文档直接创建")
    path = tmp_path / "created.docx"
    formatter.doc.save(path)

    reopened = WordFormatter(path)
    assert reopened.doc.paragraphs[0].text == "空白文档直接创建"


def test_unknown_parts_and_namespaces_survive_a_round_trip(tmp_path):
    base = tmp_path / "base.docx"
    Document().save(base)
    source = tmp_path / "source.docx"
    custom_blob = b"<custom xmlns='urn:example'>keep-me</custom>"

    with ZipFile(base) as archive:
        files = {name: archive.read(name) for name in archive.namelist()}

    content_types = files["[Content_Types].xml"].decode("utf-8")
    content_types = content_types.replace(
        "</Types>",
        '<Override PartName="/customXml/item1.xml" '
        'ContentType="application/xml"/></Types>',
    )
    document_xml = files["word/document.xml"].decode("utf-8").replace(
        "<w:document ",
        '<w:document xmlns:x="urn:example:marker" x:marker="keep" ',
        1,
    )
    files["[Content_Types].xml"] = content_types.encode("utf-8")
    files["word/document.xml"] = document_xml.encode("utf-8")
    files["customXml/item1.xml"] = custom_blob

    with ZipFile(source, "w", ZIP_DEFLATED) as archive:
        for name, data in files.items():
            archive.writestr(name, data)

    document = Document(source)
    document.add_paragraph("新增正文")
    output = tmp_path / "output.docx"
    document.save(output)

    with ZipFile(output) as archive:
        assert archive.read("customXml/item1.xml") == custom_blob
        saved_xml = archive.read("word/document.xml").decode("utf-8")
    assert 'x:marker="keep"' in saved_xml


def test_core_properties_are_created_with_a_valid_part_name(tmp_path):
    base = tmp_path / "base.docx"
    Document().save(base)
    source = tmp_path / "source.docx"

    with ZipFile(base) as archive:
        files = {
            name: archive.read(name)
            for name in archive.namelist()
            if name != "docProps/core.xml"
        }
    with ZipFile(source, "w", ZIP_DEFLATED) as archive:
        for name, data in files.items():
            archive.writestr(name, data)

    document = Document(source)
    document.core_properties.title = "直接 OOXML"
    output = tmp_path / "output.docx"
    document.save(output)

    with ZipFile(output) as archive:
        names = archive.namelist()
        assert "docProps/core.xml" in names
        assert "docProps/core1.xml" not in names


def test_consumer_compatibility_surface(tmp_path):
    document = Document()
    section = document.sections[0]
    section.header_distance = Cm(1.2)
    section.footer_distance = Cm(2.2)
    document.add_page_break()

    table = document.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.autofit = False
    row = table.add_row()
    cell = row.cells[0]
    cell.text = "单元格"
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    run = cell.paragraphs[0].runs[0]
    run.bold = True

    assert round(section.header_distance.cm, 2) == 1.2
    assert round(section.footer_distance.cm, 2) == 2.2
    assert table.style.name == "Table Grid"
    assert table.autofit is False
    assert len(table.rows) == 2
    assert cell.text == "单元格"
    assert cell.vertical_alignment == WD_CELL_VERTICAL_ALIGNMENT.CENTER
    assert run.bold is True
    assert run._element is run.element

    output = tmp_path / "consumer-surface.docx"
    document.save(output)
    reopened = Document(output)
    assert len(reopened.tables[0].rows) == 2


def test_run_and_font_boolean_properties_are_written(tmp_path):
    """A 类静默失效回归：italic / all_caps 必须真的写进 XML。"""
    document = Document()
    run = document.add_paragraph("斜体").runs[0]

    run.italic = True
    run.font.all_caps = True

    assert run.italic is True
    assert run.font.italic is True
    assert run.font.all_caps is True

    path = tmp_path / "run-bools.docx"
    document.save(path)
    root = etree.fromstring(_document_xml(path).encode("utf-8"))
    assert root.find(".//%si" % _W).get(_W + "val") is None
    assert root.find(".//%siCs" % _W).get(_W + "val") is None
    assert root.find(".//%scaps" % _W).get(_W + "val") is None

    reopened = Document(path).paragraphs[0].runs[0]
    assert reopened.italic is True
    assert reopened.font.all_caps is True

    # False 写成 w:val="0"（可覆盖继承的开启），None 表示删除该属性节点
    reopened.italic = False
    assert reopened.italic is False
    reopened.italic = None
    assert reopened.italic is False
    r_pr = reopened._r.find(qn("w:rPr"))
    assert r_pr.find(qn("w:i")) is None
    assert r_pr.find(qn("w:iCs")) is None


def test_paragraph_format_keep_together_and_widow_control(tmp_path):
    """A 类静默失效回归：keep_together / widow_control 必须真的写进 XML。"""
    document = Document()
    fmt = document.add_paragraph("段落").paragraph_format

    fmt.keep_together = True
    fmt.widow_control = True
    assert fmt.keep_together is True
    assert fmt.widow_control is True

    on_path = tmp_path / "pf-on.docx"
    document.save(on_path)
    root = etree.fromstring(_document_xml(on_path).encode("utf-8"))
    assert root.find(".//%skeepLines" % _W) is not None
    assert root.find(".//%swidowControl" % _W) is not None

    reopened = Document(on_path).paragraphs[0].paragraph_format
    assert reopened.keep_together is True
    assert reopened.widow_control is True

    # False 写成 w:val="0"（而不是删元素），才能覆盖样式继承的「开」
    reopened.keep_together = False
    reopened.widow_control = False
    assert reopened.keep_together is False
    assert reopened.widow_control is False
    reopened._paragraph._document.save(tmp_path / "pf-off.docx")
    off_root = etree.fromstring(_document_xml(tmp_path / "pf-off.docx").encode("utf-8"))
    assert off_root.find(".//%skeepLines" % _W).get(_W + "val") == "0"
    assert off_root.find(".//%swidowControl" % _W).get(_W + "val") == "0"

    # 只有 None 才删除属性节点
    reopened.keep_together = None
    reopened.widow_control = None
    assert reopened.keep_together is False
    assert reopened.widow_control is False
    reopened._paragraph._document.save(tmp_path / "pf-none.docx")
    none_root = etree.fromstring(_document_xml(tmp_path / "pf-none.docx").encode("utf-8"))
    assert none_root.find(".//%skeepLines" % _W) is None
    assert none_root.find(".//%swidowControl" % _W) is None


def test_false_overrides_style_inherited_paragraph_flags(tmp_path):
    """F1 回归：段落上显式 False 必须能覆盖样式里继承的 True。"""
    document = Document()
    document.styles["Normal"].paragraph_format.keep_together = True
    paragraph = document.add_paragraph("继承覆盖")
    paragraph.paragraph_format.keep_together = False

    path = tmp_path / "override.docx"
    document.save(path)

    style_root = etree.fromstring(_part_xml(path, "word/styles.xml").encode("utf-8"))
    normal = next(
        style
        for style in style_root.iter("%sstyle" % _W)
        if style.get(_W + "styleId") == "Normal"
    )
    style_flag = normal.find("%spPr/%skeepLines" % (_W, _W))
    assert style_flag is not None
    assert style_flag.get(_W + "val") is None

    doc_root = etree.fromstring(_document_xml(path).encode("utf-8"))
    override = doc_root.find(".//%skeepLines" % _W)
    assert override is not None
    assert override.get(_W + "val") == "0"

    reopened = Document(path)
    reopened_paragraph = reopened.paragraphs[0]
    assert reopened_paragraph.text == "继承覆盖"
    assert reopened_paragraph.paragraph_format.keep_together is False


def test_line_spacing_rule_normalises_and_rejects_unknown_values(tmp_path):
    """A 类静默失效回归：line_spacing_rule 现在可读可写，且拒绝非法值。"""
    document = Document()
    fmt = document.add_paragraph("行距").paragraph_format

    assert fmt.line_spacing_rule is None

    fmt.line_spacing_rule = "exact"
    assert fmt.line_spacing_rule == "exact"

    # line_spacing 按值类型决定 lineRule，仍与显式规则一致
    fmt.line_spacing = Pt(20)
    assert fmt.line_spacing_rule == "exact"

    fmt.line_spacing_rule = "atleast"
    assert fmt.line_spacing_rule == "atLeast"

    with pytest.raises(ValueError):
        fmt.line_spacing_rule = "固定值"

    path = tmp_path / "rule.docx"
    document.save(path)
    root = etree.fromstring(_document_xml(path).encode("utf-8"))
    spacing = root.find(".//%sspacing" % _W)
    assert spacing.get(_W + "lineRule") == "atLeast"

    fmt.line_spacing_rule = None
    assert fmt.line_spacing_rule is None


def test_new_properties_keep_schema_child_order(tmp_path):
    """新属性必须按 OOXML 架构顺序插入，不能破坏 pPr / rPr 序列。"""
    document = Document()
    paragraph = document.add_paragraph("顺序")
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    fmt = paragraph.paragraph_format
    fmt.line_spacing = 1.5
    fmt.keep_together = True
    fmt.widow_control = True
    fmt.keep_with_next = True

    run = paragraph.runs[0]
    run.font.underline = True
    run.font.size = Pt(12)
    run.italic = True
    run.font.all_caps = True
    run.font.color.rgb = RGBColor(0x11, 0x22, 0x33)
    run.bold = True

    p_pr = _child_names(paragraph._p.find(qn("w:pPr")))
    assert p_pr.index("keepNext") < p_pr.index("keepLines")
    assert p_pr.index("keepLines") < p_pr.index("widowControl")
    assert p_pr.index("widowControl") < p_pr.index("spacing")
    assert p_pr.index("spacing") < p_pr.index("jc")

    r_pr = _child_names(run._r.find(qn("w:rPr")))
    assert r_pr.index("b") < r_pr.index("bCs") < r_pr.index("i") < r_pr.index("iCs")
    assert r_pr.index("iCs") < r_pr.index("caps") < r_pr.index("color")
    assert r_pr.index("color") < r_pr.index("sz") < r_pr.index("u")

    path = tmp_path / "order.docx"
    document.save(path)
    assert "i" in _document_xml(path)


def test_save_accepts_a_binary_stream(tmp_path):
    """C 类回归：save 同时接受路径和可写二进制流。"""
    document = Document()
    document.add_paragraph("流式保存")

    stream = io.BytesIO()
    document.save(stream)
    data = stream.getvalue()
    assert ZipFile(io.BytesIO(data)).testzip() is None

    path = tmp_path / "stream.docx"
    document.save(path)
    assert [p.text for p in Document(data).paragraphs] == ["流式保存"]
    assert [p.text for p in Document(path).paragraphs] == ["流式保存"]


def test_run_underline_setter_and_add_tab(tmp_path):
    """B 类回归：run.underline 可赋值，run.add_tab() 写入 w:tab。"""
    document = Document()
    run = document.add_paragraph("x").runs[0]

    run.underline = True
    assert run.underline is True
    run.add_tab()
    assert _child_names(run._r)[-1] == "tab"

    path = tmp_path / "run.docx"
    document.save(path)
    root = etree.fromstring(_document_xml(path).encode("utf-8"))
    assert root.find(".//%su" % _W).get(_W + "val") == "single"

    reopened = Document(path).paragraphs[0].runs[0]
    assert reopened.underline is True
    reopened.underline = False
    assert reopened.underline is False


def test_insert_paragraph_before_places_paragraph_in_order():
    """B 类回归：insert_paragraph_before 插在本段之前并返回新段落。"""
    document = Document()
    document.add_paragraph("第一段")
    document.add_paragraph("第三段")

    middle = document.paragraphs[1].insert_paragraph_before("第二段")
    assert [p.text for p in document.paragraphs] == ["第一段", "第二段", "第三段"]
    assert middle.text == "第二段"

    styled = document.paragraphs[0].insert_paragraph_before("标题", style="Heading 1")
    assert styled.style.name == "Heading 1"
    assert [p.text for p in document.paragraphs][0] == "标题"


def test_add_picture_and_inline_shapes(tmp_path):
    """B 类回归：add_picture 与 inline_shapes 可用。"""
    image = tmp_path / "tiny.png"
    image.write_bytes(_png_bytes())

    document = Document()
    document.add_paragraph("带图")
    shape = document.add_picture(str(image), width=Cm(2))

    assert len(document.inline_shapes) == 1
    assert round(document.inline_shapes[0].width.cm, 2) == 2.0
    assert shape.height is not None

    # 表格单元格里的图片同样计入（与 python-docx 的 inline_shapes 一致）
    cell = document.add_table(rows=1, cols=1).cell(0, 0)
    cell.paragraphs[0].add_run().add_picture(str(image))
    assert len(document.inline_shapes) == 2

    path = tmp_path / "picture.docx"
    document.save(path)
    reopened = Document(path)
    assert len(reopened.inline_shapes) == 2
    assert round(reopened.inline_shapes[0].width.cm, 2) == 2.0


def test_insert_in_order_appends_foreign_namespace_children():
    """F2 回归：同名但非 w: 命名空间的元素不做 schema 排序。"""
    parent = OxmlElement("w:pPr")
    parent.append(OxmlElement("w:jc"))
    parent.append(OxmlElement("w:rPr"))

    foreign = etree.Element("{http://example.com/myns}jc")
    insert_in_order(parent, foreign)

    assert _child_names(parent) == ["jc", "rPr", "jc"]
    assert parent[-1] is foreign

    # 正常 w: 子元素仍然排进 schema 区间（keepNext 在 jc/rPr 之前）
    insert_in_order(parent, OxmlElement("w:keepNext"))
    assert _child_names(parent) == ["keepNext", "jc", "rPr", "jc"]


def test_new_public_names_are_exported():
    """O1：批注与图形类应从 wc_word_report_tool.ooxml 直接可用。"""
    from wc_word_report_tool import ooxml

    for name in ("Comment", "Comments", "InlineShape", "InlineShapes"):
        assert name in ooxml.__all__
        assert hasattr(ooxml, name)


def test_comments_round_trip(tmp_path):
    """B 类回归：Document.comments / add_comment 可用且能往返。"""
    document = Document()
    run = document.add_paragraph("需要批注的正文").runs[0]

    comment = document.add_comment(run, text="请补充数据", author="张三", initials="ZS")
    assert len(document.comments) == 1
    assert document.comments[0].text == "请补充数据"
    assert document.comments.get(comment.comment_id).author == "张三"
    assert comment.comment_id is not None

    path = tmp_path / "comments.docx"
    document.save(path)
    with ZipFile(path) as archive:
        assert "word/comments.xml" in archive.namelist()
        document_xml = archive.read("word/document.xml").decode("utf-8")

    for marker in ("commentRangeStart", "commentRangeEnd", "commentReference"):
        assert marker in document_xml

    assert comment.timestamp is not None
    assert comment.timestamp.tzinfo is not None

    reopened = Document(path)
    reopened_comment = reopened.comments[0]
    assert reopened_comment.text == "请补充数据"
    assert reopened_comment.author == "张三"
    assert reopened_comment.initials == "ZS"
    assert reopened_comment.comment_id == comment.comment_id
    assert len(reopened.comments) == 1


def test_add_comment_supports_multiple_paragraphs():
    """B 类回归：多行批注拆成多个段落，读取时用换行连接。"""
    document = Document()
    run = document.add_paragraph("正文").runs[0]
    comment = document.add_comment(run, text="第一行\n第二行", author="王五")

    assert len(comment.paragraphs) == 2
    assert comment.text == "第一行\n第二行"
    assert document.comments[0].text == "第一行\n第二行"


def test_comments_part_is_resolved_by_relationship_not_name(tmp_path):
    """F3 回归：批注部件按 comments 关系解析，不能只认固定文件名。

    合并类工具会产出 word/comments1.xml 这种命名；若按固定名查找，读取会丢
    失批注，写入还会新增第二条 comments 关系，让 python-docx 拒绝打开。
    """
    document = Document()
    document.add_comment(document.add_paragraph("正文").runs[0], text="原批注", author="A")
    original = tmp_path / "original.docx"
    document.save(original)

    with ZipFile(original) as archive:
        files = {name: archive.read(name) for name in archive.namelist()}
    files["word/comments1.xml"] = files.pop("word/comments.xml")
    files["[Content_Types].xml"] = files["[Content_Types].xml"].replace(
        b"/word/comments.xml", b"/word/comments1.xml",
    )
    files["word/_rels/document.xml.rels"] = files[
        "word/_rels/document.xml.rels"
    ].replace(b"comments.xml", b"comments1.xml")

    renamed = tmp_path / "renamed.docx"
    with ZipFile(renamed, "w", ZIP_DEFLATED) as archive:
        for name, data in files.items():
            archive.writestr(name, data)

    reopened = Document(renamed)
    assert [comment.text for comment in reopened.comments] == ["原批注"]

    reopened.add_comment(
        reopened.add_paragraph("第二段").runs[0], text="新增批注", author="B",
    )
    output = tmp_path / "added.docx"
    reopened.save(output)

    with ZipFile(output) as archive:
        names = archive.namelist()
        rels = archive.read("word/_rels/document.xml.rels").decode("utf-8")
    assert "word/comments1.xml" in names
    assert "word/comments.xml" not in names
    assert rels.count("relationships/comments") == 1
    assert [comment.text for comment in Document(output).comments] == [
        "原批注",
        "新增批注",
    ]

_FONT_TOGGLES = {
    "strike": "strike",
    "double_strike": "dstrike",
    "outline": "outline",
    "shadow": "shadow",
    "emboss": "emboss",
    "imprint": "imprint",
    "hidden": "vanish",
    "small_caps": "smallCaps",
    "snap_to_grid": "snapToGrid",
    "spec_vanish": "specVanish",
    "web_hidden": "webHidden",
    "rtl": "rtl",
    "no_proof": "noProof",
    "math": "oMath",
    "cs_bold": "bCs",
    "cs_italic": "iCs",
    "complex_script": "cs",
}


def test_font_toggle_properties_write_expected_markup():
    """19 个字符级开关：True / False / None 三态都真的落到 rPr 上。"""
    document = Document()
    font = document.add_paragraph("x").runs[0].font

    for name, tag in _FONT_TOGGLES.items():
        setattr(font, name, True)
        assert getattr(font, name) is True, name
        node = font.element.find(qn("w:" + tag))
        assert node is not None, name

        setattr(font, name, False)
        assert getattr(font, name) is False, name
        assert node.get(qn("w:val")) == "0", name

        setattr(font, name, None)
        assert getattr(font, name) is False, name
        assert font.element.find(qn("w:" + tag)) is None, name

    # 上下标共用 w:vertAlign，互斥且各自关闭时不误删对方
    font.subscript = True
    assert font.subscript is True
    assert font.superscript is False

    font.superscript = True
    assert font.superscript is True
    assert font.subscript is False

    font.subscript = False
    assert font.superscript is True

    font.superscript = False
    assert font.element.find(qn("w:vertAlign")) is None


def test_object_scalars_and_iterators(tmp_path):
    """9 个标量 / 迭代器：clear、contains_page_break、iter_inner_content、
    start_type、grid_span、table_direction。"""
    document = Document()
    paragraph = document.add_paragraph("文本")
    run = paragraph.runs[0]

    assert list(run.iter_inner_content()) == ["文本"]
    run.add_tab()
    assert list(run.iter_inner_content()) == ["文本", "\t"]
    assert len(list(paragraph.iter_inner_content())) == 1

    assert paragraph.contains_page_break is False
    assert run.contains_page_break is False
    run._r.append(OxmlElement("w:lastRenderedPageBreak"))
    assert paragraph.contains_page_break is True
    assert run.contains_page_break is True

    run.clear()
    assert run.text == ""
    assert run._r.find(qn("w:t")) is None

    section = document.sections[0]
    assert section.start_type in (None, WD_SECTION_START.NEW_PAGE)
    section.start_type = WD_SECTION_START.ODD_PAGE
    assert section.start_type == WD_SECTION_START.ODD_PAGE

    table = document.add_table(rows=2, cols=2)
    assert table.table_direction is None
    table.table_direction = WD_TABLE_DIRECTION.RTL
    assert table.table_direction == WD_TABLE_DIRECTION.RTL
    table.table_direction = WD_TABLE_DIRECTION.LTR
    assert table.table_direction == WD_TABLE_DIRECTION.LTR

    cell = table.cell(0, 0)
    assert cell.grid_span == 1
    nested = OxmlElement("w:tbl")
    nested.append(OxmlElement("w:tblPr"))
    nested.append(OxmlElement("w:tblGrid"))
    nested.append(OxmlElement("w:tr"))
    cell._tc.append(nested)
    kinds = [type(item).__name__ for item in cell.iter_inner_content()]
    assert "Paragraph" in kinds
    assert "Table" in kinds


def test_plumbing_accessors(tmp_path):
    """内部访问器：element / part / table 返回真实对象。"""
    image = tmp_path / "tiny.png"
    image.write_bytes(_png_bytes())

    document = Document()
    paragraph = document.add_paragraph("x")
    document.add_picture(str(image))

    assert paragraph.paragraph_format.element is not None
    assert paragraph.paragraph_format.part is document.part
    assert paragraph.runs[0].font.element is not None
    assert paragraph.runs[0].font.part is document.part

    table = document.add_table(rows=1, cols=1)
    assert table.part is document.part
    assert table.table is table
    assert table.cell(0, 0).part is document.part

    assert document.sections[0].part is document.part
    assert document.styles.part is not None
    assert document.styles["Normal"].part is not None
    assert document.inline_shapes.part is document.part

