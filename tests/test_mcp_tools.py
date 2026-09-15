"""MCP 工具层测试：直接调用工具函数，不经过 MCP 协议。

工具函数是普通 Python 函数（不依赖 mcp SDK），所以可以像测普通库一样测试它们。
协议层的冒烟测试见 tests/test_mcp_server.py。
"""

from pathlib import Path
from zipfile import ZipFile

import pytest

from wc_word_report_tool.mcp import tools
from wc_word_report_tool.mcp.errors import ToolError
from wc_word_report_tool.mcp.session import ReportSessionError, SessionStore

TEST_PNG_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk"
    "/x8AAusB9Y9Z4k8AAAAASUVORK5CYII="
)


@pytest.fixture(autouse=True)
def clean_store(monkeypatch, tmp_path):
    """每个测试用独立的输出目录和空会话表。"""
    monkeypatch.setenv("WC_REPORT_MCP_OUTPUT_DIR", str(tmp_path / "out"))
    monkeypatch.delenv("WC_REPORT_MCP_DEFAULT_LOGO", raising=False)
    tools.reset_store()
    yield
    tools.reset_store()


def _read_zip_xml(docx_path: Path, inner_path: str) -> str:
    with ZipFile(docx_path) as zf:
        return zf.read(inner_path).decode("utf-8", errors="ignore")


def _zip_names(docx_path: Path) -> list:
    with ZipFile(docx_path) as zf:
        return zf.namelist()


# ====================================================================
# 会话生命周期
# ====================================================================


def test_create_report_returns_usable_doc_id():
    created = tools.word_create_report(title="测试报告", author="willcha")

    assert created["doc_id"]
    assert created["sections"] == 1
    assert created["paragraphs"] == 0

    described = tools.word_describe_report(created["doc_id"])
    assert described["doc_id"] == created["doc_id"]
    assert described["paragraphs_per_section"] == [0]


def test_unknown_doc_id_gives_actionable_error():
    with pytest.raises(ToolError) as excinfo:
        tools.word_add_body("does-not-exist", "正文")

    assert "未知的 doc_id" in str(excinfo.value)
    assert "word_create_report" in str(excinfo.value)


def test_close_report_invalidates_doc_id():
    doc_id = tools.word_create_report()["doc_id"]

    assert tools.word_close_report(doc_id)["closed"] == doc_id

    with pytest.raises(ToolError):
        tools.word_describe_report(doc_id)


def test_save_without_path_reuses_previous_target(tmp_path):
    doc_id = tools.word_create_report()["doc_id"]

    first = tools.word_save_report(doc_id)
    second = tools.word_save_report(doc_id)

    assert first["saved_to"] == second["saved_to"]
    assert Path(first["saved_to"]).parent == tmp_path / "out"
    assert first["size_bytes"] > 0


def test_store_reaps_expired_sessions():
    store = SessionStore(ttl_seconds=60)
    session = store.create()
    session.last_used -= 120

    with pytest.raises(ReportSessionError):
        store.get(session.doc_id)


def test_store_evicts_oldest_when_full():
    store = SessionStore(max_sessions=2)
    oldest = store.create()
    newest = store.create()
    newest.last_used = oldest.last_used + 10

    survivor = store.create()

    assert store.get(newest.doc_id).doc_id == newest.doc_id
    assert survivor.doc_id != oldest.doc_id
    with pytest.raises(ReportSessionError):
        store.get(oldest.doc_id)


# ====================================================================
# 无状态工具
# ====================================================================


def test_rmb_upper_needs_no_doc_id():
    assert tools.word_rmb_upper(123456.78)["chinese_upper"] == (
        "壹拾贰万叁仟肆佰伍拾陆元柒角捌分"
    )
    assert tools.word_rmb_upper("1,234.5")["chinese_upper"] == "壹仟贰佰叁拾肆元伍角"


def test_describe_api_lists_formatter_methods():
    text = tools.word_describe_api()

    for method in (
        "heading",
        "body",
        "add_table",
        "add_toc",
        "fengmian_jiesuan",
        "qianfaye",
        "set_page_number_from_section",
    ):
        assert method in text


# ====================================================================
# 兜底工具
# ====================================================================


def test_word_call_invokes_arbitrary_method():
    doc_id = tools.word_create_report()["doc_id"]

    result = tools.word_call(doc_id, "body", {"text": "来自兜底工具", "indent": False})

    assert result["method"] == "body"
    assert result["paragraphs"] == 1


def test_word_call_resolves_section_index_to_object():
    doc_id = tools.word_create_report()["doc_id"]

    tools.word_call(doc_id, "set_header", {"section": -1, "text": "兜底页眉"})

    saved = tools.word_save_report(doc_id, "call-header.docx")
    head_path = "word/header1.xml"
    assert head_path in _zip_names(Path(saved["saved_to"]))
    assert "兜底页眉" in _read_zip_xml(Path(saved["saved_to"]), head_path)


def test_word_call_rejects_private_and_unknown_methods():
    doc_id = tools.word_create_report()["doc_id"]

    with pytest.raises(ToolError):
        tools.word_call(doc_id, "_ensure_update_fields_on_open", {})
    with pytest.raises(ToolError):
        tools.word_call(doc_id, "not_a_real_method", {})


# ====================================================================
# 内容构建
# ====================================================================


def test_add_body_defaults_use_character_indent(tmp_path):
    doc_id = tools.word_create_report()["doc_id"]
    tools.word_add_body(doc_id, "正文内容")

    saved = tools.word_save_report(doc_id, "body.docx")
    document_xml = _read_zip_xml(Path(saved["saved_to"]), "word/document.xml")

    assert 'w:eastAsia="宋体"' in document_xml
    assert 'w:firstLineChars="200"' in document_xml


def test_add_body_list_writes_every_paragraph():
    doc_id = tools.word_create_report()["doc_id"]

    result = tools.word_add_body_list(doc_id, ["第一段。", "第二段。", "第三段。"])

    assert result["added"] == 3
    assert result["paragraphs"] == 3


def test_add_table_reports_shape():
    doc_id = tools.word_create_report()["doc_id"]

    result = tools.word_add_table(
        doc_id,
        headers=["序号", "项目", "金额（万元）"],
        rows=[["1", "建筑工程", "1200.50"], ["2", "安装工程", "850.00"]],
        col_widths=[2, 6, 4],
    )

    assert result["rows"] == 3
    assert result["columns"] == 3
    assert result["table_index"] == 0


def test_format_cell_out_of_range_explains_bounds():
    doc_id = tools.word_create_report()["doc_id"]
    tools.word_add_table(doc_id, ["标题"], [["内容"]])

    with pytest.raises(ToolError) as excinfo:
        tools.word_format_cell(doc_id, row=5, col=0, text="越界")

    assert "超出范围" in str(excinfo.value)
    assert "2 行" in str(excinfo.value)


def test_add_date_chinese_style():
    doc_id = tools.word_create_report()["doc_id"]
    tools.word_add_date(doc_id, value="2026-07-16", style="chinese")

    saved = tools.word_save_report(doc_id, "date.docx")

    assert "二零二六年七月十六日" in _read_zip_xml(
        Path(saved["saved_to"]), "word/document.xml"
    )


def test_add_date_rejects_unknown_style():
    doc_id = tools.word_create_report()["doc_id"]

    with pytest.raises(ToolError) as excinfo:
        tools.word_add_date(doc_id, style="lunar")

    assert "chinese" in str(excinfo.value)


def test_target_validation_is_converted_to_tool_error():
    """参数校验发生在会话上下文内，才能被转换成模型可见的 ToolError。"""
    doc_id = tools.word_create_report()["doc_id"]

    with pytest.raises(ToolError) as excinfo:
        tools.word_clear_header_footer(doc_id, target="sidebar")
    assert "header/footer/both" in str(excinfo.value)

    with pytest.raises(ToolError) as excinfo:
        tools.word_set_header_image(doc_id, target="sidebar", width_cm=2)
    assert "header/footer" in str(excinfo.value)


def test_section_index_out_of_range():
    doc_id = tools.word_create_report()["doc_id"]

    with pytest.raises(ToolError) as excinfo:
        tools.word_set_header(doc_id, "页眉", section_index=9)

    assert "section_index 9 超出范围" in str(excinfo.value)


# ====================================================================
# 图片双通道
# ====================================================================


def test_insert_image_accepts_base64():
    doc_id = tools.word_create_report()["doc_id"]

    tools.word_insert_image(doc_id, width_cm=3, image_base64=TEST_PNG_B64)
    saved = tools.word_save_report(doc_id, "image.docx")

    assert any(
        name.startswith("word/media/") for name in _zip_names(Path(saved["saved_to"]))
    )


def test_insert_image_accepts_data_uri(tmp_path):
    doc_id = tools.word_create_report()["doc_id"]

    tools.word_insert_image(
        doc_id, width_cm=3, image_base64=f"data:image/png;base64,{TEST_PNG_B64}"
    )
    saved = tools.word_save_report(doc_id, "image-datauri.docx")

    assert any(
        name.startswith("word/media/") for name in _zip_names(Path(saved["saved_to"]))
    )


def test_insert_image_requires_a_source():
    doc_id = tools.word_create_report()["doc_id"]

    with pytest.raises(ToolError) as excinfo:
        tools.word_insert_image(doc_id, width_cm=3)

    assert "image_path" in str(excinfo.value)


def test_insert_image_rejects_both_sources():
    doc_id = tools.word_create_report()["doc_id"]

    with pytest.raises(ToolError) as excinfo:
        tools.word_insert_image(
            doc_id, width_cm=3, image_path="a.png", image_base64=TEST_PNG_B64
        )

    assert "只能传一个" in str(excinfo.value)


def test_cover_page_without_logo_explains_the_fix():
    doc_id = tools.word_create_report()["doc_id"]

    with pytest.raises(ToolError) as excinfo:
        tools.word_add_cover_page(doc_id)

    assert "Logo" in str(excinfo.value)
    assert "WC_REPORT_MCP_DEFAULT_LOGO" in str(excinfo.value)


def test_cover_page_falls_back_to_default_logo_env(monkeypatch, tmp_path):
    logo = tmp_path / "logo.png"
    import base64

    logo.write_bytes(base64.b64decode(TEST_PNG_B64))
    monkeypatch.setenv("WC_REPORT_MCP_DEFAULT_LOGO", str(logo))

    doc_id = tools.word_create_report()["doc_id"]
    tools.word_add_cover_page(doc_id, project_name="某某工程")
    saved = tools.word_save_report(doc_id, "cover.docx")

    document_xml = _read_zip_xml(Path(saved["saved_to"]), "word/document.xml")
    assert "某某工程" in document_xml


# ====================================================================
# 端到端
# ====================================================================


def test_full_report_end_to_end():
    doc_id = tools.word_create_report(title="测试项目结算审核报告")["doc_id"]

    tools.word_add_cover_page(
        doc_id,
        project_name="某某工程",
        entrusting_unit="某委托单位",
        compiling_unit="某编制单位",
        date="2026-07-16",
        image_base64=TEST_PNG_B64,
    )
    tools.word_set_default_font(doc_id, font_size="三号", cn_font="仿宋_GB2312")
    tools.word_insert_section(doc_id, inherit_header=False, inherit_footer=False)
    tools.word_add_toc(doc_id, title="目  录", levels_min=1, levels_max=3)
    tools.word_add_heading(doc_id, "一、项目概况", level=1)
    tools.word_add_body_list(doc_id, ["本项目位于……。", "项目背景说明……。"])
    tools.word_add_table(
        doc_id,
        headers=["序号", "项目", "金额（万元）"],
        rows=[["1", "建筑工程", "1200.50"], ["2", "安装工程", "850.00"]],
    )
    tools.word_set_header(doc_id, "某某工程 结算审核报告", section_index=-1)
    tools.word_set_page_numbers(
        doc_id, start_section_index=1, start=1, prefix="第 ", suffix=" 页",
    )

    saved = tools.word_save_report(doc_id, "demo.docx")
    path = Path(saved["saved_to"])

    assert path.is_file()
    assert path.suffix == ".docx"

    document_xml = _read_zip_xml(path, "word/document.xml")
    assert "一、项目概况" in document_xml
    assert 'w:firstLineChars="200"' in document_xml
    assert "TOC" in document_xml

    settings_xml = _read_zip_xml(path, "word/settings.xml")
    assert "w:updateFields" in settings_xml
    assert "zh-CN" in settings_xml

    footers = [name for name in _zip_names(path) if name.startswith("word/footer")]
    assert footers
    assert any("PAGE" in _read_zip_xml(path, name) for name in footers)


def test_open_report_round_trips_an_existing_file():
    doc_id = tools.word_create_report()["doc_id"]
    tools.word_add_body(doc_id, "原始内容")
    saved = tools.word_save_report(doc_id, "source.docx")

    reopened = tools.word_open_report(saved["saved_to"])
    assert reopened["source_path"] == saved["saved_to"]

    tools.word_add_body(reopened["doc_id"], "追加内容")

    again = tools.word_save_report(reopened["doc_id"], "appended.docx")
    document_xml = _read_zip_xml(Path(again["saved_to"]), "word/document.xml")

    assert "原始内容" in document_xml
    assert "追加内容" in document_xml
