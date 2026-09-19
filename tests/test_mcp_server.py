"""MCP 服务端协议层测试。

分两层：
1. 直接查询服务端实例的工具清单与 JSON Schema；
2. 起一个真实的 ``python -m wc_word_report_tool.mcp`` 子进程，走 stdio JSON-RPC
   完成握手、tools/list 和 tools/call —— 这是对入口和传输最完整的验证。
"""

import asyncio
import json
import os
import selectors
import subprocess
import sys
import time

import pytest

pytest.importorskip("mcp")

from wc_word_report_tool.mcp.server import build_server, load_server_class

EXPECTED_TOOLS = {
    "word_create_report",
    "word_open_report",
    "word_save_report",
    "word_close_report",
    "word_describe_report",
    "word_set_default_font",
    "word_set_page_margins",
    "word_set_document_language",
    "word_add_heading",
    "word_add_body",
    "word_add_body_list",
    "word_add_body_segments",
    "word_add_blank_lines",
    "word_add_page_break",
    "word_add_cover_text",
    "word_add_right_text",
    "word_add_date",
    "word_insert_image",
    "word_add_table",
    "word_merge_cells",
    "word_format_cell",
    "word_set_table_borders",
    "word_list_blocks",
    "word_replace_text",
    "word_insert_cell_image",
    "word_set_header",
    "word_set_footer",
    "word_set_header_parts",
    "word_set_footer_parts",
    "word_set_different_first_page",
    "word_set_different_odd_even",
    "word_clear_header_footer",
    "word_set_header_image",
    "word_insert_section",
    "word_set_section_page",
    "word_set_page_numbers",
    "word_restart_page_numbering",
    "word_set_page_number_start",
    "word_set_page_number_format",
    "word_add_toc",
    "word_set_toc_level_style",
    "word_set_paragraph_style",
    "word_add_custom_heading",
    "word_add_cover_page",
    "word_add_signature_page",
    "word_rmb_upper",
    "word_describe_api",
    "word_call",
}


def _list_tools():
    return asyncio.run(build_server().list_tools())


def _schemas():
    return {tool.name: tool.input_schema for tool in _list_tools()}


# ====================================================================
# 服务端实例
# ====================================================================


def test_expected_tool_surface_is_registered():
    assert {tool.name for tool in _list_tools()} == EXPECTED_TOOLS


def test_server_class_detection_returns_known_major():
    _, major = load_server_class()
    assert major in (1, 2)


def test_schemas_expose_real_json_types():
    schema = _schemas()["word_add_heading"]

    assert schema["properties"]["doc_id"]["type"] == "string"
    assert schema["properties"]["text"]["type"] == "string"
    assert schema["properties"]["level"]["type"] == "integer"
    assert schema["properties"]["indent"]["type"] == "boolean"
    assert set(schema["required"]) == {"doc_id", "text"}


def test_font_size_accepts_number_or_chinese_size():
    font_size = _schemas()["word_add_heading"]["properties"]["font_size"]

    assert {"type": "number"} in font_size["anyOf"]
    assert {"type": "string"} in font_size["anyOf"]


def test_image_tools_expose_both_input_channels():
    properties = _schemas()["word_insert_image"]["properties"]

    assert "image_path" in properties
    assert "image_base64" in properties


def test_every_tool_has_a_description():
    for tool in _list_tools():
        assert tool.description, f"{tool.name} 缺少描述"


# ====================================================================
# 真实 stdio 子进程
# ====================================================================


class StdioClient:
    """最小 stdio JSON-RPC 客户端，避免依赖 mcp SDK 的测试辅助 API（各版本名称不同）。"""

    def __init__(self, process):
        self._process = process
        self._selector = selectors.DefaultSelector()
        self._selector.register(process.stdout, selectors.EVENT_READ)
        self._next_id = 0

    def _write(self, payload):
        self._process.stdin.write(json.dumps(payload) + "\n")
        self._process.stdin.flush()

    def notify(self, method, params=None):
        message = {"jsonrpc": "2.0", "method": method}
        if params is not None:
            message["params"] = params
        self._write(message)

    def request(self, method, params=None, timeout=30.0):
        self._next_id += 1
        request_id = self._next_id
        message = {"jsonrpc": "2.0", "id": request_id, "method": method}
        if params is not None:
            message["params"] = params
        self._write(message)
        return self._await(request_id, timeout)

    def _await(self, request_id, timeout):
        deadline = time.monotonic() + timeout
        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0 or not self._selector.select(remaining):
                raise TimeoutError(f"等待 {request_id} 的响应超时")
            line = self._process.stdout.readline()
            if not line:
                raise RuntimeError(f"服务端提前退出。stderr:\n{self._stderr_tail()}")
            line = line.strip()
            if not line:
                continue
            message = json.loads(line)
            if message.get("id") == request_id:
                return message

    def _stderr_tail(self):
        if self._process.poll() is None:
            return "<进程仍在运行，未读取 stderr>"
        return self._process.stderr.read()[-2000:]

    def handshake(self):
        response = self.request(
            "initialize",
            {
                "protocolVersion": "2025-06-18",
                "capabilities": {},
                "clientInfo": {"name": "pytest", "version": "0"},
            },
        )
        self.notify("notifications/initialized")
        return response


@pytest.fixture
def stdio_client(tmp_path):
    process = subprocess.Popen(
        [sys.executable, "-m", "wc_word_report_tool.mcp"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        bufsize=1,
        env={
            **os.environ,
            "WC_REPORT_MCP_OUTPUT_DIR": str(tmp_path / "out"),
            "PYTHONUNBUFFERED": "1",
        },
    )
    client = StdioClient(process)
    try:
        yield client
    finally:
        process.kill()
        process.wait(timeout=10)


def test_stdio_handshake_reports_server_identity(stdio_client):
    response = stdio_client.handshake()

    assert "result" in response, response
    assert response["result"]["serverInfo"]["name"] == "wc-word-report"


def test_stdio_lists_the_same_tool_surface(stdio_client):
    stdio_client.handshake()

    listing = stdio_client.request("tools/list")

    assert {tool["name"] for tool in listing["result"]["tools"]} == EXPECTED_TOOLS


def test_stdio_tool_call_round_trip(stdio_client):
    stdio_client.handshake()

    response = stdio_client.request(
        "tools/call",
        {"name": "word_rmb_upper", "arguments": {"value": 100200.03}},
    )

    assert response["result"]["isError"] is False
    assert "壹拾万零贰佰元零叁分" in json.dumps(response["result"], ensure_ascii=False)


def test_stdio_tool_error_surfaces_as_is_error(stdio_client):
    stdio_client.handshake()

    response = stdio_client.request(
        "tools/call",
        {"name": "word_describe_report", "arguments": {"doc_id": "nope"}},
    )

    assert response["result"]["isError"] is True
    assert "未知的 doc_id" in json.dumps(response["result"], ensure_ascii=False)
