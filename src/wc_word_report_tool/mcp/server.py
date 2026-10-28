"""MCP 服务端：把 ``WordFormatter`` 暴露成可被 AI 调用的工具集合。

启动方式::

    wc-report-mcp                       # stdio，供 Claude Desktop / Cursor 等本地客户端
    wc-report-mcp --transport streamable-http --port 8000
"""

import argparse
import ipaddress
import sys

if sys.version_info < (3, 10):
    raise SystemExit(
        "MCP 服务端需要 Python 3.10 及以上（当前 {}.{}）。".format(*sys.version_info[:2])
        + "请升级 Python，或仅使用 pip install wc_word_report_tool 的基础功能。"
    )

from . import tools

INSTRUCTIONS = """中文 Word 报告编制工具。用 python-docx 生成符合中国公文格式的报告。

## 基本流程

1. `word_create_report` 新建文档，拿到 `doc_id`（后续所有调用都要带上它）
2. 用 `word_add_cover_page` / `word_add_toc` / `word_add_heading` / `word_add_body_list`
   / `word_add_table` 等按顺序写入内容
3. 用 `word_set_page_numbers` 设置页码，`word_set_header` / `word_set_footer` 设置页眉页脚
4. `word_save_report` 落盘

关闭、TTL 回收或容量淘汰前，未保存的会话会自动保存到输出目录；结果里会给出
`auto_saved_to` 或 `maintenance_events`。如果自动保存失败，会话不会被丢弃。

## 填模板（在已有文档里补内容）

上面的流程是「从空白开始、按顺序追加」；如果要往**一份现成的 .docx**（招投标文件、
表单）里填空，走另一条路，两者不要混：

1. `word_open_report` 打开模板 → `word_list_blocks` 看清有哪些段落和单元格
2. 文本占位用 `word_replace_text`（给 `paragraph_index`，或 `table_index`/`row`/`col`）。
   它只改写原有 run 的文本，字体字号自动沿用原文，**不需要你重新指定格式**
3. 证件照这类要贴进格子的图片用 `word_insert_cell_image`
4. `word_save_report` 存到目标路径（建议传绝对路径）

注意：`word_add_body` / `word_add_body_list` / `word_insert_image` 都是**往文档末尾追加**，
在模板上误用会把内容写到最后一页，而不是填进占位处。

## 读取格式（照着一份现成文档复刻版式）

`word_open_report` 打开源文档后，用 `word_describe_format_profile` 读取格式：
返回 `format_profile`，其中 `styles` 给出标题一/标题二/正文等样式的字体、字号、
行距、首行缩进、左右缩进、段前段后和对齐（含每个字段来自默认值/样式/直接格式），
`observed` 给出正文逐块实测结果，`document` 覆盖页面、页眉页脚、页码与自动编号，
`python_code` 是可直接粘贴执行的 `WordFormatter` 调用。

自动编号、页眉页脚中的域/图片、文档网格等无法由现有写接口重建，结果里的
`notes` 会逐项列出，交付时请提醒用户手工处理。

## 重要约定

- **顺序即文档顺序**：工具按调用顺序追加内容，请按报告的阅读顺序调用。
- **批量写入**：多段正文用 `word_add_body_list` 一次传数组，比逐段调用高效。
- **字号**支持数字（pt）或中文字号字符串：`14`、`"三号"`、`"小五"`。
  不确定 pt 对应哪一号时，用 `word_convert_font_size` 换算：它按**吸附**规则
  取最接近的一档，返回 `font_size` 与 `exact`，`exact=False` 表示该 pt 是被吸附
  过去的（`delta_pt` 是差值）；正好落在两档正中间时用 `ties="larger"|"smaller"` 指定取向。
- **对齐**支持 `"居中"` / `"两端对齐"` / `"居左"` / `"居右"`，也接受 `center` / `justify` / `left` / `right`。
- **单位**：页边距、图片宽度、列宽是 cm；字号、段间距是 pt。
- `section_index` / `table_index` 默认 `-1`，表示最后一节 / 最后一个表格。
- MCP HTTP 仅允许绑定本机回环地址；当前版本没有认证和路径隔离。

## 限制（请提前告知用户）

- 目录页码需在 Word/OnlyOffice 中刷新域（F9）后才显示；生成的文档已标记
  `updateFields`，Word 打开时会自动刷新。你无法自行校验目录页码，请在交付时提醒用户确认。
- 字体必须在运行本服务端的机器上已安装，否则 Word 会回退到其他字体。
- 封面和签发页需要 Logo 图片：提供 `image_path` 或 `image_base64`，
  或在服务端配置环境变量 `WC_REPORT_MCP_DEFAULT_LOGO`。

## 兜底能力

如果某个 `WordFormatter` 方法没有被包装成独立工具，用 `word_describe_api` 查看
全部方法签名，再用 `word_call(doc_id, method, params)` 直接调用。

`word_call` 对象参数支持：
- `section` / `table`：整数索引，`-1` 表示最后一个。
- `cell`：`{"table_index": -1, "row": 0, "col": 0}`。
- `paragraph`：正文段落下标，或
  `{"table_index": -1, "row": 0, "col": 0, "paragraph_index": 0}`。
- `run`：`{"paragraph": <paragraph 引用>, "run_index": 0}`。
"""


def load_server_class():
    """返回 ``(服务端类, 主版本号)``，兼容 mcp 1.x 与 2.x。

    mcp 2.x 把 ``FastMCP`` 改名为 ``MCPServer``，并且把 host/port 从构造函数挪到了
    ``run()`` 参数上，因此需要区分处理。
    """
    try:
        from mcp.server.mcpserver import MCPServer
    except ImportError:
        from mcp.server.fastmcp import FastMCP

        return FastMCP, 1
    return MCPServer, 2


def _is_loopback_host(host: str) -> bool:
    if host.lower() == "localhost":
        return True
    try:
        return ipaddress.ip_address(host).is_loopback
    except ValueError:
        return False


def build_server(host: str = "127.0.0.1", port: int = 8000):
    """构造 MCP 服务端实例并注册全部工具。"""
    if not _is_loopback_host(host):
        raise ValueError(
            f"拒绝绑定非本机地址 {host!r}。当前版本没有认证和路径隔离，"
            "仅允许 localhost / 127.0.0.1 / ::1。"
        )
    server_class, major = load_server_class()
    options = {"instructions": INSTRUCTIONS}
    if major == 1:
        # mcp 1.x 通过构造函数的 settings 配置监听地址。
        options.update(host=host, port=port)
    server = server_class("wc-word-report", **options)
    tools.register(server)
    return server


mcp = build_server()


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="wc-report-mcp",
        description="中文 Word 报告编制的 MCP 服务端",
    )
    parser.add_argument(
        "--transport",
        choices=("stdio", "streamable-http"),
        default="stdio",
        help="传输方式，默认 stdio（本地客户端）；streamable-http 用于常驻服务",
    )
    parser.add_argument("--host", default="127.0.0.1", help="HTTP 传输监听地址")
    parser.add_argument("--port", type=int, default=8000, help="HTTP 传输监听端口")
    args = parser.parse_args()

    try:
        server = build_server(host=args.host, port=args.port)
    except ValueError as exc:
        parser.error(str(exc))
    if args.transport == "stdio":
        server.run(transport="stdio")
    else:
        server.run(transport=args.transport, host=args.host, port=args.port)
