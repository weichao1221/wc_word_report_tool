"""MCP 服务端：把 :class:`~wc_word_report_tool.word.WordFormatter` 暴露成 AI 可调用的工具。

安装::

    pip install wc_word_report_tool[mcp]

启动::

    wc-report-mcp

``mcp`` SDK 是可选依赖，因此这里对 ``server`` 采用惰性导入 —— 未安装 SDK 时
``import wc_word_report_tool.mcp`` 仍然可用，只是取不到 ``main``。
"""

__all__ = ["main"]


def main() -> None:
    """启动 MCP 服务端（命令行入口 ``wc-report-mcp``）。"""
    from .server import main as _main

    _main()
