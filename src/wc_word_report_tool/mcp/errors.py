"""把内部异常转换成 MCP 客户端能看到的消息。

mcp SDK 只把 ``ToolError`` 的消息透传给模型；**其它任何异常一律被掩码**成
``Error executing tool <name>``（见 ``mcp.server.mcpserver.exceptions`` 的说明）。
所以如果直接抛 ``ReportSessionError``，AI 就看不到"未知的 doc_id: 'xxx'"这类
可以自行纠正的提示，只能看到一句无信息量的报错。

使用 mcp 2.x 的 ``mcp.server.mcpserver``；SDK 缺失时退化为普通异常，
保证工具层仍可脱离 mcp 单独测试。
"""

from __future__ import annotations

from zipfile import BadZipFile, LargeZipFile

from docx.opc.exceptions import PackageNotFoundError

from .session import ReportSessionError


def _load_tool_error():
    try:
        from mcp.server.mcpserver.exceptions import ToolError

        return ToolError
    except ImportError:
        pass

    class ToolError(Exception):
        """mcp SDK 未安装时的占位实现。"""

    return ToolError


ToolError = _load_tool_error()

#: 这些异常都代表"可预期的输入问题"，消息应当原样透传给模型
CONVERTIBLE_ERRORS = (
    ReportSessionError,
    ValueError,
    TypeError,
    KeyError,
    IndexError,
    OSError,
    PackageNotFoundError,
    BadZipFile,
    LargeZipFile,
)


def to_tool_error(exc: BaseException) -> Exception:
    """把异常转换成消息可见的 ToolError。"""
    message = str(exc).strip() or exc.__class__.__name__
    return ToolError(message)
