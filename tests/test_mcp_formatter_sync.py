"""MCP 工具层与 ``WordFormatter`` 核心类的同步性校验。

这组测试专门守住「MCP 有工具、底层却没有可调用的方法」这类漂移，而且不靠人工维护清单：

* **正向**：解析 ``tools.py`` 的 AST，把所有 ``WordFormatter.<name>`` 与
  ``session.formatter.<name>`` 触达的方法逐个断言真实存在。谁把底层方法改名/删掉，
  却忘了改工具层，这里立刻红。
* **反向**：断言这些方法都是**公开**的（不以 ``_`` 开头），否则「MCP 能调、普通代码
  不好调」的不对称就又回来了（少数明确的内部方法走白名单）。
* **可达性**：断言 ``TOOLS`` 里的每个工具都能按名字从 ``tools`` 模块取到、可调用、
  有描述，且注册到服务端的工具集合与 ``TOOLS`` 完全一致。
"""

import ast
import asyncio
import inspect
from pathlib import Path

import pytest

from wc_word_report_tool import WordFormatter
from wc_word_report_tool.mcp import tools

# tools.py 里合法使用的私有方法（保存前刷新域，属于工具层的实现细节而非对外能力）
_ALLOWED_PRIVATE = {"_ensure_update_fields_on_open"}

_TREE = ast.parse(Path(inspect.getsourcefile(tools)).read_text(encoding="utf-8"))


def _formatter_method_names() -> set:
    """抽出 tools.py 中通过 WordFormatter 或 formatter 实例触达的属性名。"""
    names = set()
    for node in ast.walk(_TREE):
        if not isinstance(node, ast.Attribute):
            continue
        # WordFormatter.set_header(...)
        if isinstance(node.value, ast.Name) and node.value.id == "WordFormatter":
            names.add(node.attr)
        # session.formatter.heading(...)
        elif isinstance(node.value, ast.Attribute) and node.value.attr == "formatter":
            names.add(node.attr)
    return names


def test_tools_module_reaches_into_formatter_at_all():
    """先确认探测逻辑本身有效，避免 AST 匹配写错后测试永远通过。"""
    reached = _formatter_method_names()

    assert "heading" in reached
    assert "add_table" in reached
    assert len(reached) > 15


def test_every_formatter_method_used_by_tools_exists():
    missing = sorted(
        name for name in _formatter_method_names()
        if not callable(getattr(WordFormatter, name, None))
    )

    assert not missing, f"tools.py 调用了 WordFormatter 上不存在的方法: {missing}"


def test_every_formatter_method_used_by_tools_is_public():
    """保证「MCP 能做的事，普通代码也能直接调」——不靠私有方法走后门。"""
    private = sorted(
        name for name in _formatter_method_names()
        if name.startswith("_") and name not in _ALLOWED_PRIVATE
    )

    assert not private, f"tools.py 依赖了未登记的内部方法: {private}"


def test_tool_names_match_their_module_functions():
    for tool in tools.TOOLS:
        assert tool.__name__.startswith("word_"), f"{tool.__name__} 缺少 word_ 前缀"
        assert getattr(tools, tool.__name__, None) is tool, (
            f"{tool.__name__} 与实际函数对象不一致"
        )


def test_tool_registry_has_no_duplicates():
    names = [tool.__name__ for tool in tools.TOOLS]

    assert len(names) == len(set(names)), "TOOLS 中存在重复注册"


def test_every_tool_is_callable_and_documented():
    for tool in tools.TOOLS:
        assert callable(tool), f"{tool.__name__} 不可调用"
        assert (tool.__doc__ or "").strip(), f"{tool.__name__} 缺少描述"


def test_registered_server_surface_equals_tools_registry():
    """服务端真实注册的工具集合必须与 TOOLS 完全一致（不多不少）。"""
    pytest.importorskip("mcp")

    from wc_word_report_tool.mcp.server import build_server

    registered = {tool.name for tool in asyncio.run(build_server().list_tools())}

    assert registered == {tool.__name__ for tool in tools.TOOLS}


def test_api_help_covers_every_public_formatter_method():
    """兜底能力（word_call）必须能覆盖到全部公开方法，否则长尾方法就是不可达的。"""
    api_text = tools.word_describe_api()

    public = {
        name for name, value in vars(WordFormatter).items()
        if not name.startswith("_")
        and (inspect.isfunction(value) or isinstance(value, staticmethod))
    }

    missing = sorted(name for name in public if f"{name}(" not in api_text)
    assert not missing, f"word_describe_api 未列出这些公开方法: {missing}"
