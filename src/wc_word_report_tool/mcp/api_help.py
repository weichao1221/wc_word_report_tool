"""反射 ``WordFormatter`` 的公开方法，生成给 AI 看的 API 清单。

这是 MCP 兜底能力（``word_describe_api`` + ``word_call``）的支撑：即使某个长尾
方法没有被包装成独立工具，AI 也能查到这里再去调用。

本模块不依赖 ``mcp`` SDK，可独立测试。
"""

from __future__ import annotations

import inspect
from functools import lru_cache

from ..word import WordFormatter


def _unwrap(name: str) -> tuple[object, bool]:
    """返回 (函数, 是否静态方法)。"""
    raw = inspect.getattr_static(WordFormatter, name)
    if isinstance(raw, staticmethod):
        return raw.__func__, True
    return raw, False


def _signature_text(name: str) -> str:
    func, is_static = _unwrap(name)
    try:
        signature = inspect.signature(func)
    except (TypeError, ValueError):
        return f"{name}(...)"
    if not is_static:
        params = list(signature.parameters.values())[1:]
        signature = signature.replace(parameters=params)
    return f"{name}{signature}"


def _first_doc_line(name: str) -> str:
    func, _ = _unwrap(name)
    doc = (func.__doc__ or "").strip()
    return doc.splitlines()[0] if doc else ""


def _public_methods() -> list[tuple[str, bool]]:
    collected = []
    for name in sorted(vars(WordFormatter)):
        if name.startswith("_"):
            continue
        attribute = inspect.getattr_static(WordFormatter, name)
        if isinstance(attribute, staticmethod):
            collected.append((name, True))
        elif inspect.isfunction(attribute):
            collected.append((name, False))
    return collected


@lru_cache(maxsize=1)
def formatter_api_text() -> str:
    """生成 ``WordFormatter`` 的完整 API 说明文本。"""
    methods = _public_methods()
    instance_methods = [(n, s) for n, s in methods if not s]
    static_methods = [(n, s) for n, s in methods if s]

    lines = [
        f"WordFormatter 可用方法（共 {len(methods)} 个）。",
        "",
        "调用方式：word_call(doc_id, method, params)，params 为关键字参数字典。",
        "文档级方法（下面前两组）会自动绑定当前 doc_id 对应的文档。",
        "",
        f"【实例方法 · {len(instance_methods)} 个】只需 doc_id，无需手动传 doc",
    ]
    for name, _ in instance_methods:
        doc = _first_doc_line(name)
        lines.append(f"  {_signature_text(name)}" + (f"\n      {doc}" if doc else ""))

    lines.append("")
    lines.append(
        f"【静态方法 · {len(static_methods)} 个】"
        "部分需要传入 section / table / cell 对象；"
        "在 word_call 中，section 和 table 参数可直接传整数索引（-1 表示最后一个）"
    )
    for name, _ in static_methods:
        doc = _first_doc_line(name)
        lines.append(f"  {_signature_text(name)}" + (f"\n      {doc}" if doc else ""))

    lines.append("")
    lines.append("【模块级函数】")
    lines.append(
        "  number_to_chinese_upper(value) — 把数字转换为人民币金额中文大写"
        "（已包装成 word_rmb_upper 工具）"
    )
    return "\n".join(lines)
