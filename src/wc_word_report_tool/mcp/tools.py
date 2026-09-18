"""MCP 工具函数。

每个函数都是普通 Python 函数，模块本身不依赖 ``mcp`` SDK，因此可以直接单元测试；
由 :func:`register` 统一注册到 FastMCP 实例。

命名统一加 ``word_`` 前缀：MCP 客户端会把多个 server 的工具名拍平到同一命名空间，
前缀可以避免与其他文档类 server 冲突。

设计分层：

* 会话工具 —— 创建 / 打开 / 保存 / 关闭 / 查看文档
* 语义工具 —— 覆盖 AI 生成报告的绝大多数路径，参数带完整默认值
* 兜底工具 —— ``word_describe_api`` + ``word_call`` 保证对 ``WordFormatter`` 的
  全量覆盖，长尾方法不会因为没被单独包装而不可用
"""

import time
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Dict, Iterator, List, Optional, Union

from docx.text.paragraph import Paragraph

from ..numbers import number_to_chinese_upper
from ..word import (
    DEFAULT_CN_FONT,
    DEFAULT_EN_FONT,
    DEFAULT_HEADER_FOOTER_FONT,
    DEFAULT_HEADER_FOOTER_SIZE,
    WordFormatter,
)
from .api_help import formatter_api_text
from .errors import CONVERTIBLE_ERRORS, ToolError, to_tool_error
from .session import (
    ReportSessionError,
    Session,
    SessionStore,
    default_logo_path,
    materialize_image,
    resolve_input_path,
    resolve_output_path,
)

FontSize = Union[float, str]

_STORE = SessionStore()


def reset_store() -> None:
    """清空全部会话，供测试使用。"""
    _STORE.close_all()


@contextmanager
def _op(doc_id: str) -> Iterator[Session]:
    """取出会话并在全局锁内执行文档改写。

    可预期的异常在这里统一转成 ``ToolError``，否则 mcp SDK 会把消息掩码成
    ``Error executing tool <name>``，AI 就拿不到可自纠的中文提示了。
    """
    try:
        session = _STORE.get(doc_id)
        with _STORE.lock:
            session.last_used = time.time()
            yield session
            session.last_used = time.time()
    except ToolError:
        raise
    except CONVERTIBLE_ERRORS as exc:
        raise to_tool_error(exc) from exc


def _result(session: Session, **extra: Any) -> Dict[str, Any]:
    payload = session.describe()
    payload.update(extra)
    return payload


def _section(session: Session, index: int):
    sections = session.doc.sections
    resolved = index + len(sections) if index < 0 else index
    if resolved < 0 or resolved >= len(sections):
        raise ReportSessionError(
            f"section_index {index} 超出范围"
            f"（共 {len(sections)} 节，0~{len(sections) - 1}）"
        )
    return sections[resolved]


def _table(session: Session, index: int):
    tables = session.doc.tables
    if not tables:
        raise ReportSessionError("文档中还没有表格，请先调用 word_add_table。")
    resolved = index + len(tables) if index < 0 else index
    if resolved < 0 or resolved >= len(tables):
        raise ReportSessionError(
            f"table_index {index} 超出范围"
            f"（共 {len(tables)} 个表格，0~{len(tables) - 1}）"
        )
    return tables[resolved]


def _resolve_logo(
    session: Session,
    image_path: Optional[str],
    image_base64: Optional[str],
    label: str = "Logo 图片",
) -> Path:
    if image_path or image_base64:
        return materialize_image(
            session, image_path=image_path, image_base64=image_base64, label=label,
        )
    fallback = default_logo_path()
    if fallback is not None:
        return fallback
    raise ReportSessionError(
        f"缺少{label}：请提供 image_path 或 image_base64，"
        "或在服务端配置环境变量 WC_REPORT_MCP_DEFAULT_LOGO 指向一个默认 Logo。"
    )


# ====================================================================
# Layer 1 · 会话生命周期
# ====================================================================


def word_create_report(
    title: Optional[str] = None,
    author: Optional[str] = None,
    apply_default_margins: bool = True,
    language: Optional[str] = "zh-CN",
) -> Dict[str, Any]:
    """新建一份空白 Word 报告，返回后续所有操作都要带上的 doc_id。

    :param title: 文档标题，写入 docx 核心属性。
    :param author: 作者，写入 docx 核心属性。
    :param apply_default_margins: 是否套用公文默认页边距（上 3.7 / 下 3.5 / 左 2.8 / 右 2.6 cm）。
    :param language: 文档校对语言，默认 zh-CN；传 None 则不设置。
    """
    session = _STORE.create()
    formatter = session.formatter
    if apply_default_margins:
        formatter.setup_defaults()
    if language:
        formatter.set_document_language(language)
    if title:
        session.doc.core_properties.title = title
    if author:
        session.doc.core_properties.author = author
    return _result(session)


def word_open_report(path: str) -> Dict[str, Any]:
    """打开一份已存在的 .docx 继续编辑，返回新的 doc_id。

    :param path: 文件路径；相对路径先按当前目录解析，再按 WC_REPORT_MCP_OUTPUT_DIR 解析。
    """
    resolved = resolve_input_path(path, what="Word 文档")
    session = _STORE.create(source_path=resolved)
    return _result(session, opened=str(resolved))


def word_save_report(
    doc_id: str,
    path: Optional[str] = None,
    refresh_fields_on_open: bool = True,
) -> Dict[str, Any]:
    """把当前文档保存为 .docx 并返回落盘路径。

    :param doc_id: 会话 id。
    :param path: 保存路径；相对路径挂在 WC_REPORT_MCP_OUTPUT_DIR（默认 ~/wc-reports）下。
                 省略时复用上次保存路径，从未保存过则自动生成带时间戳的文件名。
    :param refresh_fields_on_open: 是否让 Word 打开文件时自动刷新目录与页码域。
    """
    with _op(doc_id) as session:
        if path:
            target = resolve_output_path(path)
        elif session.saved_path is not None:
            target = session.saved_path
        else:
            target = resolve_output_path(None)
        if refresh_fields_on_open:
            session.formatter._ensure_update_fields_on_open()
        session.doc.save(str(target))
        session.saved_path = target
        return _result(session, saved_to=str(target), size_bytes=target.stat().st_size)


def word_close_report(doc_id: str) -> Dict[str, Any]:
    """释放会话占用的内存。

    :param doc_id: 会话 id。关闭后该 doc_id 立即失效。
    """
    session = _STORE.get(doc_id)
    saved_path = session.saved_path
    source_path = session.source_path
    _STORE.close(doc_id)
    return {
        "closed": doc_id,
        "saved_path": str(saved_path) if saved_path else None,
        "source_path": str(source_path) if source_path else None,
    }


def word_describe_report(doc_id: str) -> Dict[str, Any]:
    """查看当前文档的结构摘要（段落数、节数、表格数及其形状）。

    :param doc_id: 会话 id。
    """
    with _op(doc_id) as session:
        payload = _result(session)
        # python-docx 的 Section 没有 paragraphs 属性，只能按节遍历正文内容
        payload["paragraphs_per_section"] = [
            sum(
                1
                for item in section.iter_inner_content()
                if isinstance(item, Paragraph)
            )
            for section in session.doc.sections
        ]
        payload["table_shapes"] = [
            {"rows": len(table.rows), "columns": len(table.columns)}
            for table in session.doc.tables
        ]
        return payload


# ====================================================================
# Layer 2 · 文档级设置
# ====================================================================


def word_set_default_font(
    doc_id: str,
    font_size: FontSize = 14,
    cn_font: str = DEFAULT_CN_FONT,
    en_font: str = DEFAULT_EN_FONT,
) -> Dict[str, Any]:
    """设置文档 Normal 样式的默认字号与中英文字体。

    :param doc_id: 会话 id。
    :param font_size: 字号，支持数字（pt）或中文字号字符串（三号/小四/小五…）。
    :param cn_font: 中文字体名。注意：字体必须在运行本服务端的机器上已安装。
    :param en_font: 西文字体名。
    """
    with _op(doc_id) as session:
        session.formatter.set_default_font(
            font_size=font_size, cn_font=cn_font, en_font=en_font,
        )
        return _result(session)


def word_set_page_margins(
    doc_id: str,
    top: float = 3.7,
    bottom: float = 3.5,
    left: float = 2.8,
    right: float = 2.6,
    gutter: float = 0,
    horizontal_alignment: str = "L",
) -> Dict[str, Any]:
    """设置全文页边距（单位 cm）。

    :param doc_id: 会话 id。
    :param top: 上边距，单位 cm，默认 3.7。
    :param bottom: 下边距，单位 cm，默认 3.5。
    :param left: 左边距，单位 cm，默认 2.8。
    :param right: 右边距，单位 cm，默认 2.6。
    :param gutter: 装订线宽度，单位 cm。
    :param horizontal_alignment: 页面水平对齐，支持 L/C/R 或 左对齐/居中/右对齐。
    """
    with _op(doc_id) as session:
        session.formatter.set_page_margins(
            top=top, bottom=bottom, left=left, right=right,
            gutter=gutter, horizontal_alignment=horizontal_alignment,
        )
        return _result(session)


def word_set_document_language(doc_id: str, lang: str = "zh-CN") -> Dict[str, Any]:
    """设置文档校对语言，避免 Word/OnlyOffice 按英文处理中文排版。

    :param doc_id: 会话 id。
    :param lang: BCP 47 语言标签，默认 zh-CN。
    """
    with _op(doc_id) as session:
        session.formatter.set_document_language(lang)
        return _result(session)


# ====================================================================
# Layer 2 · 段落与标题
# ====================================================================


def word_add_heading(
    doc_id: str,
    text: str,
    level: int = 1,
    font_name: str = "黑体",
    font_size: FontSize = "三号",
    bold: bool = False,
    indent: bool = False,
    alignment: Optional[str] = None,
    line_spacing: float = 1.5,
    space_before: Optional[float] = None,
    space_after: Optional[float] = None,
) -> Dict[str, Any]:
    """添加标题段落。

    :param doc_id: 会话 id。
    :param text: 标题文本。
    :param level: 标题层级 1~9，默认 1。层级会写入 Word 大纲级别，也是目录的收集依据。
    :param font_name: 中文字体名，默认黑体。
    :param font_size: 字号，默认三号。
    :param bold: 是否加粗。
    :param indent: 是否首行缩进 2 字符。
    :param alignment: 对齐方式，省略时沿用标题样式的居左。做居中的封面大标题时传 "居中"。
    :param line_spacing: 行距倍数，默认 1.5。
    :param space_before: 段前空白（pt）；省略时按字号的 0.5 行计算。
    :param space_after: 段后空白（pt）；省略时按字号的 0.5 行计算。
    """
    with _op(doc_id) as session:
        session.formatter.heading(
            text, level=level, font_name=font_name, font_size=font_size,
            bold=bold, indent=indent, line_spacing=line_spacing,
            alignment=alignment, space_before=space_before, space_after=space_after,
        )
        return _result(session)


def word_add_body(
    doc_id: str,
    text: str,
    font_name: str = "宋体",
    font_size: FontSize = "四号",
    indent: bool = True,
    bold: bool = False,
    alignment: str = "两端对齐",
    line_spacing: float = 1.5,
    space_before: float = 0,
    space_after: float = 0,
) -> Dict[str, Any]:
    """添加一个正文段落（默认宋体四号、首行缩进 2 字符、两端对齐、1.5 倍行距）。

    :param doc_id: 会话 id。
    :param text: 正文内容。
    :param font_name: 中文字体名。
    :param font_size: 字号，默认四号。
    :param indent: 是否首行缩进 2 字符。
    :param bold: 是否加粗。
    :param alignment: 对齐方式，支持 两端对齐/居中/居左/居右 或 left/center/right/justify。
    :param line_spacing: 行距倍数。
    :param space_before: 段前空白，单位 pt。
    :param space_after: 段后空白，单位 pt。
    """
    with _op(doc_id) as session:
        session.formatter.body(
            text, font_name=font_name, font_size=font_size, indent=indent,
            bold=bold, alignment=alignment, line_spacing=line_spacing,
            space_before=space_before, space_after=space_after,
        )
        return _result(session)


def word_add_body_list(
    doc_id: str,
    texts: List[str],
    font_name: str = "宋体",
    font_size: FontSize = "四号",
    indent: bool = True,
    bold: bool = False,
    alignment: str = "两端对齐",
    line_spacing: float = 1.5,
    space_before: float = 0,
    space_after: float = 0,
) -> Dict[str, Any]:
    """批量添加多个正文段落，共用同一套格式参数（长报告建议用它减少调用次数）。

    :param doc_id: 会话 id。
    :param texts: 正文段落列表，按顺序逐段写入。
    其余参数含义同 word_add_body。
    """
    with _op(doc_id) as session:
        for text in texts:
            session.formatter.body(
                text, font_name=font_name, font_size=font_size, indent=indent,
                bold=bold, alignment=alignment, line_spacing=line_spacing,
                space_before=space_before, space_after=space_after,
            )
        return _result(session, added=len(texts))


def word_add_blank_lines(doc_id: str, count: int = 1) -> Dict[str, Any]:
    """添加若干空行（封面、签发页等留白用）。

    :param doc_id: 会话 id。
    :param count: 空行数量，默认 1。
    """
    with _op(doc_id) as session:
        session.formatter.blank_lines(count)
        return _result(session, added=count)


def word_add_cover_text(
    doc_id: str,
    text: str,
    font_name: str = "方正小标宋简体",
    font_size: FontSize = 24,
    bold: bool = True,
    alignment: str = "居中",
    line_spacing: float = 2,
    space_before: Optional[float] = None,
    space_after: Optional[float] = None,
) -> Dict[str, Any]:
    """添加封面文本（默认居中加粗、2 倍行距）。

    :param doc_id: 会话 id。
    :param text: 封面文字。
    :param font_name: 中文字体名。
    :param font_size: 字号。
    :param bold: 是否加粗。
    :param alignment: 对齐方式。
    :param line_spacing: 行距倍数。
    :param space_before: 段前空白（pt）；省略时按字号的 1 行计算。
    :param space_after: 段后空白（pt）；省略时按字号的 1 行计算。
    """
    with _op(doc_id) as session:
        session.formatter.cover_text(
            text, font_name=font_name, font_size=font_size, bold=bold,
            alignment=alignment, line_spacing=line_spacing,
            space_before=space_before, space_after=space_after,
        )
        return _result(session)


def word_add_right_text(
    doc_id: str,
    text: str,
    font_name: str = DEFAULT_CN_FONT,
    font_size: FontSize = 14,
    indent: bool = True,
) -> Dict[str, Any]:
    """添加右对齐段落（公司名、落款等）。

    :param doc_id: 会话 id。
    :param text: 段落文本。
    :param font_name: 中文字体名。
    :param font_size: 字号。
    :param indent: 是否首行缩进 2 字符。
    """
    with _op(doc_id) as session:
        session.formatter.right_text(
            text, font_name=font_name, font_size=font_size, indent=indent,
        )
        return _result(session)


def word_add_date(
    doc_id: str,
    value: Optional[str] = None,
    style: str = "chinese",
    font_name: str = DEFAULT_CN_FONT,
    font_size: FontSize = 14,
    bold: bool = False,
    alignment: str = "居中",
) -> Dict[str, Any]:
    """添加日期文本。

    :param doc_id: 会话 id。
    :param value: 日期，支持 2026年07月16日 / 2026-07-16 / 2026/07/16 / 2026.07.16；省略时用当天。
    :param style: chinese = 二〇二六年七月十六日；year_month = 二〇二六年七月；
                  numeric = 2026年07月16日（右对齐）。
    :param font_name: 中文字体名。
    :param font_size: 字号。
    :param bold: 是否加粗。
    :param alignment: 对齐方式（numeric 样式固定右对齐，该参数不生效）。
    """
    with _op(doc_id) as session:
        formatter = session.formatter
        if style == "chinese":
            formatter.chinese_date(
                value, font_name=font_name, font_size=font_size,
                bold=bold, alignment=alignment,
            )
        elif style == "year_month":
            formatter.chinese_year_month(
                value, font_name=font_name, font_size=font_size,
                bold=bold, alignment=alignment,
            )
        elif style == "numeric":
            formatter.created_time(
                value, font_name=font_name, font_size=font_size, bold=bold,
            )
        else:
            raise ReportSessionError(
                f"未知的日期样式 {style!r}，可选 chinese / year_month / numeric。"
            )
        return _result(session, style=style)


def word_insert_image(
    doc_id: str,
    width_cm: Optional[float] = None,
    height_cm: Optional[float] = None,
    image_path: Optional[str] = None,
    image_base64: Optional[str] = None,
    alignment: str = "居左",
    floating: bool = False,
    y_offset_pt: float = 0,
) -> Dict[str, Any]:
    """插入一张图片到正文。

    ``width_cm`` 与 ``height_cm`` 都不传时按图片自身的 DPI 使用原始尺寸，
    因此从 PDF/扫描件裁切出来的图片可以直接插入，不需要先量一个宽度。

    :param doc_id: 会话 id。
    :param width_cm: 图片宽度，单位 cm；省略时按图片原始尺寸。
    :param height_cm: 图片高度，单位 cm；省略时按图片原始尺寸。
    :param image_path: 图片路径（与 image_base64 二选一）。
    :param image_base64: 图片的 base64，支持裸串或 data:image/png;base64,... 形式。
    :param alignment: 段落对齐方式。
    :param floating: 是否让图片浮于文字上方。印章要压在落款文字上时传 True，
                     从 PDF 裁出的图表、公式同理。
    :param y_offset_pt: 浮动图片的垂直偏移，单位 pt，正数上移、负数下移。
    """
    with _op(doc_id) as session:
        path = materialize_image(
            session, image_path=image_path, image_base64=image_base64, label="正文图片",
        )
        session.formatter.insert_img(
            str(path), width_cm, height=height_cm, alignment=alignment,
            floating=floating, y_offset_pt=y_offset_pt,
        )
        return _result(
            session,
            width_cm=width_cm,
            height_cm=height_cm,
            auto_sized=width_cm is None and height_cm is None,
            floating=floating,
        )


# ====================================================================
# Layer 2 · 表格
# ====================================================================


def word_add_table(
    doc_id: str,
    headers: List[str],
    rows: List[List[Any]],
    col_widths: Optional[List[float]] = None,
    font_size: FontSize = 12,
    font_name: str = DEFAULT_CN_FONT,
    header_bold: bool = True,
    alignment: str = "居中",
    border_color: str = "000000",
    border_size: int = 4,
    merges: Optional[List[Dict[str, int]]] = None,
    row_heights: Optional[List[Optional[float]]] = None,
    table_width_cm: Optional[float] = None,
    cell_font_names: Optional[List[Optional[List[Optional[str]]]]] = None,
) -> Dict[str, Any]:
    """在末尾追加一个带表头的表格。

    :param doc_id: 会话 id。
    :param headers: 表头文本列表，例如 ["序号", "项目", "金额（万元）"]。
    :param rows: 二维数据列表，每个子列表是一行。
    :param col_widths: 每列宽度列表，单位 cm，可选。
    :param font_size: 表格字号。
    :param font_name: 单元格中文字体名。
    :param header_bold: 表头是否加粗。
    :param alignment: 单元格对齐方式。
    :param border_color: 边框颜色，6 位十六进制 RGB。
    :param border_size: 边框粗细，单位 1/8 pt（4 = 0.5pt，8 = 1pt）。
    :param merges: 合并区域列表，例如
                   [{"row": 0, "col": 0, "rowspan": 1, "colspan": 4}]。
                   行列索引从 0 开始、按完整网格计数；合并区域的内容取左上角单元格的值。
    :param row_heights: 每行高度，单位 cm，可选；按「最小值」规则设置，内容超出会自动撑高。
    :param table_width_cm: 表格总宽度，单位 cm，可选。
    :param cell_font_names: 逐单元格中文字体名的二维列表（含表头行），可选；
                            未覆盖的位置回落到 font_name。
    """
    with _op(doc_id) as session:
        table = session.formatter.add_table(
            headers, rows, col_widths=col_widths, font_size=font_size,
            font_name=font_name, header_bold=header_bold, alignment=alignment,
            border_color=border_color, border_size=border_size,
            merges=merges, row_heights=row_heights,
            table_width_cm=table_width_cm, cell_font_names=cell_font_names,
        )
        return _result(
            session,
            table_index=len(session.doc.tables) - 1,
            rows=len(table.rows),
            columns=len(table.columns),
        )


def word_merge_cells(
    doc_id: str,
    merges: List[Dict[str, int]],
    table_index: int = -1,
) -> Dict[str, Any]:
    """合并指定表格里的单元格区域（对已存在的表格操作）。

    在 ``word_add_table`` 之后需要补合并时用它；新建表格时也可以直接在
    ``word_add_table`` 里传 ``merges``。

    :param doc_id: 会话 id。
    :param merges: 合并区域列表，例如
                   [{"row": 0, "col": 0, "rowspan": 1, "colspan": 3}]。
                   行列索引从 0 开始、按完整网格计数。
    :param table_index: 表格索引，默认 -1 表示文档中最后一个表格。
    """
    with _op(doc_id) as session:
        table = _table(session, table_index)
        WordFormatter.merge_cells(table, merges)
        return _result(session, table_index=table_index, merged=len(merges or []))


def word_format_cell(
    doc_id: str,
    row: int,
    col: int,
    text: str,
    table_index: int = -1,
    font_name: str = DEFAULT_CN_FONT,
    font_size: FontSize = 12,
    bold: bool = False,
    alignment: str = "居中",
    line_spacing: float = 1.25,
) -> Dict[str, Any]:
    """改写指定表格单元格的内容与格式。

    :param doc_id: 会话 id。
    :param row: 行索引，从 0 开始。
    :param col: 列索引，从 0 开始。
    :param text: 新的单元格文本（会清空原有内容）。
    :param table_index: 表格索引，默认 -1 表示文档中最后一个表格。
    :param font_name: 中文字体名。
    :param font_size: 字号。
    :param bold: 是否加粗。
    :param alignment: 对齐方式。
    :param line_spacing: 行距倍数。
    """
    with _op(doc_id) as session:
        table = _table(session, table_index)
        row_count, col_count = len(table.rows), len(table.columns)
        if row < 0 or row >= row_count or col < 0 or col >= col_count:
            raise ReportSessionError(
                f"单元格 ({row}, {col}) 超出范围（该表共 {row_count} 行 × {col_count} 列）。"
            )
        WordFormatter.set_cell(
            table.cell(row, col), text, font_name=font_name, font_size=font_size,
            bold=bold, alignment=alignment, line_spacing=line_spacing,
        )
        return _result(session, table_index=table_index, row=row, col=col)


def word_set_table_borders(
    doc_id: str,
    table_index: int = -1,
    color: str = "000000",
    size: int = 4,
) -> Dict[str, Any]:
    """为指定表格添加四边及内部横竖边框。

    :param doc_id: 会话 id。
    :param table_index: 表格索引，默认 -1 表示最后一个表格。
    :param color: 边框颜色，6 位十六进制 RGB。
    :param size: 边框粗细，单位 1/8 pt。
    """
    with _op(doc_id) as session:
        WordFormatter.set_table_borders(
            _table(session, table_index), color=color, size=size,
        )
        return _result(session, table_index=table_index)


# ====================================================================
# Layer 2 · 页眉页脚
# ====================================================================


def word_set_header(
    doc_id: str,
    text: str,
    section_index: int = -1,
    variant: str = "primary",
    alignment: str = "居中",
    font_name: str = DEFAULT_HEADER_FOOTER_FONT,
    font_size: FontSize = DEFAULT_HEADER_FOOTER_SIZE,
    bottom_border: bool = True,
    line_length: Optional[float] = None,
    line_alignment: str = "居左",
) -> Dict[str, Any]:
    """设置某节的页眉文字（默认楷体小五号，页眉默认带下划线）。

    :param doc_id: 会话 id。
    :param text: 页眉文本。
    :param section_index: 节索引，默认 -1 表示最后一节。
    :param variant: 页眉变体：primary（默认/奇数页）、first（首页）、even（偶数页）。
                    使用 first 前需先用 word_set_different_first_page 打开本节开关；
                    使用 even 前需先用 word_set_different_odd_even 打开全文开关。
    :param alignment: 文本对齐方式。
    :param font_name: 中文字体名。
    :param font_size: 字号。
    :param bottom_border: 是否给页眉段落加下横线。
    :param line_length: 横线长度，单位 cm；省略时使用页面可用宽度。
    :param line_alignment: 横线对齐方式。
    """
    with _op(doc_id) as session:
        section = _section(session, section_index)
        WordFormatter.set_header(
            section, text, variant=variant, alignment=alignment, font_name=font_name,
            font_size=font_size, bottom_border=bottom_border,
            line_length=line_length, line_alignment=line_alignment,
        )
        return _result(session, section_index=section_index, variant=variant)


def word_set_footer(
    doc_id: str,
    text: str,
    section_index: int = -1,
    variant: str = "primary",
    alignment: str = "居中",
    font_name: str = DEFAULT_HEADER_FOOTER_FONT,
    font_size: FontSize = DEFAULT_HEADER_FOOTER_SIZE,
    top_border: bool = True,
    line_length: Optional[float] = None,
    line_alignment: str = "居左",
) -> Dict[str, Any]:
    """设置某节的页脚文字（默认楷体小五号，页脚默认带上横线）。

    :param doc_id: 会话 id。
    :param text: 页脚文本。已有的页码域会被保留。
    :param section_index: 节索引，默认 -1 表示最后一节。
    :param variant: 页脚变体：primary（默认/奇数页）、first（首页）、even（偶数页）。
    :param alignment: 文本对齐方式。
    :param font_name: 中文字体名。
    :param font_size: 字号。
    :param top_border: 是否给页脚段落加上横线。
    :param line_length: 横线长度，单位 cm。
    :param line_alignment: 横线对齐方式。
    """
    with _op(doc_id) as session:
        section = _section(session, section_index)
        WordFormatter.set_footer(
            section, text, variant=variant, alignment=alignment, font_name=font_name,
            font_size=font_size, top_border=top_border,
            line_length=line_length, line_alignment=line_alignment,
        )
        return _result(session, section_index=section_index, variant=variant)


def word_set_header_parts(
    doc_id: str,
    left: str = "",
    center: str = "",
    right: str = "",
    section_index: int = -1,
    variant: str = "primary",
    font_name: str = DEFAULT_HEADER_FOOTER_FONT,
    font_size: FontSize = DEFAULT_HEADER_FOOTER_SIZE,
    bottom_border: bool = True,
    line_length: Optional[float] = None,
    line_alignment: str = "居左",
) -> Dict[str, Any]:
    """在一行内分左、中、右三段设置页眉（中文报告最常见的页眉形式）。

    用制表位实现：左段贴版心左边界、中段居中、右段贴版心右边界。

    :param doc_id: 会话 id。
    :param left: 左段文本，如公司名。
    :param center: 中段文本，如报告名。
    :param right: 右段文本，如页码或密级。
    :param section_index: 节索引，默认 -1 表示最后一节。
    :param variant: primary / first / even。
    :param font_name: 中文字体名。
    :param font_size: 字号。
    :param bottom_border: 是否给页眉段落加下横线。
    :param line_length: 横线长度，单位 cm。
    :param line_alignment: 横线对齐方式。
    """
    with _op(doc_id) as session:
        WordFormatter.set_header_parts(
            _section(session, section_index),
            left=left, center=center, right=right, variant=variant,
            font_name=font_name, font_size=font_size, bottom_border=bottom_border,
            line_length=line_length, line_alignment=line_alignment,
        )
        return _result(session, section_index=section_index, variant=variant)


def word_set_footer_parts(
    doc_id: str,
    left: str = "",
    center: str = "",
    right: str = "",
    section_index: int = -1,
    variant: str = "primary",
    font_name: str = DEFAULT_HEADER_FOOTER_FONT,
    font_size: FontSize = DEFAULT_HEADER_FOOTER_SIZE,
    top_border: bool = True,
    line_length: Optional[float] = None,
    line_alignment: str = "居左",
) -> Dict[str, Any]:
    """在一行内分左、中、右三段设置页脚；参数含义同 word_set_header_parts。

    :param top_border: 是否给页脚段落加上横线。
    """
    with _op(doc_id) as session:
        WordFormatter.set_footer_parts(
            _section(session, section_index),
            left=left, center=center, right=right, variant=variant,
            font_name=font_name, font_size=font_size, top_border=top_border,
            line_length=line_length, line_alignment=line_alignment,
        )
        return _result(session, section_index=section_index, variant=variant)


def word_set_different_first_page(
    doc_id: str,
    enabled: bool = True,
    section_index: int = -1,
) -> Dict[str, Any]:
    """设置某节首页是否使用独立的页眉页脚（w:titlePg）。

    打开后即可用 variant="first" 单独设置首页页眉页脚，常见于「首页不显示页码」。

    :param doc_id: 会话 id。
    :param enabled: 是否启用独立首页页眉页脚。
    :param section_index: 节索引，默认 -1 表示最后一节。
    """
    with _op(doc_id) as session:
        WordFormatter.set_different_first_page(
            _section(session, section_index), enabled,
        )
        return _result(session, section_index=section_index, enabled=enabled)


def word_set_different_odd_even(
    doc_id: str,
    enabled: bool = True,
) -> Dict[str, Any]:
    """设置全文奇偶页是否使用不同的页眉页脚（文档级 w:evenAndOddHeaders）。

    打开后 variant="even" 的页眉页脚才会生效。注意这是**文档级**开关，
    一旦打开，未单独设置偶页页眉的节会显示为空页眉。

    :param doc_id: 会话 id。
    :param enabled: 是否启用奇偶页不同。
    """
    with _op(doc_id) as session:
        session.formatter.set_different_odd_even(enabled)
        return _result(session, enabled=enabled)


def word_clear_header_footer(
    doc_id: str,
    target: str = "header",
    section_index: int = -1,
    variant: str = "primary",
) -> Dict[str, Any]:
    """清空某节的页眉或页脚。

    :param doc_id: 会话 id。
    :param target: header / footer / both。
    :param section_index: 节索引，默认 -1 表示最后一节。
    :param variant: primary / first / even。
    """
    with _op(doc_id) as session:
        if target not in ("header", "footer", "both"):
            raise ReportSessionError(
                f"target 只能是 header/footer/both，收到 {target!r}。"
            )
        section = _section(session, section_index)
        if target in ("header", "both"):
            WordFormatter.clear_header(section, variant=variant)
        if target in ("footer", "both"):
            WordFormatter.clear_footer(section, variant=variant)
        return _result(session, section_index=section_index,
                       variant=variant, cleared=target)


def word_set_header_image(
    doc_id: str,
    target: str = "header",
    section_index: int = -1,
    variant: str = "primary",
    width_cm: Optional[float] = None,
    height_cm: Optional[float] = None,
    image_path: Optional[str] = None,
    image_base64: Optional[str] = None,
    alignment: str = "居中",
    floating: bool = True,
    bottom_border: bool = True,
    y_offset_pt: float = 0,
) -> Dict[str, Any]:
    """设置页眉（或页脚）的图片；已有文字会保留，已有图片会被替换。

    :param doc_id: 会话 id。
    :param target: header 或 footer。
    :param section_index: 节索引，默认 -1 表示最后一节。
    :param variant: primary / first / even。
    :param width_cm: 图片宽度，单位 cm（与 height_cm 至少传一个）。
    :param height_cm: 图片高度，单位 cm。
    :param image_path: 图片路径（与 image_base64 二选一）。
    :param image_base64: 图片的 base64。
    :param alignment: 对齐方式。
    :param floating: 是否浮于文字上方（Word 兼容性更好）。
    :param bottom_border: 是否加下横线。
    :param y_offset_pt: 垂直偏移，单位 pt，正数上移、负数下移。
    """
    with _op(doc_id) as session:
        if target not in ("header", "footer"):
            raise ReportSessionError(
                f"target 只能是 header/footer，收到 {target!r}。"
            )
        section = _section(session, section_index)
        path = materialize_image(
            session, image_path=image_path, image_base64=image_base64,
            label="页眉页脚图片",
        )
        setter = (
            WordFormatter.set_header_image
            if target == "header"
            else WordFormatter.set_footer_image
        )
        setter(
            section, str(path), variant=variant, width=width_cm, height=height_cm,
            alignment=alignment, floating=floating, bottom_border=bottom_border,
            y_offset_pt=y_offset_pt,
        )
        return _result(session, section_index=section_index,
                       variant=variant, target=target)


# ====================================================================
# Layer 2 · 节与页码
# ====================================================================


def word_insert_section(
    doc_id: str,
    add_page_number: bool = False,
    restart_page_number: bool = False,
    start_page_number: int = 1,
    inherit_header: bool = True,
    inherit_footer: bool = True,
    prefix: str = "",
    suffix: str = "",
    alignment: str = "居中",
    number_format: Optional[str] = None,
    orientation: Optional[str] = None,
    page_width_cm: Optional[float] = None,
    page_height_cm: Optional[float] = None,
    margin_top_cm: Optional[float] = None,
    margin_bottom_cm: Optional[float] = None,
    margin_left_cm: Optional[float] = None,
    margin_right_cm: Optional[float] = None,
) -> Dict[str, Any]:
    """插入一个新节（默认从新页开始）。

    横向插页与纵向正文混排时，用 orientation / margin_* 直接给本节设定纸张与边距，
    不必先插节再单独调用 word_set_section_page。

    :param doc_id: 会话 id。
    :param add_page_number: 是否在本节设置页码。
    :param restart_page_number: 是否从 start_page_number 重新编号。
    :param start_page_number: 重新编号的起始页码。
    :param inherit_header: 是否继承上一节页眉。
    :param inherit_footer: 是否继承上一节页脚。
    :param prefix: 页码前缀，如 "第 "。
    :param suffix: 页码后缀，如 " 页"。
    :param alignment: 页码对齐方式。
    :param number_format: 页码编号格式，如 "upperRoman"/"大写罗马"/"chineseCounting"。
    :param orientation: 本节纸张方向，"纵向"/"横向" 或 portrait/landscape。
    :param page_width_cm/page_height_cm: 本节纸张宽高，单位 cm，可选。
    :param margin_top_cm/margin_bottom_cm/margin_left_cm/margin_right_cm: 本节页边距，单位 cm，可选。
    """
    with _op(doc_id) as session:
        session.formatter.insert_section(
            add_page_number=add_page_number,
            restart_page_number=restart_page_number,
            start_page_number=start_page_number,
            inherit_header=inherit_header, inherit_footer=inherit_footer,
            prefix=prefix, suffix=suffix, alignment=alignment,
            number_format=number_format, orientation=orientation,
            page_width=page_width_cm, page_height=page_height_cm,
            top=margin_top_cm, bottom=margin_bottom_cm,
            left=margin_left_cm, right=margin_right_cm,
        )
        return _result(session, section_index=len(session.doc.sections) - 1)


def word_set_section_page(
    doc_id: str,
    section_index: int = -1,
    orientation: Optional[str] = None,
    page_width_cm: Optional[float] = None,
    page_height_cm: Optional[float] = None,
    margin_top_cm: Optional[float] = None,
    margin_bottom_cm: Optional[float] = None,
    margin_left_cm: Optional[float] = None,
    margin_right_cm: Optional[float] = None,
    gutter_cm: Optional[float] = None,
    horizontal_alignment: Optional[str] = None,
) -> Dict[str, Any]:
    """单独设置某一节的纸张方向、尺寸与页边距。

    word_set_page_margins 会把边距应用到全文所有节；本节级设置有需要时用这个工具。

    :param doc_id: 会话 id。
    :param section_index: 节索引，默认 -1 表示最后一节。
    :param orientation: "纵向"/"横向" 或 portrait/landscape；与当前方向不同时自动交换宽高。
    :param page_width_cm/page_height_cm: 纸张宽高，单位 cm，可选。
    :param margin_top_cm/margin_bottom_cm/margin_left_cm/margin_right_cm: 页边距，单位 cm，可选。
    :param gutter_cm: 装订线宽度，单位 cm，可选。
    :param horizontal_alignment: 页面水平对齐，支持 L/C/R 或 左对齐/居中/右对齐。
    """
    with _op(doc_id) as session:
        section = _section(session, section_index)
        session.formatter.set_section_page(
            section, orientation=orientation,
            page_width=page_width_cm, page_height=page_height_cm,
            top=margin_top_cm, bottom=margin_bottom_cm,
            left=margin_left_cm, right=margin_right_cm,
            gutter=gutter_cm, horizontal_alignment=horizontal_alignment,
        )
        return _result(
            session,
            section_index=section_index,
            page_width_cm=round(section.page_width.cm, 2),
            page_height_cm=round(section.page_height.cm, 2),
        )


def word_set_page_numbers(
    doc_id: str,
    start_section_index: int = 0,
    start: int = 1,
    prefix: str = "",
    suffix: str = "",
    alignment: str = "居中",
    font_name: Optional[str] = None,
    font_size: Optional[FontSize] = None,
    number_format: Optional[str] = None,
) -> Dict[str, Any]:
    """从指定节开始编页码，之前的节不带页码（跨节页码的主入口）。

    典型用法：封面/目录在前置节（无页码），正文从第 1 节开始编号为"第 1 页"。

    :param doc_id: 会话 id。
    :param start_section_index: 从 0 开始的节索引，页码从此节开始。
    :param start: 起始页码。
    :param prefix: 页码前缀，如 "第 "。
    :param suffix: 页码后缀，如 " 页"。
    :param alignment: 页码对齐方式。
    :param font_name: 页码字体名，省略时用默认中文字体。
    :param font_size: 页码字号，省略时用默认正文字号。
    :param number_format: 页码编号格式，如 "upperRoman"/"大写罗马"、
                          "chineseCounting"/"中文数字"；省略时用阿拉伯数字。
    """
    with _op(doc_id) as session:
        session.formatter.set_page_number_from_section(
            start_section_index, start=start, prefix=prefix, suffix=suffix,
            alignment=alignment, font_name=font_name, font_size=font_size,
            number_format=number_format,
        )
        return _result(session, start_section_index=start_section_index)


def word_restart_page_numbering(
    doc_id: str,
    section_index: int = -1,
    start: int = 1,
    add_footer_number: bool = True,
    prefix: str = "",
    suffix: str = "",
    alignment: str = "居中",
    number_format: Optional[str] = None,
) -> Dict[str, Any]:
    """重启某节的页码编号。

    :param doc_id: 会话 id。
    :param section_index: 节索引，默认 -1 表示最后一节。
    :param start: 起始页码。
    :param add_footer_number: 是否同时在页脚渲染页码。
    :param prefix: 页码前缀。
    :param suffix: 页码后缀。
    :param alignment: 页码对齐方式。
    :param number_format: 页码编号格式，如 "upperRoman"/"大写罗马"；省略时沿用原格式。
    """
    with _op(doc_id) as session:
        session.formatter.restart_page_numbering(
            _section(session, section_index), start=start,
            add_footer_number=add_footer_number, prefix=prefix, suffix=suffix,
            alignment=alignment, number_format=number_format,
        )
        return _result(session, section_index=section_index)


def word_set_page_number_start(
    doc_id: str,
    start: int = 1,
    section_index: int = -1,
    number_format: Optional[str] = None,
) -> Dict[str, Any]:
    """设置某节页码的起始值（可选同时设置编号格式），不添加页脚页码域。

    :param doc_id: 会话 id。
    :param start: 起始页码。
    :param section_index: 节索引，默认 -1 表示最后一节。
    :param number_format: 页码编号格式，如 "upperRoman"/"大写罗马"/"chineseCounting"。
    """
    with _op(doc_id) as session:
        WordFormatter.set_page_number_start(
            _section(session, section_index), start, number_format=number_format,
        )
        return _result(session, section_index=section_index, start=start)


def word_set_page_number_format(
    doc_id: str,
    number_format: str,
    section_index: int = -1,
    start: Optional[int] = None,
) -> Dict[str, Any]:
    """只设置某节的页码编号格式（不改动已有的页码域）。

    封面用罗马数字、正文用阿拉伯数字时，配合 word_insert_section 分节后分别设置。

    :param doc_id: 会话 id。
    :param number_format: 编号格式，支持 decimal / upperRoman / lowerRoman /
                          chineseCounting / decimalEnclosedCircle 等，也接受
                          大写罗马 / 小写罗马 / 中文数字 / 带圈数字。
    :param section_index: 节索引，默认 -1 表示最后一节。
    :param start: 同时指定起始页码，可选。
    """
    with _op(doc_id) as session:
        WordFormatter.set_page_number_format(
            _section(session, section_index), number_format, start=start,
        )
        return _result(session, section_index=section_index,
                       number_format=number_format, start=start)


# ====================================================================
# Layer 2 · 目录与样式
# ====================================================================


def word_add_toc(
    doc_id: str,
    title: Optional[str] = "目录",
    levels_min: int = 1,
    levels_max: int = 3,
    title_font_name: str = "黑体",
    title_font_size: FontSize = 16,
    title_bold: bool = True,
    title_alignment: str = "居中",
    toc_level_styles: Optional[Dict[str, Any]] = None,
    use_hyperlinks: bool = True,
) -> Dict[str, Any]:
    """插入 Word 目录域（TOC）。

    注意：目录内容需要 Word/OnlyOffice 打开后刷新域才会显示页码。本工具会写入
    updateFields 标记，让 Word 打开时自动刷新；LibreOffice/Google Docs 不保证。

    :param doc_id: 会话 id。
    :param title: 目录标题；传 None 则不生成标题。
    :param levels_min: 收集的最小标题层级。
    :param levels_max: 收集的最大标题层级。
    :param title_font_name: 目录标题字体。
    :param title_font_size: 目录标题字号。
    :param title_bold: 目录标题是否加粗。
    :param title_alignment: 目录标题对齐方式。
    :param toc_level_styles: 各级目录样式，格式 {"1": {"font_name": "黑体", "font_size": 14}}。
    :param use_hyperlinks: 是否生成超链接。
    """
    with _op(doc_id) as session:
        styles = None
        if toc_level_styles:
            styles = {int(level): options for level, options in toc_level_styles.items()}
        session.formatter.add_toc(
            title=title, levels=(levels_min, levels_max),
            title_font_name=title_font_name, title_font_size=title_font_size,
            title_bold=title_bold, title_alignment=title_alignment,
            toc_level_styles=styles, use_hyperlinks=use_hyperlinks,
        )
        return _result(session, levels=[levels_min, levels_max])


def word_set_toc_level_style(
    doc_id: str,
    level: int,
    font_name: Optional[str] = None,
    font_size: Optional[FontSize] = None,
    bold: Optional[bool] = None,
    line_spacing: Optional[float] = None,
    space_before: Optional[float] = None,
    space_after: Optional[float] = None,
    first_line_indent_chars: Optional[float] = None,
    left_indent: Optional[float] = None,
    alignment: Optional[str] = None,
) -> Dict[str, Any]:
    """设置某一级 TOC 目录样式。

    :param doc_id: 会话 id。
    :param level: 目录层级。
    :param font_name: 字体名，可选。
    :param font_size: 字号，可选。
    :param bold: 是否加粗，可选。
    :param line_spacing: 行距倍数，可选。
    :param space_before: 段前空白（pt），可选。
    :param space_after: 段后空白（pt），可选。
    :param first_line_indent_chars: 首行缩进字符数，可选。
    :param left_indent: 左缩进（pt），可选。
    :param alignment: 对齐方式，可选。
    """
    with _op(doc_id) as session:
        session.formatter.set_toc_level_style(
            level, font_name=font_name, font_size=font_size, bold=bold,
            line_spacing=line_spacing, space_before=space_before,
            space_after=space_after,
            first_line_indent_chars=first_line_indent_chars,
            left_indent=left_indent, alignment=alignment,
        )
        return _result(session, level=level)


def word_set_paragraph_style(
    doc_id: str,
    style_name: str,
    base_style_name: str = "Normal",
    font_name: Optional[str] = None,
    font_size: Optional[FontSize] = None,
    bold: Optional[bool] = None,
    line_spacing: Optional[float] = None,
    space_before: Optional[float] = None,
    space_after: Optional[float] = None,
    first_line_indent_chars: Optional[float] = None,
    left_indent: Optional[float] = None,
    alignment: Optional[str] = None,
) -> Dict[str, Any]:
    """创建或更新一个自定义段落样式（配合 word_add_custom_heading 使用）。

    :param doc_id: 会话 id。
    :param style_name: 样式名称。
    :param base_style_name: 基础样式名，默认 Normal。
    其余参数均为可选，含义同 word_set_toc_level_style。
    """
    with _op(doc_id) as session:
        session.formatter.set_paragraph_style(
            style_name, base_style_name=base_style_name, font_name=font_name,
            font_size=font_size, bold=bold, line_spacing=line_spacing,
            space_before=space_before, space_after=space_after,
            first_line_indent_chars=first_line_indent_chars,
            left_indent=left_indent, alignment=alignment,
        )
        return _result(session, style_name=style_name)


def word_add_custom_heading(
    doc_id: str,
    text: str,
    style_name: str,
    level: Optional[int] = None,
    font_name: Optional[str] = None,
    font_size: Optional[FontSize] = None,
) -> Dict[str, Any]:
    """使用自定义样式添加标题。

    :param doc_id: 会话 id。
    :param text: 标题文本。
    :param style_name: 之前用 word_set_paragraph_style 创建（或文档自带）的样式名。
    :param level: 大纲级别，传入后该标题可被目录收集。
    :param font_name: 覆盖字体名，可选。
    :param font_size: 覆盖字号，可选。
    """
    with _op(doc_id) as session:
        session.formatter.add_custom_heading(
            text, style_name=style_name, level=level,
            font_name=font_name, font_size=font_size,
        )
        return _result(session, style_name=style_name)


# ====================================================================
# Layer 2 · 模板化组合
# ====================================================================


def word_add_cover_page(
    doc_id: str,
    project_name: str = "默认工程名称",
    entrusting_unit: str = "默认委托单位",
    compiling_unit: str = "默认编制单位",
    report_title: str = "结算审核报告",
    date: Optional[str] = None,
    logo_width_cm: float = 4,
    info_blank_lines: int = 11,
    image_path: Optional[str] = None,
    image_base64: Optional[str] = None,
) -> Dict[str, Any]:
    """插入结算审核报告封面（Logo + 工程名 + 报告标题 + 落款表格 + 中文年月）。

    :param doc_id: 会话 id。
    :param project_name: 工程名称。
    :param entrusting_unit: 委托单位名称。
    :param compiling_unit: 编制单位名称。
    :param report_title: 报告标题。
    :param date: 封面日期，省略时用当天。
    :param logo_width_cm: Logo 宽度，单位 cm。
    :param info_blank_lines: 报告标题与落款表格之间的空行数。
    :param image_path: Logo 路径；省略时回退到环境变量 WC_REPORT_MCP_DEFAULT_LOGO。
    :param image_base64: Logo 的 base64（与 image_path 二选一）。
    """
    with _op(doc_id) as session:
        logo = _resolve_logo(session, image_path, image_base64)
        session.formatter.fengmian_jiesuan(
            logo, project_name, entrusting_unit, compiling_unit=compiling_unit,
            report_title=report_title, date=date, logo_width=logo_width_cm,
            info_blank_lines=info_blank_lines,
        )
        return _result(session)


def word_add_signature_page(
    doc_id: str,
    project_name: str = "默认工程名称",
    entrusting_unit: str = "默认委托单位",
    personnel: Optional[Dict[str, Any]] = None,
    participants: Optional[List[Dict[str, Any]]] = None,
    compiling_unit: str = "默认编制单位",
    report_title: str = "结算审核报告",
    footer_text: str = "公司从业方针：客观、公正、严谨、专业",
    logo_width_cm: float = 2,
    logo_y_offset_pt: float = 0,
    image_path: Optional[str] = None,
    image_base64: Optional[str] = None,
) -> Dict[str, Any]:
    """新建一节并插入结算审核报告签发页。

    :param doc_id: 会话 id。
    :param project_name: 工程名称。
    :param entrusting_unit: 委托单位名称。
    :param personnel: 各签发岗位信息，键为 公司签发/部门核准/部门审核/小组初审/项目负责人，
                      值可含 姓名、部门、职务、职称，项目负责人还可含 联系电话。
    :param participants: 项目参与者列表，每项可含 姓名、职称。
    :param compiling_unit: 报告编制单位。
    :param report_title: 页眉中的报告标题。
    :param footer_text: 页脚文字。
    :param logo_width_cm: 页眉 Logo 宽度，单位 cm。
    :param logo_y_offset_pt: 页眉 Logo 垂直偏移，单位 pt，正数上移、负数下移。
    :param image_path: 页眉 Logo 路径；省略时回退到环境变量 WC_REPORT_MCP_DEFAULT_LOGO。
    :param image_base64: 页眉 Logo 的 base64。
    """
    with _op(doc_id) as session:
        logo = _resolve_logo(session, image_path, image_base64)
        session.formatter.qianfaye(
            logo, project_name, entrusting_unit, personnel=personnel,
            participants=participants, compiling_unit=compiling_unit,
            report_title=report_title, footer_text=footer_text,
            logo_width=logo_width_cm, logo_y_offset_pt=logo_y_offset_pt,
        )
        return _result(session, section_index=len(session.doc.sections) - 1)


# ====================================================================
# Layer 2 · 无状态工具
# ====================================================================


def word_rmb_upper(value: Union[float, int, str]) -> Dict[str, Any]:
    """把数字转换为人民币金额中文大写（不需要 doc_id）。

    :param value: 数字或数字字符串，支持千分位逗号，例如 123456.78 或 "1,234.5"。
    """
    try:
        chinese_upper = number_to_chinese_upper(value)
    except CONVERTIBLE_ERRORS as exc:
        raise to_tool_error(exc) from exc
    return {"value": str(value), "chinese_upper": chinese_upper}


# ====================================================================
# Layer 3 · 兜底工具（保证对 WordFormatter 的全量覆盖）
# ====================================================================

_INDEXABLE_PARAMS = {"section": "sections", "table": "tables"}


def _resolve_object_params(session: Session, params: Dict[str, Any]) -> Dict[str, Any]:
    """把 section / table 的整数索引换成真实对象，供静态方法使用。"""
    resolved = dict(params)
    for key, collection in _INDEXABLE_PARAMS.items():
        value = resolved.get(key)
        if isinstance(value, bool) or not isinstance(value, int):
            continue
        items = getattr(session.doc, collection)
        if not items:
            raise ReportSessionError(f"文档中还没有{collection}，无法解析 {key}={value}。")
        index = value + len(items) if value < 0 else value
        if index < 0 or index >= len(items):
            raise ReportSessionError(
                f"{key} 索引 {value} 超出范围（共 {len(items)} 个，0~{len(items) - 1}）。"
            )
        resolved[key] = items[index]
    return resolved


def _describe_result(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return type(value).__name__


def word_describe_api() -> str:
    """列出 WordFormatter 的全部公开方法、签名与说明（配合 word_call 使用）。"""
    return formatter_api_text()


def word_call(
    doc_id: str,
    method: str,
    params: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """直接调用 WordFormatter 上的任意公开方法（兜底能力，覆盖所有长尾参数）。

    先用 word_describe_api 查看可用方法与签名，再把关键字参数放进 params。
    section / table 这类对象参数可以直接传整数索引，-1 表示最后一个。

    :param doc_id: 会话 id。
    :param method: WordFormatter 的方法名，例如 "heading"、"set_toc_level_style"。
    :param params: 该方法的关字参数字典。
    """
    with _op(doc_id) as session:
        if method.startswith("_"):
            raise ReportSessionError(f"不允许调用私有方法: {method!r}。")
        target = getattr(session.formatter, method, None)
        if target is None or not callable(target):
            raise ReportSessionError(
                f"WordFormatter 没有公开方法 {method!r}。请先调用 word_describe_api 查看可用方法。"
            )
        resolved = _resolve_object_params(session, params or {})
        result = target(**resolved)
        return _result(session, method=method, returned=_describe_result(result))


TOOLS = (
    # Layer 1
    word_create_report,
    word_open_report,
    word_save_report,
    word_close_report,
    word_describe_report,
    # Layer 2 · 文档级设置
    word_set_default_font,
    word_set_page_margins,
    word_set_document_language,
    # Layer 2 · 段落与标题
    word_add_heading,
    word_add_body,
    word_add_body_list,
    word_add_blank_lines,
    word_add_cover_text,
    word_add_right_text,
    word_add_date,
    word_insert_image,
    # Layer 2 · 表格
    word_add_table,
    word_merge_cells,
    word_format_cell,
    word_set_table_borders,
    # Layer 2 · 页眉页脚
    word_set_header,
    word_set_footer,
    word_set_header_parts,
    word_set_footer_parts,
    word_set_different_first_page,
    word_set_different_odd_even,
    word_clear_header_footer,
    word_set_header_image,
    # Layer 2 · 节与页码
    word_insert_section,
    word_set_section_page,
    word_set_page_numbers,
    word_restart_page_numbering,
    word_set_page_number_start,
    word_set_page_number_format,
    # Layer 2 · 目录与样式
    word_add_toc,
    word_set_toc_level_style,
    word_set_paragraph_style,
    word_add_custom_heading,
    # Layer 2 · 模板化组合
    word_add_cover_page,
    word_add_signature_page,
    # Layer 2 · 无状态
    word_rmb_upper,
    # Layer 3 · 兜底
    word_describe_api,
    word_call,
)


def register(mcp) -> None:
    """把所有工具注册到 FastMCP 实例。"""
    for tool in TOOLS:
        mcp.tool()(tool)
