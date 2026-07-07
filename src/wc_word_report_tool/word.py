"""wc_word_report_tool.word
=========================
基于 python-docx 的中文 Word 报告格式化工具。

设计目标
--------
1. **覆盖完整**：表格、页眉页脚、目录、页码、封面等报告常见需求一站式覆盖。
2. **命名一致**：所有公开方法使用 snake_case，参数命名清晰，单位明确。
3. **默认合理**：默认值贴近中国公文标准（仿宋_GB2312、三号 14pt、1.5 倍行距）。
4. **严格校验**：未知参数抛 ValueError 而非静默降级，便于调试。
5. **向后兼容**：v0.1.x 的旧 API（Heading_1、Normal_doc、fengmian_doc1 等）保留为
   别名，已有代码无需修改即可升级。

版本：v0.2.0
作者：willcha
"""

from __future__ import annotations

import datetime as _datetime
import warnings

from docx.enum.style import WD_STYLE_TYPE
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX, WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


# ====================================================================
# 字号常量（单位：pt）
# ====================================================================
# 中文字号 -> pt 对照，便于按号数指定字号
FONT_SIZE_MAP = {
    "初号": 42, "小初": 36,
    "一号": 26, "小一": 24,
    "二号": 22, "小二": 18,
    "三号": 16, "小三": 15,
    "四号": 14, "小四": 12,
    "五号": 10.5, "小五": 9,
    "六号": 7.5, "小六": 6.5,
    "七号": 5.5, "八号": 5,
}

# 公文/报告常用默认值
DEFAULT_CN_FONT = "仿宋_GB2312"
DEFAULT_EN_FONT = "Times New Roman"
DEFAULT_BODY_SIZE = 14          # 三号
DEFAULT_HEADER_FOOTER_FONT = "楷体"
DEFAULT_HEADER_FOOTER_SIZE = 9  # 小五号


def _to_pt(size) -> float:
    """将字号参数统一转为 pt。
    - int/float：视为 pt 数值
    - str："三号"/"小五"/"16" 等中文号数或字符串数字
    """
    if isinstance(size, (int, float)):
        return float(size)
    if isinstance(size, str):
        key = size.strip()
        if key in FONT_SIZE_MAP:
            return FONT_SIZE_MAP[key]
        try:
            return float(key)
        except ValueError:
            raise ValueError(
                f"无法识别的字号: {size!r}。支持数字（pt）或中文字号"
                f"（如 '三号'/'小五'），可选值: {list(FONT_SIZE_MAP.keys())}"
            )
    raise TypeError(f"字号参数类型不支持: {type(size).__name__}")


class WordFormatter:
    """Word 报告格式化工具。

    所有方法均为 staticmethod，可直接 ``WordFormatter.method(doc, ...)`` 调用，
    无需实例化。设计上无状态，便于在不同文档间复用。

    命名约定（v0.2.0+）
    ------------------
    - snake_case 全小写：``heading1`` / ``body`` / ``set_header``
    - 设置类前缀 ``set_``：``set_cell`` / ``set_header`` / ``set_table_borders``
    - 添加类前缀 ``add_``：``add_table`` / ``add_toc`` / ``add_page_number``
    - 单位：长度 cm，字号 pt 或中文字号（"三号"），页边距 cm

    向后兼容
    --------
    v0.1.x 的旧方法名（Heading_1 / Normal_doc / fengmian_doc1 等）仍可用，
    调用时会发出 DeprecationWarning，建议逐步迁移到新 API。
    """

    # ----------------------------------------------------------------
    # 对齐方式映射
    # ----------------------------------------------------------------
    _ALIGNMENT_MAP = {
        # 中文（支持"居X"与"X对齐"两种说法）
        "左对齐": WD_PARAGRAPH_ALIGNMENT.LEFT,
        "居左": WD_PARAGRAPH_ALIGNMENT.LEFT,
        "居中": WD_PARAGRAPH_ALIGNMENT.CENTER,
        "居右": WD_PARAGRAPH_ALIGNMENT.RIGHT,
        "右对齐": WD_PARAGRAPH_ALIGNMENT.RIGHT,
        "两端对齐": WD_PARAGRAPH_ALIGNMENT.JUSTIFY,
        # 英文（不区分大小写）
        "left": WD_PARAGRAPH_ALIGNMENT.LEFT,
        "center": WD_PARAGRAPH_ALIGNMENT.CENTER,
        "right": WD_PARAGRAPH_ALIGNMENT.RIGHT,
        "justify": WD_PARAGRAPH_ALIGNMENT.JUSTIFY,
        # 单字母 / 数字
        "L": WD_PARAGRAPH_ALIGNMENT.LEFT,
        "LEFT": WD_PARAGRAPH_ALIGNMENT.LEFT,
        "1": WD_PARAGRAPH_ALIGNMENT.LEFT,
        1: WD_PARAGRAPH_ALIGNMENT.LEFT,
        "M": WD_PARAGRAPH_ALIGNMENT.CENTER,
        "C": WD_PARAGRAPH_ALIGNMENT.CENTER,
        "CENTER": WD_PARAGRAPH_ALIGNMENT.CENTER,
        "2": WD_PARAGRAPH_ALIGNMENT.CENTER,
        2: WD_PARAGRAPH_ALIGNMENT.CENTER,
        "R": WD_PARAGRAPH_ALIGNMENT.RIGHT,
        "RIGHT": WD_PARAGRAPH_ALIGNMENT.RIGHT,
        "3": WD_PARAGRAPH_ALIGNMENT.RIGHT,
        3: WD_PARAGRAPH_ALIGNMENT.RIGHT,
    }
    _PAGE_JUSTIFICATION_MAP = {
        "左对齐": "left",
        "居左": "left",
        "居中": "center",
        "居右": "right",
        "右对齐": "right",
        "L": "left", "LEFT": "left", "1": "left", 1: "left",
        "M": "center", "C": "center", "CENTER": "center", "2": "center", 2: "center",
        "R": "right", "RIGHT": "right", "3": "right", 3: "right",
    }

    # ================================================================
    # 内部工具方法（公开可用，但建议优先使用高层 API）
    # ================================================================

    @staticmethod
    def check_item(item) -> bool:
        """判断单元格内容是否为非零数值（用于表格非空判断）。"""
        try:
            return float(item) != 0
        except (TypeError, ValueError):
            return False

    @staticmethod
    def set_run_font(run, *, cn_font: str, en_font: str = DEFAULT_EN_FONT,
                     size=14, bold: bool = False,
                     color: tuple[int, int, int] | None = None,
                     highlight=None) -> None:
        """设置 run 的字体（中英文分别设置）、字号、加粗、颜色、高亮。

        :param run: python-docx 的 Run 对象
        :param cn_font: 中文字体名（写入 w:eastAsia）
        :param en_font: 西文字体名（写入 w:ascii / w:hAnsi）
        :param size: 字号，可为数字（pt）或中文字号字符串（如 "三号"）
        :param bold: 是否加粗
        :param color: RGB 三元组，如 (255, 0, 0)
        :param highlight: WD_COLOR_INDEX 高亮颜色
        """
        run.font.bold = bold
        run.font.size = Pt(_to_pt(size))
        run.font.name = en_font
        run.element.rPr.rFonts.set(qn("w:eastAsia"), cn_font)
        if color is not None:
            run.font.color.rgb = RGBColor(*color)
        if highlight is not None:
            run.font.highlight_color = highlight

    # 旧名保留（私有方法被外部依赖，保留为公开别名）
    _set_run_font = set_run_font

    @staticmethod
    def set_paragraph_format(paragraph, *, alignment=None, line_spacing=1.5,
                              space_before=0, space_after=0,
                              first_line_indent_pt=None, strict: bool = False):
        """设置段落基础格式。

        :param alignment: 对齐方式，可为字符串（"居中"/"居左"/"居右"/"两端对齐"）
                          或 WD_PARAGRAPH_ALIGNMENT 枚举
        :param line_spacing: 行距倍数，1.0 / 1.5 / 2.0
        :param space_before: 段前空白（pt）
        :param space_after: 段后空白（pt）
        :param first_line_indent_pt: 首行缩进（pt），None 表示不设置
        :param strict: True 时未知对齐方式抛错，False 时静默回退到 LEFT（默认）
        """
        if alignment is not None:
            paragraph.alignment = WordFormatter.resolve_alignment(alignment, strict=strict)
        paragraph.paragraph_format.line_spacing = line_spacing
        paragraph.paragraph_format.space_before = Pt(space_before)
        paragraph.paragraph_format.space_after = Pt(space_after)
        if first_line_indent_pt is not None:
            paragraph.paragraph_format.first_line_indent = Pt(first_line_indent_pt)
        return paragraph

    _set_paragraph_basic_format = set_paragraph_format

    @staticmethod
    def resolve_alignment(alignment, *, strict: bool = False):
        """将对齐方式参数解析为 WD_PARAGRAPH_ALIGNMENT 枚举。

        支持：中英文（"居中"/"center"）、单字母（"C"/"L"/"R"）、数字（1/2/3）。
        :param strict: True 时未知值抛 ValueError；False 时回退 LEFT
        """
        if alignment is None:
            return None
        if isinstance(alignment, WD_PARAGRAPH_ALIGNMENT):
            return alignment
        if isinstance(alignment, str):
            key = alignment.strip()
            resolved = WordFormatter._ALIGNMENT_MAP.get(
                key.upper(),
                WordFormatter._ALIGNMENT_MAP.get(key),
            )
            if resolved is not None:
                return resolved
            if strict:
                raise ValueError(
                    f"未知对齐方式: {alignment!r}。支持: "
                    "'居中'/'居左'/'居右'/'两端对齐' 或 'center'/'left'/'right'/'justify'"
                )
            return WD_PARAGRAPH_ALIGNMENT.LEFT
        if isinstance(alignment, int):
            return WordFormatter._ALIGNMENT_MAP.get(alignment, WD_PARAGRAPH_ALIGNMENT.LEFT)
        return alignment

    _resolve_alignment = resolve_alignment

    @staticmethod
    def resolve_page_justification(alignment) -> str:
        """将页面对齐参数解析为字符串 'left'/'center'/'right'（用于 w:jc 域）。"""
        if isinstance(alignment, str):
            key = alignment.strip()
            return WordFormatter._PAGE_JUSTIFICATION_MAP.get(
                key.upper(),
                WordFormatter._PAGE_JUSTIFICATION_MAP.get(key, "left"),
            )
        if isinstance(alignment, int):
            return WordFormatter._PAGE_JUSTIFICATION_MAP.get(alignment, "left")
        return "left"

    _resolve_page_justification = resolve_page_justification

    @staticmethod
    def _append_field_run(paragraph, instruction: str):
        """向段落插入 Word 域（field），如 PAGE / TOC。"""
        begin = OxmlElement("w:fldChar")
        begin.set(qn("w:fldCharType"), "begin")

        instr = OxmlElement("w:instrText")
        instr.set(qn("xml:space"), "preserve")
        instr.text = instruction

        separate = OxmlElement("w:fldChar")
        separate.set(qn("w:fldCharType"), "separate")

        end = OxmlElement("w:fldChar")
        end.set(qn("w:fldCharType"), "end")

        begin_run = paragraph.add_run()
        begin_run._r.append(begin)
        instr_run = paragraph.add_run()
        instr_run._r.append(instr)
        separate_run = paragraph.add_run()
        separate_run._r.append(separate)
        end_run = paragraph.add_run()
        end_run._r.append(end)
        return [begin_run, instr_run, separate_run, end_run]

    @staticmethod
    def _ensure_update_fields_on_open(doc):
        """让 Word/OnlyOffice 打开文档时自动刷新域（PAGE / TOC）。"""
        settings = doc.settings.element
        existing = settings.find(qn("w:updateFields"))
        if existing is None:
            update_fields = OxmlElement("w:updateFields")
            update_fields.set(qn("w:val"), "true")
            settings.append(update_fields)
        else:
            existing.set(qn("w:val"), "true")
        return doc

    @staticmethod
    def _get_or_create_style(doc, style_name: str, base_style_name: str):
        styles = doc.styles
        try:
            return styles[style_name]
        except KeyError:
            style = styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
            try:
                style.base_style = styles[base_style_name]
            except KeyError:
                pass
            return style

    @staticmethod
    def _apply_style_options(style, options: dict):
        font_name = options.get("font_name")
        en_font = options.get("en_font")
        size = options.get("font_size")
        bold = options.get("bold")
        color = options.get("color")
        line_spacing = options.get("line_spacing")
        space_before = options.get("space_before")
        space_after = options.get("space_after")
        first_line_indent = options.get("first_line_indent")
        left_indent = options.get("left_indent")
        alignment = options.get("alignment")

        if font_name:
            style.font.name = en_font or DEFAULT_EN_FONT
            style.element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
        if size is not None:
            style.font.size = Pt(_to_pt(size))
        if bold is not None:
            style.font.bold = bold
        if color is not None:
            style.font.color.rgb = RGBColor(*color)
        if line_spacing is not None:
            style.paragraph_format.line_spacing = line_spacing
        if space_before is not None:
            style.paragraph_format.space_before = Pt(space_before)
        if space_after is not None:
            style.paragraph_format.space_after = Pt(space_after)
        if first_line_indent is not None:
            style.paragraph_format.first_line_indent = Pt(first_line_indent)
        if left_indent is not None:
            style.paragraph_format.left_indent = Pt(left_indent)
        if alignment is not None:
            style.paragraph_format.alignment = WordFormatter.resolve_alignment(alignment)
        return style

    @staticmethod
    def _configure_heading_style(doc, level: int, *, cn_font: str, en_font: str, size, bold: bool,
                                 color: tuple[int, int, int] | None = None, line_spacing=1,
                                 space_before=12, space_after=12, first_line_indent=28):
        style = doc.styles[f"Heading {level}"]
        return WordFormatter._apply_style_options(
            style,
            {
                "font_name": cn_font,
                "en_font": en_font,
                "font_size": size,
                "bold": bold,
                "color": color,
                "line_spacing": line_spacing,
                "space_before": space_before,
                "space_after": space_after,
                "first_line_indent": first_line_indent,
                "alignment": "左对齐",
            },
        )

    @staticmethod
    def _add_heading_with_style(doc, heading_text: str, *, level: int, cn_font: str, en_font: str,
                                size, bold: bool, color: tuple[int, int, int] | None = None,
                                line_spacing=1, space_before=12, space_after=12,
                                first_line_indent=28):
        WordFormatter._configure_heading_style(
            doc, level, cn_font=cn_font, en_font=en_font, size=size, bold=bold, color=color,
            line_spacing=line_spacing, space_before=space_before, space_after=space_after,
            first_line_indent=first_line_indent,
        )
        heading = doc.add_heading("", level=level)
        run = heading.add_run(heading_text)
        WordFormatter.set_run_font(run, cn_font=cn_font, en_font=en_font, size=size,
                                   bold=bold, color=color)
        return heading

    # ================================================================
    # 1. 文档级设置
    # ================================================================

    @staticmethod
    def set_default_font(doc, *, font_size=14, cn_font: str = DEFAULT_CN_FONT,
                         en_font: str = DEFAULT_EN_FONT):
        """设置文档 Normal 样式的默认字号、中文字体与西文字体。

        :param font_size: 字号，数字(pt) 或中文字号字符串（如 "三号"）
        :param cn_font: 中文字体名
        :param en_font: 西文字体名
        """
        style = doc.styles["Normal"]
        style.font.size = Pt(_to_pt(font_size))
        style.font.name = en_font
        style.element.rPr.rFonts.set(qn("w:eastAsia"), cn_font)
        return doc

    @staticmethod
    def set_language(doc, lang: str = "zh-CN"):
        """设置文档语言（影响拼写检查、字体回退、目录排序）。

        :param lang: BCP 47 语言标签，默认 "zh-CN"
        """
        settings = doc.settings.element

        theme_font_lang = settings.find(qn("w:themeFontLang"))
        if theme_font_lang is None:
            theme_font_lang = OxmlElement("w:themeFontLang")
            settings.append(theme_font_lang)
        theme_font_lang.set(qn("w:val"), lang)
        theme_font_lang.set(qn("w:eastAsia"), lang)
        theme_font_lang.set(qn("w:bidi"), lang)

        for style in doc.styles:
            rPr = style.element.get_or_add_rPr()
            lang_el = rPr.find(qn("w:lang"))
            if lang_el is None:
                lang_el = OxmlElement("w:lang")
                rPr.append(lang_el)
            lang_el.set(qn("w:val"), lang)
            lang_el.set(qn("w:eastAsia"), lang)
            lang_el.set(qn("w:bidi"), lang)
        return doc

    @staticmethod
    def set_page_margins(doc, *, top: float = 2.54, bottom: float = 2.54,
                         left: float = 3.17, right: float = 3.17,
                         gutter: float = 0, horizontal_alignment="L"):
        """设置全文页边距（单位：厘米）。

        :param top/bottom/left/right: 页边距，单位 cm
        :param gutter: 装订线宽度，单位 cm
        :param horizontal_alignment: 页面水平对齐，"L"/"C"/"R" 或 "左对齐"/"居中"/"右对齐"
        """
        jc_value = WordFormatter.resolve_page_justification(horizontal_alignment)
        for section in doc.sections:
            section.top_margin = Cm(top)
            section.bottom_margin = Cm(bottom)
            section.left_margin = Cm(left)
            section.right_margin = Cm(right)
            section.gutter = Cm(gutter)
            sect_pr = section._sectPr
            jc = sect_pr.find(qn("w:jc"))
            if jc is None:
                jc = OxmlElement("w:jc")
                sect_pr.append(jc)
            jc.set(qn("w:val"), jc_value)
        return doc

    # ================================================================
    # 2. 段落与标题
    # ================================================================

    @staticmethod
    def body(doc, text: str, *, font_name: str = DEFAULT_CN_FONT, font_size=14,
             indent: bool = True, alignment="两端对齐"):
        """添加正文段落。

        :param text: 正文文本
        :param font_name: 中文字体
        :param font_size: 字号
        :param indent: 是否首行缩进 28pt（约 2 字符）
        :param alignment: 对齐方式，默认两端对齐
        """
        par = doc.add_paragraph("")
        WordFormatter.set_paragraph_format(
            par,
            alignment=WordFormatter.resolve_alignment(alignment),
            first_line_indent_pt=28 if indent else None,
        )
        run = par.add_run(text)
        WordFormatter.set_run_font(run, cn_font=font_name, size=font_size, color=(0, 0, 0))
        return par

    @staticmethod
    def blank_lines(doc, count: int = 1):
        """添加若干空行（用于封面、签发页等留白）。"""
        for _ in range(count):
            doc.add_paragraph("")
        return doc

    @staticmethod
    def heading1(doc, text: str, *, font_name: str = "黑体", font_size=16):
        """添加一级标题（黑体 16pt，1.5 倍行距，首行缩进 28pt）。"""
        return WordFormatter._add_heading_with_style(
            doc, text, level=1, cn_font=font_name, en_font=DEFAULT_EN_FONT,
            size=font_size, bold=False, line_spacing=1.5, space_before=12,
            space_after=12, first_line_indent=28, color=(0, 0, 0),
        )

    @staticmethod
    def heading2(doc, text: str, *, font_name: str = DEFAULT_CN_FONT, font_size=14):
        """添加二级标题（仿宋 14pt 加粗）。"""
        return WordFormatter._add_heading_with_style(
            doc, text, level=2, cn_font=font_name, en_font=DEFAULT_EN_FONT,
            size=font_size, bold=True, color=(0, 0, 0),
        )

    @staticmethod
    def heading3(doc, text: str, *, font_name: str = DEFAULT_CN_FONT, font_size=14):
        """添加三级标题（仿宋 14pt）。"""
        return WordFormatter._add_heading_with_style(
            doc, text, level=3, cn_font=font_name, en_font=DEFAULT_EN_FONT,
            size=font_size, bold=False, color=(0, 0, 0),
        )

    @staticmethod
    def cover_text(doc, text: str, *, font_name: str = "宋体", font_size=22,
                   bold: bool = True):
        """添加封面文本（居中，用于项目名、报告名等）。

        替代旧版 fengmian_doc1/2/3 的统一方法。
        """
        par = doc.add_paragraph("")
        WordFormatter.set_paragraph_format(par, alignment=WD_PARAGRAPH_ALIGNMENT.CENTER)
        run = par.add_run(text)
        WordFormatter.set_run_font(run, cn_font=font_name, size=font_size, bold=bold)
        return par

    @staticmethod
    def right_text(doc, text: str, *, font_name: str = DEFAULT_CN_FONT, font_size=14,
                   indent: bool = True):
        """添加右对齐段落（用于公司名、日期等）。"""
        par = doc.add_paragraph("")
        WordFormatter.set_paragraph_format(
            par, alignment=WD_PARAGRAPH_ALIGNMENT.RIGHT,
            first_line_indent_pt=28 if indent else None,
        )
        run = par.add_run(text)
        WordFormatter.set_run_font(run, cn_font=font_name, size=font_size)
        return par

    @staticmethod
    def created_time(doc, value=None, *, font_name: str = DEFAULT_CN_FONT, font_size=14):
        """添加日期文本（右对齐）。

        :param value: None=今天，date/datetime=自动格式化，str=原样输出
        """
        if value is None:
            text = _datetime.datetime.now().strftime("%Y年%m月%d日")
        elif isinstance(value, (_datetime.date, _datetime.datetime)):
            text = value.strftime("%Y年%m月%d日")
        else:
            text = str(value)
        return WordFormatter.right_text(doc, text, font_name=font_name, font_size=font_size)

    # ================================================================
    # 3. 图片
    # ================================================================

    @staticmethod
    def insert_img(doc, img_path: str, width: float, *, alignment="居左"):
        """插入图片。

        :param img_path: 图片文件路径
        :param width: 图片宽度，单位 cm
        :param alignment: 段落对齐方式
        """
        par = doc.add_paragraph("")
        par.alignment = WordFormatter.resolve_alignment(alignment)
        run = par.add_run()
        run.add_picture(img_path, width=Cm(width))
        return par

    # ================================================================
    # 4. 表格（新增能力）
    # ================================================================

    @staticmethod
    def set_table_borders(table, *, color: str = "000000", size: int = 4):
        """为表格添加四边及内部横竖边框。

        :param table: docx.table.Table 对象
        :param color: 边框颜色，6 位十六进制 RGB 字符串
        :param size: 边框粗细，1/8 pt 为单位（4 = 0.5pt 细线，8 = 1pt）
        """
        tbl = table._tbl
        tbl_pr = tbl.tblPr
        tbl_borders = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            border = OxmlElement(f"w:{edge}")
            border.set(qn("w:val"), "single")
            border.set(qn("w:sz"), str(size))
            border.set(qn("w:space"), "0")
            border.set(qn("w:color"), color)
            tbl_borders.append(border)
        tbl_pr.append(tbl_borders)
        return table

    @staticmethod
    def set_cell(cell, text: str, *, font_name: str = DEFAULT_CN_FONT, font_size=12,
                 bold: bool = False, alignment="居中", line_spacing=1.25):
        """设置表格单元格内容与格式（增强版，支持加粗与对齐）。

        :param cell: docx.table._Cell 对象
        :param text: 单元格文本
        :param font_name: 中文字体
        :param font_size: 字号
        :param bold: 是否加粗（表头常用）
        :param alignment: 对齐方式，默认居中
        :param line_spacing: 行距倍数
        """
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WordFormatter.resolve_alignment(alignment, strict=True)
        paragraph.paragraph_format.line_spacing = line_spacing
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        # 清空原有 runs，避免叠加
        for run in list(paragraph.runs):
            run._r.getparent().remove(run._r)
        run = paragraph.add_run(text)
        WordFormatter.set_run_font(run, cn_font=font_name, size=font_size, bold=bold)
        return cell

    @staticmethod
    def add_table(doc, headers, rows, *, col_widths=None, font_size=12,
                  header_bold: bool = True, alignment="居中",
                  border_color: str = "000000", border_size: int = 4):
        """一站式创建带表头的表格。

        :param doc: Document 对象
        :param headers: 表头文本列表，如 ["序号", "项目", "金额"]
        :param rows: 二维数据列表，每个子列表为一行
        :param col_widths: 每列宽度（cm）列表，可选
        :param font_size: 表格字号
        :param header_bold: 表头是否加粗
        :param alignment: 单元格对齐方式
        :param border_color: 边框颜色
        :param border_size: 边框粗细（1/8 pt）
        :return: docx.table.Table 对象
        """
        table = doc.add_table(rows=1 + len(rows), cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        WordFormatter.set_table_borders(table, color=border_color, size=border_size)

        for i, h in enumerate(headers):
            WordFormatter.set_cell(
                table.rows[0].cells[i], h,
                font_size=font_size, bold=header_bold, alignment=alignment,
            )

        for r_idx, row in enumerate(rows, start=1):
            for c_idx, val in enumerate(row):
                WordFormatter.set_cell(
                    table.rows[r_idx].cells[c_idx], str(val),
                    font_size=font_size, alignment=alignment,
                )

        if col_widths:
            for r in table.rows:
                for i, w in enumerate(col_widths):
                    r.cells[i].width = Cm(w)
        return table

    # ================================================================
    # 5. 页眉页脚（新增能力）
    # ================================================================

    @staticmethod
    def set_header(section, text: str, *, alignment="居中",
                   font_name: str = DEFAULT_HEADER_FOOTER_FONT,
                   font_size=DEFAULT_HEADER_FOOTER_SIZE):
        """设置指定节的页眉文本（默认楷体小五号）。

        :param section: doc.sections[i]
        :param text: 页眉文本
        :param alignment: 对齐方式
        :param font_name: 字体（默认楷体）
        :param font_size: 字号（默认小五号 9pt）
        """
        return WordFormatter._fill_header_footer_part(
            section.header, text, alignment, font_name, font_size
        )

    @staticmethod
    def set_footer(section, text: str, *, alignment="居中",
                   font_name: str = DEFAULT_HEADER_FOOTER_FONT,
                   font_size=DEFAULT_HEADER_FOOTER_SIZE):
        """设置指定节的页脚文本（默认楷体小五号）。

        注意：本方法设置纯文本页脚；若需页码页脚请使用
        add_footer_page_number 或 set_page_number_from_section。
        """
        return WordFormatter._fill_header_footer_part(
            section.footer, text, alignment, font_name, font_size
        )

    @staticmethod
    def _fill_header_footer_part(part, text, alignment, font_name, font_size):
        """页眉/页脚共用填充实现。"""
        part.is_linked_to_previous = False
        paragraph = part.paragraphs[0] if part.paragraphs else part.add_paragraph()
        for run in list(paragraph.runs):
            run._r.getparent().remove(run._r)
        paragraph.alignment = WordFormatter.resolve_alignment(alignment, strict=True)
        paragraph.paragraph_format.line_spacing = 1
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        if text:
            run = paragraph.add_run(text)
            WordFormatter.set_run_font(
                run, cn_font=font_name, en_font=DEFAULT_EN_FONT,
                size=font_size, bold=False, color=(0, 0, 0),
            )
        return part

    @staticmethod
    def clear_footer(section):
        """清空指定节的页脚内容。"""
        section.footer.is_linked_to_previous = False
        for paragraph in section.footer.paragraphs:
            for run in list(paragraph.runs):
                run._r.getparent().remove(run._r)
        return section

    @staticmethod
    def clear_header(section):
        """清空指定节的页眉内容。"""
        section.header.is_linked_to_previous = False
        for paragraph in section.header.paragraphs:
            for run in list(paragraph.runs):
                run._r.getparent().remove(run._r)
        return section

    # ================================================================
    # 6. 节与页码
    # ================================================================

    @staticmethod
    def insert_section(doc, start_type=WD_SECTION_START.NEW_PAGE):
        """插入新节。"""
        return doc.add_section(start_type=start_type)

    @staticmethod
    def set_page_number_start(section, start: int = 1):
        """设置指定节页码起始值。"""
        sect_pr = section._sectPr
        pg_num_type = sect_pr.find(qn("w:pgNumType"))
        if pg_num_type is None:
            pg_num_type = OxmlElement("w:pgNumType")
            sect_pr.append(pg_num_type)
        pg_num_type.set(qn("w:start"), str(start))
        return section

    @staticmethod
    def add_page_number(paragraph, *, prefix: str = "", suffix: str = "",
                       alignment="居中", font_name: str | None = None,
                       font_size=None):
        """向段落添加页码域（PAGE field）。

        :param paragraph: 段落对象（通常是页脚段落）
        :param prefix: 页码前文本，如 "第 "
        :param suffix: 页码后文本，如 " 页"
        :param alignment: 对齐方式
        :param font_name: 字体（None 时使用默认）
        :param font_size: 字号
        """
        paragraph.alignment = WordFormatter.resolve_alignment(alignment)
        cn_font = font_name or DEFAULT_CN_FONT
        size = font_size if font_size is not None else DEFAULT_BODY_SIZE

        if prefix:
            prefix_run = paragraph.add_run(prefix)
            WordFormatter.set_run_font(prefix_run, cn_font=cn_font, size=size)
        field_runs = WordFormatter._append_field_run(paragraph, "PAGE")
        for run in field_runs:
            WordFormatter.set_run_font(run, cn_font=cn_font, size=size)
        if suffix:
            suffix_run = paragraph.add_run(suffix)
            WordFormatter.set_run_font(suffix_run, cn_font=cn_font, size=size)
        return paragraph

    @staticmethod
    def add_footer_page_number(section, *, prefix: str = "", suffix: str = "",
                                alignment="居中", font_name: str | None = None,
                                font_size=None):
        """向指定节页脚添加页码。"""
        section.footer.is_linked_to_previous = False
        footer = section.footer
        paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        if paragraph.text:
            paragraph = footer.add_paragraph()
        return WordFormatter.add_page_number(
            paragraph, prefix=prefix, suffix=suffix, alignment=alignment,
            font_name=font_name, font_size=font_size,
        )

    @staticmethod
    def restart_page_numbering(section, *, start: int = 1, add_footer_number: bool = True,
                               prefix: str = "", suffix: str = "", alignment="居中",
                               font_name: str | None = None, font_size=None):
        """重启指定节的页码编号（单节）。"""
        section.header.is_linked_to_previous = False
        section.footer.is_linked_to_previous = False
        WordFormatter.set_page_number_start(section, start=start)
        if add_footer_number:
            WordFormatter.add_footer_page_number(
                section, prefix=prefix, suffix=suffix, alignment=alignment,
                font_name=font_name, font_size=font_size,
            )
        return section

    @staticmethod
    def insert_section_with_page_numbering(doc, *, start_type=WD_SECTION_START.NEW_PAGE,
                                           start_page_number: int = 1,
                                           add_footer_number: bool = True,
                                           prefix: str = "", suffix: str = "",
                                           alignment="居中", font_name: str | None = None,
                                           font_size=None):
        """插入新节并重启页码编号（便捷组合方法）。"""
        section = WordFormatter.insert_section(doc, start_type=start_type)
        return WordFormatter.restart_page_numbering(
            section, start=start_page_number, add_footer_number=add_footer_number,
            prefix=prefix, suffix=suffix, alignment=alignment,
            font_name=font_name, font_size=font_size,
        )

    @staticmethod
    def set_page_number_from_section(doc, start_section_idx: int, *, start: int = 1,
                                    prefix: str = "", suffix: str = "",
                                    alignment="居中", font_name: str | None = None,
                                    font_size=None):
        """从指定节开始设置页码（包含此节），之前的节无页码。

        一次性完成：
        1) start_section_idx 之前的节：清空页脚，无页码
        2) start_section_idx 节：页码从 start 重新开始，添加页脚页码
        3) start_section_idx 之后的节：链接到前一节页脚，页码连续

        :param doc: Document 对象
        :param start_section_idx: 从 0 开始的节索引
        :param start: 起始页码
        """
        sections = doc.sections
        total = len(sections)
        if start_section_idx < 0 or start_section_idx >= total:
            raise ValueError(
                f"start_section_idx {start_section_idx} 超出范围"
                f"（共 {total} 节，0~{total - 1}）"
            )

        for i in range(start_section_idx):
            WordFormatter.clear_footer(sections[i])

        WordFormatter.restart_page_numbering(
            sections[start_section_idx], start=start, add_footer_number=True,
            prefix=prefix, suffix=suffix, alignment=alignment,
            font_name=font_name, font_size=font_size,
        )

        for i in range(start_section_idx + 1, total):
            sections[i].footer.is_linked_to_previous = True
        return doc

    # ================================================================
    # 7. 目录
    # ================================================================

    @staticmethod
    def set_toc_level_style(doc, level: int, *, font_name: str | None = None, font_size=None,
                            bold: bool | None = None, color: tuple[int, int, int] | None = None,
                            line_spacing=None, space_before=None, space_after=None,
                            first_line_indent=None, left_indent=None, alignment=None):
        """设置某一级 TOC 样式。"""
        style_name = f"TOC {level}"
        style = WordFormatter._get_or_create_style(doc, style_name, "Normal")
        return WordFormatter._apply_style_options(
            style,
            {
                "font_name": font_name, "font_size": font_size, "bold": bold,
                "color": color, "line_spacing": line_spacing,
                "space_before": space_before, "space_after": space_after,
                "first_line_indent": first_line_indent, "left_indent": left_indent,
                "alignment": alignment,
            },
        )

    @staticmethod
    def set_paragraph_style(doc, style_name: str, *, base_style_name: str = "Normal",
                            font_name: str | None = None, font_size=None,
                            bold: bool | None = None,
                            color: tuple[int, int, int] | None = None,
                            line_spacing=None, space_before=None, space_after=None,
                            first_line_indent=None, left_indent=None, alignment=None):
        """创建或更新自定义段落样式。"""
        style = WordFormatter._get_or_create_style(doc, style_name, base_style_name)
        return WordFormatter._apply_style_options(
            style,
            {
                "font_name": font_name, "font_size": font_size, "bold": bold,
                "color": color, "line_spacing": line_spacing,
                "space_before": space_before, "space_after": space_after,
                "first_line_indent": first_line_indent, "left_indent": left_indent,
                "alignment": alignment,
            },
        )

    @staticmethod
    def add_custom_heading(doc, text_content: str, *, style_name: str,
                          level: int | None = None, font_name: str | None = None,
                          font_size=None):
        """使用自定义段落样式添加标题（用于不使用内置 Heading 样式的场景）。"""
        par = doc.add_paragraph(text_content, style=style_name)
        if par.runs and (font_name is not None or font_size is not None):
            WordFormatter.set_run_font(
                par.runs[0],
                cn_font=font_name or DEFAULT_CN_FONT,
                size=font_size or DEFAULT_BODY_SIZE,
            )
        if level is not None:
            p_pr = par._p.get_or_add_pPr()
            outline_lvl = p_pr.find(qn("w:outlineLvl"))
            if outline_lvl is None:
                outline_lvl = OxmlElement("w:outlineLvl")
                p_pr.append(outline_lvl)
            outline_lvl.set(qn("w:val"), str(max(level - 1, 0)))
        return par

    @staticmethod
    def add_toc(doc, *, title: str = "目录", levels: tuple[int, int] = (1, 3),
                title_font_name: str = "黑体", title_font_size=16,
                title_bold: bool = True, title_alignment="居中",
                toc_level_styles: dict[int, dict] | None = None,
                use_hyperlinks: bool = True, hide_page_numbers_in_web: bool = True,
                use_outline_levels: bool = False,
                custom_style_levels: dict[str, int] | None = None,
                page_number_separator: str | None = None,
                right_align_page_numbers: bool | None = None):
        """插入 Word 目录域（TOC field）。

        生成的目录需在 Word/OnlyOffice 中刷新域（F9）才会显示页码。

        :param title: 目录标题，None 时不生成标题
        :param levels: 收集的标题层级范围，如 (1, 3)
        :param toc_level_styles: 各级 TOC 样式配置
        :param use_hyperlinks: 是否生成超链接
        :param use_outline_levels: 是否使用大纲级别
        :param custom_style_levels: 自定义样式 -> 目录层级映射
        """
        min_level, max_level = levels
        if min_level < 1 or max_level < min_level:
            raise ValueError("levels 参数不合法，例如 (1, 3)")

        WordFormatter._ensure_update_fields_on_open(doc)

        if title:
            title_par = doc.add_paragraph("")
            WordFormatter.set_paragraph_format(
                title_par,
                alignment=WordFormatter.resolve_alignment(title_alignment),
                space_before=12, space_after=12,
            )
            title_run = title_par.add_run(title)
            WordFormatter.set_run_font(
                title_run, cn_font=title_font_name, size=title_font_size, bold=title_bold,
            )

        if toc_level_styles:
            for level, options in toc_level_styles.items():
                WordFormatter.set_toc_level_style(doc, level, **options)

        switches = [
            f'\\o "{min_level}-{max_level}"',
            "\\h" if use_hyperlinks else "",
            "\\z" if hide_page_numbers_in_web else "",
        ]
        if use_outline_levels:
            switches.append("\\u")
        if custom_style_levels:
            style_level_text = ",".join(
                f"{name},{lvl}" for name, lvl in custom_style_levels.items()
            )
            switches.append(f'\\t "{style_level_text}"')
        if page_number_separator is not None:
            switches.append(f'\\p "{page_number_separator}"')
        # right_align_page_numbers: 预留接口，python-docx 暂不支持直接控制
        instruction = f'TOC {" ".join(s for s in switches if s)}'

        toc_par = doc.add_paragraph("")
        WordFormatter._append_field_run(toc_par, instruction)
        return toc_par

    # ================================================================
    # 8. 向后兼容（v0.1.x 旧 API，调用时发 DeprecationWarning）
    # ================================================================

    @staticmethod
    def _deprecate(old_name: str, new_name: str):
        warnings.warn(
            f"`WordFormatter.{old_name}` 已废弃，请使用 `WordFormatter.{new_name}`。",
            DeprecationWarning, stacklevel=3,
        )

    @staticmethod
    def fengmian_doc1(doc, text_content: str, font_name: str = "宋体", font_size: int = 22):
        """[已废弃] 使用 cover_text(doc, text, font_name=..., font_size=..., bold=True)。"""
        WordFormatter._deprecate("fengmian_doc1", "cover_text")
        return WordFormatter.cover_text(doc, text_content, font_name=font_name,
                                        font_size=font_size, bold=True)

    @staticmethod
    def fengmian_doc2(doc, text_content: str, font_name: str = "方正小标宋简体",
                      font_size: int = 16):
        """[已废弃] 使用 cover_text(doc, text, font_name=..., font_size=..., bold=False)。"""
        WordFormatter._deprecate("fengmian_doc2", "cover_text")
        return WordFormatter.cover_text(doc, text_content, font_name=font_name,
                                        font_size=font_size, bold=False)

    @staticmethod
    def fengmian_doc3(doc, text_content: str, font_name: str = "黑体", font_size: int = 22):
        """[已废弃] 使用 cover_text。"""
        WordFormatter._deprecate("fengmian_doc3", "cover_text")
        return WordFormatter.cover_text(doc, text_content, font_name=font_name,
                                        font_size=font_size, bold=True)

    @staticmethod
    def Heading_1(doc, heading_text: str, font_name: str = "黑体", font_size: int = 16):
        """[已废弃] 使用 heading1。"""
        WordFormatter._deprecate("Heading_1", "heading1")
        return WordFormatter.heading1(doc, heading_text, font_name=font_name, font_size=font_size)

    @staticmethod
    def Heading_2(doc, heading_text: str, font_name: str = DEFAULT_CN_FONT,
                  font_size: int = 14):
        """[已废弃] 使用 heading2。"""
        WordFormatter._deprecate("Heading_2", "heading2")
        return WordFormatter.heading2(doc, heading_text, font_name=font_name, font_size=font_size)

    @staticmethod
    def Heading_3(doc, heading_text: str, font_name: str = DEFAULT_CN_FONT,
                   font_size: int = 14):
        """[已废弃] 使用 heading3。"""
        WordFormatter._deprecate("Heading_3", "heading3")
        return WordFormatter.heading3(doc, heading_text, font_name=font_name, font_size=font_size)

    @staticmethod
    def Heading_union(doc, text_content: str, layout: str, font_name: str, font_size: int):
        """[已废弃] 使用 cover_text 加 alignment 参数。"""
        WordFormatter._deprecate("Heading_union", "cover_text")
        par = doc.add_heading("", level=2)
        par.alignment = WordFormatter.resolve_alignment(layout) or WD_PARAGRAPH_ALIGNMENT.RIGHT
        par.paragraph_format.line_spacing = 1
        par.paragraph_format.space_before = Pt(12)
        par.paragraph_format.space_after = Pt(12)
        run = par.add_run(text_content)
        WordFormatter.set_run_font(run, cn_font=font_name, size=font_size,
                                   bold=True, color=(0, 0, 0))
        return par

    @staticmethod
    def Normal_doc(doc, text_content: str, font_name: str = DEFAULT_CN_FONT,
                   font_size: int = 14):
        """[已废弃] 使用 body。"""
        WordFormatter._deprecate("Normal_doc", "body")
        return WordFormatter.body(doc, text_content, font_name=font_name, font_size=font_size)

    @staticmethod
    def Normal_doc_Highlight(doc, text_content: str, font_name: str = DEFAULT_CN_FONT,
                             font_size: int = 14):
        """[已废弃] 使用 body 加自定 run 高亮。"""
        WordFormatter._deprecate("Normal_doc_Highlight", "body")
        par = WordFormatter.body(doc, text_content, font_name=font_name, font_size=font_size)
        par.runs[-1].font.highlight_color = WD_COLOR_INDEX.YELLOW
        return par

    @staticmethod
    def Normal_doc_仿宋三号加粗(doc, text_content: str,
                              font_name: str = DEFAULT_CN_FONT, font_size: int = 16):
        """[已废弃] 使用 cover_text(doc, text, font_size=16, bold=True)。"""
        WordFormatter._deprecate("Normal_doc_仿宋三号加粗", "cover_text")
        par = doc.add_paragraph("")
        WordFormatter.set_paragraph_format(par, alignment=WD_PARAGRAPH_ALIGNMENT.CENTER)
        run = par.add_run(text_content)
        WordFormatter.set_run_font(run, cn_font=font_name, size=font_size,
                                   bold=True, color=(0, 0, 0))
        return par

    @staticmethod
    def set_cell_format(cell, text_content: str, font_name: str = DEFAULT_CN_FONT,
                        font_size: int = 12):
        """[已废弃] 使用 set_cell（支持加粗、对齐方式）。"""
        WordFormatter._deprecate("set_cell_format", "set_cell")
        return WordFormatter.set_cell(
            cell, text_content, font_name=font_name, font_size=font_size,
            bold=False, alignment="两端对齐", line_spacing=1.5,
        )

    @staticmethod
    def company_name(doc, company_name: str, font_name: str = DEFAULT_CN_FONT,
                     font_size: int = 16):
        """[已废弃] 使用 right_text。"""
        WordFormatter._deprecate("company_name", "right_text")
        return WordFormatter.right_text(doc, company_name, font_name=font_name,
                                        font_size=font_size, indent=True)

    @staticmethod
    def set_all_layout(doc, *, top: float = 2.54, bottom: float = 2.54,
                       left: float = 3.17, right: float = 3.17, gutter: float = 0):
        """[已废弃] 使用 set_page_margins。"""
        WordFormatter._deprecate("set_all_layout", "set_page_margins")
        return WordFormatter.set_page_margins(
            doc, top=top, bottom=bottom, left=left, right=right, gutter=gutter,
        )

    @staticmethod
    def set_document_layout(doc, *, top: float = 2.54, bottom: float = 2.54,
                            left: float = 3.17, right: float = 3.17, gutter: float = 0,
                            horizontal_alignment="L"):
        """[已废弃] 使用 set_page_margins。"""
        WordFormatter._deprecate("set_document_layout", "set_page_margins")
        return WordFormatter.set_page_margins(
            doc, top=top, bottom=bottom, left=left, right=right, gutter=gutter,
            horizontal_alignment=horizontal_alignment,
        )

    @staticmethod
    def insert_new_section(doc, start_type=WD_SECTION_START.NEW_PAGE):
        """[已废弃] 使用 insert_section。"""
        WordFormatter._deprecate("insert_new_section", "insert_section")
        return WordFormatter.insert_section(doc, start_type=start_type)


about_word = WordFormatter
