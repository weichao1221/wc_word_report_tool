"""wc_word_report_tool.word
=========================
基于 python-docx 的中文 Word 报告格式化工具。

设计目标
--------
1. **覆盖完整**：表格、页眉页脚、目录、页码、封面等报告常见需求一站式覆盖。
2. **命名一致**：所有公开方法使用 snake_case，参数命名清晰，单位明确。
3. **默认合理**：默认值贴近中国公文标准（仿宋_GB2312、三号 14pt、1.5 倍行距）。
4. **严格校验**：未知参数抛 ValueError 而非静默降级，便于调试。

版本：v0.4.5
作者：willcha
"""

from __future__ import annotations

import datetime as _datetime
from pathlib import Path

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
DEFAULT_BODY_SIZE = 14          # 四号
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

    文档级操作通过 ``WordFormatter(doc)`` 绑定文档后调用；纯工具方法保留为静态方法。

    命名约定（v0.4.4+）
    ------------------
    - snake_case 全小写：``heading`` / ``body`` / ``set_header``
    - 设置类前缀 ``set_``：``set_cell`` / ``set_header`` / ``set_table_borders``
    - 添加类前缀 ``add_``：``add_table`` / ``add_toc`` / ``add_page_number``
    - 单位：长度 cm，字号 pt 或中文字号（"三号"），页边距 cm

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

    def __init__(self, doc):
        """创建绑定到指定 Word 文档的格式化器。"""
        self.doc = doc

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
        """判断单元格内容是否为非零数值。

        :param item: 待判断的单元格值。
        """
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
        :param size: 字号，支持中文字号字符串，如 ``"三号"``、``"四号"``、``"小五"``。
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

    @staticmethod
    def set_paragraph_format(paragraph, *, alignment=None, line_spacing=1.5,
                              space_before=0, space_after=0,
                              first_line_indent_pt=None, strict: bool = False):
        """设置段落基础格式。

        :param alignment: 对齐方式，支持中文、英文、单字母、数字和对齐枚举。
        :param paragraph: 目标段落。
        :param line_spacing: 行距倍数。
        :param space_before: 段前空白，单位 pt。
        :param space_after: 段后空白，单位 pt。
        :param first_line_indent_pt: 首行缩进，单位 pt；None 表示不设置。
        :param strict: 是否对未知对齐方式抛出 ValueError，默认否。
        """
        if alignment is not None:
            paragraph.alignment = WordFormatter.resolve_alignment(alignment, strict=strict)
        paragraph.paragraph_format.line_spacing = line_spacing
        paragraph.paragraph_format.space_before = Pt(space_before)
        paragraph.paragraph_format.space_after = Pt(space_after)
        if first_line_indent_pt is not None:
            paragraph.paragraph_format.first_line_indent = Pt(first_line_indent_pt)
        return paragraph


    @staticmethod
    def resolve_alignment(alignment, *, strict: bool = False):
        """将对齐方式参数解析为 WD_PARAGRAPH_ALIGNMENT 枚举。

        支持中文（居中/居左/居右/两端对齐）、英文（center/left/right/justify）、
        单字母（C/L/R）、数字（1/2/3）和 WD_PARAGRAPH_ALIGNMENT 枚举。
        :param alignment: 对齐方式，可传中文、英文、单字母、数字或对齐枚举。
        :param strict: True 时未知值抛 ValueError；False 时回退 LEFT。
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


    @staticmethod
    def resolve_page_justification(alignment) -> str:
        """将页面对齐参数解析为页面 XML 使用的字符串。

        :param alignment: 页面水平对齐方式，支持左/中/右对齐写法。
        """
        if isinstance(alignment, str):
            key = alignment.strip()
            return WordFormatter._PAGE_JUSTIFICATION_MAP.get(
                key.upper(),
                WordFormatter._PAGE_JUSTIFICATION_MAP.get(key, "left"),
            )
        if isinstance(alignment, int):
            return WordFormatter._PAGE_JUSTIFICATION_MAP.get(alignment, "left")
        return "left"


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

    def _ensure_update_fields_on_open(self):
        """让 Word/OnlyOffice 打开文档时自动刷新域（PAGE / TOC）。"""
        settings = self.doc.settings.element
        existing = settings.find(qn("w:updateFields"))
        if existing is None:
            update_fields = OxmlElement("w:updateFields")
            update_fields.set(qn("w:val"), "true")
            settings.append(update_fields)
        else:
            existing.set(qn("w:val"), "true")
        return self.doc

    def _get_or_create_style(self, style_name: str, base_style_name: str):
        styles = self.doc.styles
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

    def _configure_heading_style(self, level: int, *, cn_font: str, en_font: str, size, bold: bool,
                                 color: tuple[int, int, int] | None = None, line_spacing=1,
                                 space_before=12, space_after=12, first_line_indent=28):
        style = self.doc.styles[f"Heading {level}"]
        return self._apply_style_options(
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

    def _add_heading_with_style(self, heading_text: str, *, level: int, cn_font: str, en_font: str,
                                size, bold: bool, color: tuple[int, int, int] | None = None,
                                line_spacing=1, space_before=12, space_after=12,
                                first_line_indent=28):
        self._configure_heading_style(
            level, cn_font=cn_font, en_font=en_font, size=size, bold=bold, color=color,
            line_spacing=line_spacing, space_before=space_before, space_after=space_after,
            first_line_indent=first_line_indent,
        )
        heading = self.doc.add_heading("", level=level)
        run = heading.add_run(heading_text)
        self.set_run_font(run, cn_font=cn_font, en_font=en_font, size=size,
                                   bold=bold, color=color)
        return heading

    # ================================================================
    # 1. 文档级设置
    # ================================================================

    def set_default_font(self, *, font_size=14, cn_font: str = DEFAULT_CN_FONT,
                         en_font: str = DEFAULT_EN_FONT):
        """设置文档 Normal 样式的默认字号、中文字体与西文字体。

        :param font_size: 字号，支持中文字号字符串，如 ``"三号"``、``"四号"``、``"小五"``。
        :param cn_font: 中文字体名称。
        :param en_font: 西文字体名称。
        :return: 当前文档对象。
        """
        style = self.doc.styles["Normal"]
        style.font.size = Pt(_to_pt(font_size))
        style.font.name = en_font
        style.element.rPr.rFonts.set(qn("w:eastAsia"), cn_font)
        return self.doc

    def setup_defaults(self, *, top: float = 3.7, bottom: float = 3.5,
                       left: float = 2.8, right: float = 2.6, gutter: float = 0,
                       body_font_name: str = "宋体", body_font_size="四号",
                       body_line_spacing=1.5, body_indent_chars: float = 2,
                       alignment="两端对齐", language: str = "zh-CN"):
        """一次设置文档页边距、Normal 正文样式和文档语言。

        :param top/bottom/left/right: 页边距，单位 cm。
        :param gutter: 装订线，单位 cm。
        :param body_font_name: 正文字体名称。
        :param body_font_size: 正文大小，支持中文字号字符串。
        :param body_line_spacing: 正文行距倍数。
        :param body_indent_chars: 正文首行缩进字符数，默认 2。
        :param alignment: 正文对齐方式，支持中文、英文、单字母、数字和对齐枚举。
        :param language: 文档语言标签，默认 ``"zh-CN"``。
        """
        self.set_page_margins(
            top=top, bottom=bottom, left=left, right=right, gutter=gutter,
        )
        self.set_default_font(
            font_size=body_font_size, cn_font=body_font_name,
        )
        normal = self.doc.styles["Normal"]
        normal.paragraph_format.line_spacing = body_line_spacing
        normal.paragraph_format.space_before = Pt(0)
        normal.paragraph_format.space_after = Pt(0)
        normal.paragraph_format.first_line_indent = Pt(
            _to_pt(body_font_size) * body_indent_chars
        )
        normal.paragraph_format.alignment = self.resolve_alignment(
            alignment, strict=True,
        )
        self.set_language(language)
        return self.doc

    def set_language(self, lang: str = "zh-CN"):
        """设置文档语言（影响拼写检查、字体回退、目录排序）。

        :param lang: BCP 47 语言标签，默认 "zh-CN"
        """
        settings = self.doc.settings.element

        theme_font_lang = settings.find(qn("w:themeFontLang"))
        if theme_font_lang is None:
            theme_font_lang = OxmlElement("w:themeFontLang")
            settings.append(theme_font_lang)
        theme_font_lang.set(qn("w:val"), lang)
        theme_font_lang.set(qn("w:eastAsia"), lang)
        theme_font_lang.set(qn("w:bidi"), lang)

        for style in self.doc.styles:
            rPr = style.element.get_or_add_rPr()
            lang_el = rPr.find(qn("w:lang"))
            if lang_el is None:
                lang_el = OxmlElement("w:lang")
                rPr.append(lang_el)
            lang_el.set(qn("w:val"), lang)
            lang_el.set(qn("w:eastAsia"), lang)
            lang_el.set(qn("w:bidi"), lang)
        return self.doc

    def set_page_margins(self, *, top: float = 2.54, bottom: float = 2.54,
                         left: float = 3.17, right: float = 3.17,
                         gutter: float = 0, horizontal_alignment="L"):
        """设置全文页边距（单位：厘米）。

        :param top/bottom/left/right: 页边距，单位 cm。
        :param gutter: 装订线宽度，单位 cm。
        :param horizontal_alignment: 页面水平对齐，支持 ``L``/``C``/``R``、
                                     ``左对齐``/``居中``/``右对齐``。
        :return: 当前文档对象。
        """
        jc_value = self.resolve_page_justification(horizontal_alignment)
        for section in self.doc.sections:
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
        return self.doc

    # ================================================================
    # 2. 段落与标题
    # ================================================================

    def body(self, text: str, *, font_name: str = "宋体", font_size="四号",
             indent: bool = True, bold: bool = False, highlight: bool = False,
             alignment="两端对齐", line_spacing=1.5,
             space_before=0, space_after=0):
        """添加正文段落。

        默认格式：宋体、四号、首行缩进 2 字符、两端对齐、1.5 倍行距。

        :param text: 正文文本。
        :param font_name: 中文字体名称，默认 ``"宋体"``；可传任意本机已安装字体。
        :param font_size: 字号，默认 ``"四号"``。支持数字/浮点数（单位 pt）、
                          数字字符串，以及 ``FONT_SIZE_MAP`` 中的中文字号，
                          例如 ``"三号"``、``"四号"``、``"小五"``。
        :param indent: 是否首行缩进 2 字符，默认 ``True``。缩进值按当前字号
                       乘以 2 换算，因此不是固定 28pt。
        :param bold: 是否加粗，默认 ``False``。
        :param highlight: 是否使用黄色高亮，默认 ``False``。
        :param alignment: 段落对齐方式，默认 ``"两端对齐"``。支持
                          ``"居左"``/``"左对齐"``、``"居中"``、
                          ``"居右"``/``"右对齐"``、``"两端对齐"``，
                          英文 ``"left"``/``"center"``/``"right"``/
                          ``"justify"``，单字母 ``"L"``/``"C"``/``"R"``，
                          数字 ``1``/``2``/``3``，以及
                          ``WD_PARAGRAPH_ALIGNMENT`` 枚举。
        :param line_spacing: 行距倍数，默认 ``1.5``；例如 ``1.0``、``1.5``、``2.0``。
        :param space_before: 段前空白，默认 ``0``，单位 pt。
        :param space_after: 段后空白，默认 ``0``，单位 pt。
        :return: 新创建的 ``Paragraph`` 对象。
        """
        par = self.doc.add_paragraph("")
        self.set_paragraph_format(
            par,
            alignment=self.resolve_alignment(alignment),
            line_spacing=line_spacing,
            space_before=space_before,
            space_after=space_after,
            first_line_indent_pt=_to_pt(font_size) * 2 if indent else None,
        )
        run = par.add_run(text)
        self.set_run_font(
            run, cn_font=font_name, size=font_size, color=(0, 0, 0),
            bold=bold,
            highlight=WD_COLOR_INDEX.YELLOW if highlight else None,
        )
        return par

    def blank_lines(self, count: int = 1):
        """添加若干空行（用于封面、签发页等留白）。

        :param count: 添加数量，默认 1。
        """
        for _ in range(count):
            self.doc.add_paragraph("")
        return self.doc

    def heading(self, text: str, *, level: int = 1,
                font_name: str = "黑体", font_size="三号",
                bold: bool = False,
                indent: bool = False, line_spacing=1.5,
                 space_before=None, space_after=None):
        """添加通用标题，默认一级标题格式。

        :param text: 标题文本。
        :param level: 标题层级，1 到 9，默认 1。
        :param font_name: 标题字体名称。
        :param font_size: 标题字号，支持中文字号字符串。
        :param bold: 是否加粗，默认否。
        :param indent: 是否首行缩进 2 字符，默认否。
        :param line_spacing: 行距倍数，默认 1.5。
        :param space_before: 段前空白，单位 pt；不传时按字号的 0.5 行计算。
        :param space_after: 段后空白，单位 pt；不传时按字号的 0.5 行计算。
        """
        if level < 1 or level > 9:
            raise ValueError("标题 level 必须在 1 到 9 之间")
        return self._add_heading_with_style(
            text, level=level, cn_font=font_name, en_font=DEFAULT_EN_FONT,
            size=font_size, bold=bold, line_spacing=line_spacing,
            space_before=_to_pt(font_size) * 0.5 if space_before is None else space_before,
            space_after=_to_pt(font_size) * 0.5 if space_after is None else space_after,
            first_line_indent=_to_pt(font_size) * 2 if indent else None, color=(0, 0, 0),
        )

    def cover_text(self, text: str, *, font_name: str = "方正小标宋简体",
                   font_size=24, bold: bool = True, alignment="居中",
                   indent: bool = False, line_spacing=2,
                   space_before=None, space_after=None):
        """添加封面文本，默认居中加粗、2 倍行距。

        :param text: 封面文本。
        :param font_name: 字体名称。
        :param font_size: 字号，支持中文字号字符串。
        :param bold: 是否加粗，默认是。
        :param alignment: 对齐方式，支持中文、英文、单字母、数字和对齐枚举。
        :param indent: 是否首行缩进 2 字符，默认否。
        :param line_spacing: 行距倍数，默认 2。
        :param space_before: 段前空白，单位 pt；不传时按字号的 1 行计算。
        :param space_after: 段后空白，单位 pt；不传时按字号的 1 行计算。
        """
        par = self.doc.add_paragraph("")
        self.set_paragraph_format(
            par, alignment=alignment, line_spacing=line_spacing,
            space_before=_to_pt(font_size) if space_before is None else space_before,
            space_after=_to_pt(font_size) if space_after is None else space_after,
            first_line_indent_pt=_to_pt(font_size) * 2 if indent else None,
        )
        run = par.add_run(text)
        self.set_run_font(run, cn_font=font_name, size=font_size, bold=bold)
        return par

    def right_text(self, text: str, *, font_name: str = DEFAULT_CN_FONT, font_size=14,
                   indent: bool = True):
        """添加右对齐段落（用于公司名、日期等）。

        :param text: 段落文本。
        :param font_name: 字体名称。
        :param font_size: 字号，支持中文字号字符串。
        :param indent: 是否首行缩进 2 字符，默认是。
        """
        par = self.doc.add_paragraph("")
        self.set_paragraph_format(
            par, alignment=WD_PARAGRAPH_ALIGNMENT.RIGHT,
            first_line_indent_pt=28 if indent else None,
        )
        run = par.add_run(text)
        self.set_run_font(run, cn_font=font_name, size=font_size)
        return par

    @staticmethod
    def _date_value(value=None):
        if value is None:
            return _datetime.datetime.now()
        if isinstance(value, _datetime.datetime):
            return value
        if isinstance(value, _datetime.date):
            return _datetime.datetime.combine(value, _datetime.time())
        text = str(value).strip()
        for fmt in ("%Y年%m月%d日", "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
            try:
                return _datetime.datetime.strptime(text, fmt)
            except ValueError:
                continue
        raise ValueError("日期格式不正确，请使用 2026年07月16日 或 2026-07-16")

    def created_time(self, value=None, *, font_name: str = DEFAULT_CN_FONT,
                     font_size=14, bold: bool = False):
        """添加日期文本（右对齐）。

        :param value: 不传时使用今天；支持 date/datetime，或 ``2026年07月16日``、
                      ``2026-07-16``、``2026/07/16``、``2026.07.16`` 字符串。
        :param font_name: 字体名称。
        :param font_size: 字号，支持中文字号字符串。
        :param bold: 是否加粗，默认否。
        """
        text = self._date_value(value).strftime("%Y年%m月%d日")
        par = self.right_text(text, font_name=font_name, font_size=font_size)
        par.runs[0].font.bold = bold
        return par

    @staticmethod
    def _chinese_date_text(value, *, year_month_only: bool = False):
        date = WordFormatter._date_value(value)
        digits = "零一二三四五六七八九"
        year = "".join(digits[int(char)] for char in str(date.year))

        def chinese_number(number):
            if number < 10:
                return digits[number]
            if number < 20:
                return "十" if number == 10 else "十" + digits[number % 10]
            tens, ones = divmod(number, 10)
            return digits[tens] + "十" + (digits[ones] if ones else "")

        text = f"{year}年{chinese_number(date.month)}月"
        if not year_month_only:
            text += f"{chinese_number(date.day)}日"
        return text

    def _chinese_date(self, value=None, *, year_month_only=False,
                      font_name: str = DEFAULT_CN_FONT, font_size=14,
                      bold: bool = False, alignment="居中"):
        par = self.doc.add_paragraph("")
        self.set_paragraph_format(par, alignment=alignment, first_line_indent_pt=None)
        run = par.add_run(self._chinese_date_text(
            value, year_month_only=year_month_only,
        ))
        self.set_run_font(run, cn_font=font_name, size=font_size, bold=bold)
        return par

    def chinese_date(self, value=None, *, font_name: str = DEFAULT_CN_FONT,
                     font_size=14, bold: bool = False, alignment="居中"):
        """添加中文完整日期，如“二零二六年七月十六日”。

        :param value: 日期值，省略时使用今天；支持 date/datetime 和标准日期字符串。
        :param font_name: 字体名称。
        :param font_size: 字号，支持中文字号字符串。
        :param bold: 是否加粗，默认否。
        :param alignment: 对齐方式，默认居中，支持中文、英文、单字母、数字和对齐枚举。
        """
        return self._chinese_date(
            value, font_name=font_name, font_size=font_size,
            bold=bold, alignment=alignment,
        )

    def chinese_year_month(self, value=None, *, font_name: str = DEFAULT_CN_FONT,
                           font_size=14, bold: bool = False, alignment="居中"):
        """添加中文年月，如“二零二六年七月”。

        :param value: 日期值，省略时使用今天；支持 date/datetime 和标准日期字符串。
        :param font_name: 字体名称。
        :param font_size: 字号，支持中文字号字符串。
        :param bold: 是否加粗，默认否。
        :param alignment: 对齐方式，默认居中，支持中文、英文、单字母、数字和对齐枚举。
        """
        return self._chinese_date(
            value, year_month_only=True, font_name=font_name,
            font_size=font_size, bold=bold, alignment=alignment,
        )

    # ================================================================
    # 3. 图片
    # ================================================================

    def insert_img(self, img_path: str, width: float, *, alignment="居左"):
        """插入图片。

        :param img_path: 图片文件路径。
        :param width: 图片宽度，单位 cm。
        :param alignment: 段落对齐方式，支持中文、英文、单字母、数字和对齐枚举。
        """
        par = self.doc.add_paragraph("")
        par.alignment = self.resolve_alignment(alignment)
        run = par.add_run()
        run.add_picture(img_path, width=Cm(width))
        return par

    # ================================================================
    # 4. 表格（新增能力）
    # ================================================================

    @staticmethod
    def set_table_borders(table, *, color: str = "000000", size: int = 4):
        """为表格添加四边及内部横竖边框。

        :param table: docx.table.Table 对象。
        :param color: 边框颜色，6 位十六进制 RGB 字符串，例如 ``"000000"``。
        :param size: 边框粗细，单位为 1/8 pt，例如 4 = 0.5pt、8 = 1pt。
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

        :param cell: docx.table._Cell 对象。
        :param text: 单元格文本。
        :param font_name: 字体名称。
        :param font_size: 字号，支持中文字号字符串。
        :param bold: 是否加粗，默认否。
        :param alignment: 对齐方式，默认居中，支持中文、英文、单字母、数字和对齐枚举。
        :param line_spacing: 行距倍数，默认 1.25。
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

    def add_table(self, headers, rows, *, col_widths=None, font_size=12,
                  header_bold: bool = True, alignment="居中",
                  border_color: str = "000000", border_size: int = 4):
        """一站式创建带表头的表格。

        :param headers: 表头文本列表，如 ["序号", "项目", "金额"]。
        :param rows: 二维数据列表，每个子列表为一行。
        :param col_widths: 每列宽度列表，单位 cm，可选。
        :param font_size: 表格字号，支持中文字号字符串。
        :param header_bold: 表头是否加粗，默认是。
        :param alignment: 单元格对齐方式，默认居中，支持中文、英文、单字母、数字和对齐枚举。
        :param border_color: 边框颜色，6 位十六进制 RGB 字符串。
        :param border_size: 边框粗细，单位 1/8 pt。
        :return: docx.table.Table 对象
        """
        table = self.doc.add_table(rows=1 + len(rows), cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        self.set_table_borders(table, color=border_color, size=border_size)

        for i, h in enumerate(headers):
            self.set_cell(
                table.rows[0].cells[i], h,
                font_size=font_size, bold=header_bold, alignment=alignment,
            )

        for r_idx, row in enumerate(rows, start=1):
            for c_idx, val in enumerate(row):
                self.set_cell(
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

        :param section: 目标节。
        :param text: 页眉文本。
        :param alignment: 对齐方式，默认居中，支持中文、英文、单字母、数字和对齐枚举。
        :param font_name: 字体名称。
        :param font_size: 字号，支持中文字号字符串。
        """
        return WordFormatter._fill_header_footer_part(
            section.header, text, alignment, font_name, font_size
        )

    @staticmethod
    def set_footer(section, text: str, *, alignment="居中",
                   font_name: str = DEFAULT_HEADER_FOOTER_FONT,
                   font_size=DEFAULT_HEADER_FOOTER_SIZE):
        """设置指定节的页脚文本（默认楷体小五号）。

        :param section: 目标节。
        :param text: 页脚文本。
        :param alignment: 对齐方式，默认居中，支持中文、英文、单字母、数字和对齐枚举。
        :param font_name: 字体名称。
        :param font_size: 字号，支持中文字号字符串。
        :note: 页码请使用 add_footer_page_number 或 set_page_number_from_section。
        """
        return WordFormatter._fill_header_footer_part(
            section.footer, text, alignment, font_name, font_size
        )

    @staticmethod
    def set_header_image(section, image_path, *, width: float | None = None,
                         height: float | None = None, alignment="居中"):
        """设置页眉图片，默认替换原页眉内容。

        :param section: 目标节。
        :param image_path: 图片路径，必须存在。
        :param width: 图片宽度，单位 cm，可选。
        :param height: 图片高度，单位 cm，可选；width 和 height 至少传一个。
        :param alignment: 对齐方式，默认居中，支持中文、英文、单字母、数字和对齐枚举。
        """
        return WordFormatter._set_header_footer_image(
            section.header, image_path, width=width, height=height,
            alignment=alignment,
        )

    @staticmethod
    def set_footer_image(section, image_path, *, width: float | None = None,
                         height: float | None = None, alignment="居中"):
        """设置页脚图片，默认替换原页脚内容。

        :param section: 目标节。
        :param image_path: 图片路径，必须存在。
        :param width: 图片宽度，单位 cm，可选。
        :param height: 图片高度，单位 cm，可选；width 和 height 至少传一个。
        :param alignment: 对齐方式，默认居中，支持中文、英文、单字母、数字和对齐枚举。
        """
        return WordFormatter._set_header_footer_image(
            section.footer, image_path, width=width, height=height,
            alignment=alignment,
        )

    @staticmethod
    def _set_header_footer_image(part, image_path, *, width=None, height=None,
                                 alignment="居中"):
        path = Path(image_path)
        if not path.is_file():
            raise FileNotFoundError(f"页眉页脚图片不存在: {path}")
        if width is None and height is None:
            raise ValueError("width 和 height 至少设置一个")

        part.is_linked_to_previous = False
        paragraph = part.paragraphs[0] if part.paragraphs else part.add_paragraph()
        for extra in part.paragraphs[1:]:
            extra._element.getparent().remove(extra._element)
        for run in list(paragraph.runs):
            run._r.getparent().remove(run._r)
        paragraph.alignment = WordFormatter.resolve_alignment(alignment, strict=True)
        paragraph.paragraph_format.line_spacing = 1
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        run = paragraph.add_run()
        kwargs = {}
        if width is not None:
            kwargs["width"] = Cm(width)
        if height is not None:
            kwargs["height"] = Cm(height)
        run.add_picture(str(path), **kwargs)
        return part

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
        """清空指定节的页脚内容。

        :param section: 目标节。
        """
        section.footer.is_linked_to_previous = False
        for paragraph in section.footer.paragraphs:
            for run in list(paragraph.runs):
                run._r.getparent().remove(run._r)
        return section

    @staticmethod
    def clear_header(section):
        """清空指定节的页眉内容。

        :param section: 目标节。
        """
        section.header.is_linked_to_previous = False
        for paragraph in section.header.paragraphs:
            for run in list(paragraph.runs):
                run._r.getparent().remove(run._r)
        return section

    # ================================================================
    # 6. 节与页码
    # ================================================================

    def insert_section(self, start_type=WD_SECTION_START.NEW_PAGE, *,
                       add_page_number: bool = False,
                       restart_page_number: bool = False,
                       start_page_number: int = 1,
                       prefix: str = "", suffix: str = "",
                       alignment="居中", font_name: str | None = None,
                       font_size=None):
        """插入新节。

        默认不添加页码；``add_page_number=True`` 时默认跟随上一节。
        ``restart_page_number=True`` 时从 ``start_page_number`` 重新编号。

        :param start_type: 分节方式，默认 NEW_PAGE；已有分页符时自动避免重复换页。
        :param add_page_number: 是否设置页码，默认否。
        :param restart_page_number: 是否重新开始编号，默认否。
        :param start_page_number: 重新编号的起始页码，默认 1。
        :param prefix: 页码前缀。
        :param suffix: 页码后缀。
        :param alignment: 页码对齐方式，默认居中，支持中文、英文、单字母、数字和对齐枚举。
        :param font_name: 页码字体名称。
        :param font_size: 页码字号，支持中文字号字符串。
        """
        # 上一段已经分页时，再使用 NEW_PAGE 会叠加两次换页，产生空白页。
        # 改用连续分节，既保留上一段的分页效果，又能创建新的节。
        effective_start_type = start_type
        if start_type == WD_SECTION_START.NEW_PAGE and self.doc.paragraphs:
            previous = self.doc.paragraphs[-1]
            has_page_break = bool(previous._p.xpath('.//w:br[@w:type="page"]'))
            if previous.paragraph_format.page_break_before or has_page_break:
                effective_start_type = WD_SECTION_START.CONTINUOUS

        section = self.doc.add_section(start_type=effective_start_type)
        if add_page_number:
            if restart_page_number:
                self.restart_page_numbering(
                    section, start=start_page_number, add_footer_number=True,
                    prefix=prefix, suffix=suffix, alignment=alignment,
                    font_name=font_name, font_size=font_size,
                )
            else:
                section.footer.is_linked_to_previous = True
        else:
            # 新节默认不继承页脚，避免上一节的页码或空段落带过来。
            self.clear_footer(section)
        return section

    @staticmethod
    def set_page_number_start(section, start: int = 1):
        """设置指定节页码起始值。

        :param section: 目标节。
        :param start: 起始页码。
        """
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

        :param paragraph: 目标段落，通常是页脚段落。
        :param prefix: 页码前文本，如 ``"第 "``。
        :param suffix: 页码后文本，如 ``" 页"``。
        :param alignment: 对齐方式，默认居中，支持中文、英文、单字母、数字和对齐枚举。
        :param font_name: 字体名称；不传时使用默认字体。
        :param font_size: 字号，支持中文字号字符串；不传时使用默认字号。
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
        """向指定节页脚添加页码，并复用现有首段。

        :param section: 目标节。
        :param prefix: 页码前文本。
        :param suffix: 页码后文本。
        :param alignment: 对齐方式，默认居中，支持中文、英文、单字母、数字和对齐枚举。
        :param font_name: 字体名称；不传时使用默认字体。
        :param font_size: 字号，支持中文字号字符串；不传时使用默认字号。
        """
        section.footer.is_linked_to_previous = False
        footer = section.footer
        paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        for extra in footer.paragraphs[1:]:
            extra._element.getparent().remove(extra._element)
        for run in list(paragraph.runs):
            run._r.getparent().remove(run._r)
        paragraph.paragraph_format.line_spacing = 1
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        return WordFormatter.add_page_number(
            paragraph, prefix=prefix, suffix=suffix, alignment=alignment,
            font_name=font_name, font_size=font_size,
        )

    @staticmethod
    def restart_page_numbering(section, *, start: int = 1, add_footer_number: bool = True,
                               prefix: str = "", suffix: str = "", alignment="居中",
                               font_name: str | None = None, font_size=None):
        """重启指定节的页码编号。

        :param section: 目标节。
        :param start: 起始页码，默认 1。
        :param add_footer_number: 是否同时添加页脚页码，默认是。
        :param prefix: 页码前文本。
        :param suffix: 页码后文本。
        :param alignment: 对齐方式，默认居中，支持中文、英文、单字母、数字和对齐枚举。
        :param font_name: 页码字体名称。
        :param font_size: 页码字号，支持中文字号字符串。
        """
        section.header.is_linked_to_previous = False
        section.footer.is_linked_to_previous = False
        WordFormatter.set_page_number_start(section, start=start)
        if add_footer_number:
            WordFormatter.add_footer_page_number(
                section, prefix=prefix, suffix=suffix, alignment=alignment,
                font_name=font_name, font_size=font_size,
            )
        return section

    def insert_section_with_page_numbering(self, *, start_type=WD_SECTION_START.NEW_PAGE,
                                           start_page_number: int = 1,
                                           add_footer_number: bool = True,
                                           prefix: str = "", suffix: str = "",
                                           alignment="居中", font_name: str | None = None,
                                           font_size=None):
        """插入新节并设置页码；默认从指定数字重新开始。

        :param start_type: 分节方式，默认 NEW_PAGE。
        :param start_page_number: 新节起始页码，默认 1。
        :param add_footer_number: 是否添加页脚页码，默认是。
        :param prefix/suffix: 页码前后文本。
        :param alignment: 页码对齐方式，支持中文、英文、单字母、数字和对齐枚举。
        :param font_name: 页码字体名称。
        :param font_size: 页码字号，支持中文字号字符串。
        """
        return self.insert_section(
            start_type=start_type, add_page_number=add_footer_number,
            restart_page_number=True, start_page_number=start_page_number,
            prefix=prefix, suffix=suffix, alignment=alignment,
            font_name=font_name, font_size=font_size,
        )

    def set_page_number_from_section(self, start_section_idx: int, *, start: int = 1,
                                    prefix: str = "", suffix: str = "",
                                    alignment="居中", font_name: str | None = None,
                                    font_size=None):
        """从指定节开始设置页码（包含此节），之前的节无页码。

        一次性完成：
        1) start_section_idx 之前的节：清空页脚，无页码
        2) start_section_idx 节：页码从 start 重新开始，添加页脚页码
        3) start_section_idx 之后的节：链接到前一节页脚，页码连续

        :param start_section_idx: 从 0 开始的节索引
        :param start: 起始页码
        :param prefix/suffix: 页码前后文本。
        :param alignment: 页码对齐方式，支持中文、英文、单字母、数字和对齐枚举。
        :param font_name: 页码字体名称。
        :param font_size: 页码字号，支持中文字号字符串。
        """
        sections = self.doc.sections
        total = len(sections)
        if start_section_idx < 0 or start_section_idx >= total:
            raise ValueError(
                f"start_section_idx {start_section_idx} 超出范围"
                f"（共 {total} 节，0~{total - 1}）"
            )

        for i in range(start_section_idx):
            self.clear_footer(sections[i])

        self.restart_page_numbering(
            sections[start_section_idx], start=start, add_footer_number=True,
            prefix=prefix, suffix=suffix, alignment=alignment,
            font_name=font_name, font_size=font_size,
        )

        for i in range(start_section_idx + 1, total):
            sections[i].footer.is_linked_to_previous = True
        return self.doc

    # ================================================================
    # 7. 目录
    # ================================================================

    def set_toc_level_style(self, level: int, *, font_name: str | None = None, font_size=None,
                            bold: bool | None = None, color: tuple[int, int, int] | None = None,
                            line_spacing=None, space_before=None, space_after=None,
                            first_line_indent=None, left_indent=None, alignment=None):
        """设置某一级 TOC 样式。

        :param level: 目录层级。
        :param font_name: 字体名称，可选。
        :param font_size: 字号，支持中文字号字符串，可选。
        :param bold: 是否加粗，可选。
        :param color: RGB 颜色元组，可选。
        :param line_spacing: 行距倍数，可选。
        :param space_before/space_after: 段前/段后空白，单位 pt，可选。
        :param first_line_indent: 首行缩进，单位 pt，可选。
        :param left_indent: 左缩进，单位 pt，可选。
        :param alignment: 对齐方式，可选。
        """
        style_name = f"TOC {level}"
        style = self._get_or_create_style(style_name, "Normal")
        return self._apply_style_options(
            style,
            {
                "font_name": font_name, "font_size": font_size, "bold": bold,
                "color": color, "line_spacing": line_spacing,
                "space_before": space_before, "space_after": space_after,
                "first_line_indent": first_line_indent, "left_indent": left_indent,
                "alignment": alignment,
            },
        )

    def set_paragraph_style(self, style_name: str, *, base_style_name: str = "Normal",
                            font_name: str | None = None, font_size=None,
                            bold: bool | None = None,
                            color: tuple[int, int, int] | None = None,
                            line_spacing=None, space_before=None, space_after=None,
                            first_line_indent=None, left_indent=None, alignment=None):
        """创建或更新自定义段落样式。

        :param style_name: 样式名称。
        :param base_style_name: 基础样式名称，默认 Normal。
        :param font_name: 字体名称，可选。
        :param font_size: 字号，支持中文字号字符串，可选。
        :param bold: 是否加粗，可选。
        :param color: RGB 颜色元组，可选。
        :param line_spacing: 行距倍数，可选。
        :param space_before/space_after: 段前/段后空白，单位 pt，可选。
        :param first_line_indent/left_indent: 缩进，单位 pt，可选。
        :param alignment: 对齐方式，可选。
        """
        style = self._get_or_create_style(style_name, base_style_name)
        return self._apply_style_options(
            style,
            {
                "font_name": font_name, "font_size": font_size, "bold": bold,
                "color": color, "line_spacing": line_spacing,
                "space_before": space_before, "space_after": space_after,
                "first_line_indent": first_line_indent, "left_indent": left_indent,
                "alignment": alignment,
            },
        )

    def add_custom_heading(self, text_content: str, *, style_name: str,
                          level: int | None = None, font_name: str | None = None,
                          font_size=None):
        """使用自定义段落样式添加标题。

        :param text_content: 标题文本。
        :param style_name: 使用的段落样式名称。
        :param level: 大纲层级，可选；传入后写入 Word 大纲级别。
        :param font_name: 字体名称，可选。
        :param font_size: 字号，支持中文字号字符串，可选。
        """
        par = self.doc.add_paragraph(text_content, style=style_name)
        if par.runs and (font_name is not None or font_size is not None):
            self.set_run_font(
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

    def add_toc(self, *, title: str = "目录", levels: tuple[int, int] = (1, 3),
                title_font_name: str = "黑体", title_font_size=16,
                title_bold: bool = True, title_alignment="居中",
                toc_level_styles: dict[int, dict] | None = None,
                use_hyperlinks: bool = True, hide_page_numbers_in_web: bool = True,
                use_outline_levels: bool = True,
                custom_style_levels: dict[str, int] | None = None,
                page_number_separator: str | None = None,
                right_align_page_numbers: bool | None = None):
        """插入 Word 目录域（TOC field）。

        生成的目录需在 Word/OnlyOffice 中刷新域（F9）才会显示页码。

        :param title: 目录标题，None 时不生成标题。
        :param levels: 收集的标题层级范围，如 ``(1, 3)``。
        :param title_font_name: 标题字体名称。
        :param title_font_size: 标题字号，支持中文字号字符串。
        :param title_bold: 目录标题是否加粗。
        :param title_alignment: 目录标题对齐方式。
        :param toc_level_styles: 各级 TOC 样式配置，格式为 ``{层级: 参数字典}``。
        :param use_hyperlinks: 是否生成超链接，默认是。
        :param hide_page_numbers_in_web: 网页模式是否隐藏页码，默认是。
        :param use_outline_levels: 是否使用大纲级别，默认是。
        :param custom_style_levels: 自定义样式到目录层级的映射。
        :param page_number_separator: 页码与标题之间的分隔符，可选。
        :param right_align_page_numbers: 是否右对齐页码；当前由 Word 目录域处理。
        """
        min_level, max_level = levels
        if min_level < 1 or max_level < min_level:
            raise ValueError("levels 参数不合法，例如 (1, 3)")

        self._ensure_update_fields_on_open()

        if title:
            title_par = self.doc.add_paragraph(title, style="Title")
            self.set_paragraph_format(
                title_par,
                alignment=self.resolve_alignment(title_alignment),
                space_before=12, space_after=12,
            )
            if title_par.runs:
                self.set_run_font(
                    title_par.runs[0], cn_font=title_font_name,
                    size=title_font_size, bold=title_bold,
                )

        if toc_level_styles:
            for level, options in toc_level_styles.items():
                self.set_toc_level_style(level, **options)

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

        toc_par = self.doc.add_paragraph("")
        self._append_field_run(toc_par, instruction)
        return toc_par

    '''REMOVED_COMPATIBILITY_BLOCK'''


about_word = WordFormatter
