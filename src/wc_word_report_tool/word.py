"""wc_word_report_tool.word
=========================
基于 python-docx 的中文 Word 报告格式化工具。

设计目标
--------
1. **覆盖完整**：表格、页眉页脚、目录、页码、封面等报告常见需求一站式覆盖。
2. **命名一致**：所有公开方法使用 snake_case，参数命名清晰，单位明确。
3. **默认合理**：默认值贴近中国公文标准（仿宋_GB2312、三号 14pt、1.5 倍行距）。
4. **严格校验**：未知参数抛 ValueError 而非静默降级，便于调试。

版本：v0.7.4
作者：willcha
"""

from __future__ import annotations

import datetime as _datetime
import io
from pathlib import Path

from docx.enum.style import WD_STYLE_TYPE
from docx.enum.section import WD_ORIENT, WD_SECTION_START
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import (
    WD_ALIGN_PARAGRAPH,
    WD_COLOR_INDEX,
    WD_PARAGRAPH_ALIGNMENT,
    WD_TAB_ALIGNMENT,
    WD_UNDERLINE,
)
from docx.image.image import Image as _DocxImage
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph


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

# 页码/编号格式 -> OOXML `w:pgNumType/@w:fmt` 取值。
# 中文报告的封面与目录常用罗马数字（Ⅰ/Ⅱ），正文才用阿拉伯数字。
PAGE_NUMBER_FORMATS = {
    # 阿拉伯数字
    "decimal": "decimal", "arabic": "decimal", "阿拉伯数字": "decimal",
    "数字": "decimal", "1": "decimal",
    # 罗马数字
    "upperRoman": "upperRoman", "大写罗马": "upperRoman", "I": "upperRoman",
    "lowerRoman": "lowerRoman", "小写罗马": "lowerRoman", "i": "lowerRoman",
    # 中文数字
    "chineseCounting": "chineseCounting", "中文数字": "chineseCounting",
    "chineseCountingThousand": "chineseCountingThousand",
    "中文数字千": "chineseCountingThousand", "一": "chineseCountingThousand",
    # 带圈数字 / 其他常见样式
    "decimalEnclosedCircle": "decimalEnclosedCircle",
    "带圈数字": "decimalEnclosedCircle", "①": "decimalEnclosedCircle",
    "decimalFullWidth": "decimalFullWidth", "全角数字": "decimalFullWidth",
    "decimalZero": "decimalZero", "前导零数字": "decimalZero",
    "upperLetter": "upperLetter", "大写字母": "upperLetter", "A": "upperLetter",
    "lowerLetter": "lowerLetter", "小写字母": "lowerLetter", "a": "lowerLetter",
}

# 纸张方向 -> (WD_ORIENT 枚举, OOXML `w:orient` 取值)
ORIENTATION_MAP = {
    "纵向": (WD_ORIENT.PORTRAIT, "portrait"),
    "竖排": (WD_ORIENT.PORTRAIT, "portrait"),
    "portrait": (WD_ORIENT.PORTRAIT, "portrait"),
    "P": (WD_ORIENT.PORTRAIT, "portrait"),
    "横向": (WD_ORIENT.LANDSCAPE, "landscape"),
    "横排": (WD_ORIENT.LANDSCAPE, "landscape"),
    "landscape": (WD_ORIENT.LANDSCAPE, "landscape"),
    "L": (WD_ORIENT.LANDSCAPE, "landscape"),
}

# 页眉页脚变体 -> python-docx Section 上的属性名
_HEADER_FOOTER_VARIANTS = {
    "primary": ("header", "footer"),
    "default": ("header", "footer"),
    "odd": ("header", "footer"),
    "first": ("first_page_header", "first_page_footer"),
    "首页": ("first_page_header", "first_page_footer"),
    "even": ("even_page_header", "even_page_footer"),
    "偶页": ("even_page_header", "even_page_footer"),
}


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

    命名约定（v0.4.6+）
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
                     highlight=None, underline: bool | str = False) -> None:
        """设置 run 的字体（中英文分别设置）、字号、加粗、颜色、高亮、下划线。

        :param run: python-docx 的 Run 对象
        :param cn_font: 中文字体名（写入 w:eastAsia）
        :param en_font: 西文字体名（写入 w:ascii / w:hAnsi）
        :param size: 字号，支持中文字号字符串，如 ``"三号"``、``"四号"``、``"小五"``。
        :param bold: 是否加粗
        :param color: RGB 三元组，如 (255, 0, 0)
        :param highlight: WD_COLOR_INDEX 高亮颜色
        :param underline: 下划线。``True`` 用单实线；也可传 ``"double"``/``"wave"`` 等
                          ``WD_UNDERLINE`` 的取值名。``False`` 表示不划线。
        """
        run.font.bold = bold
        run.font.size = Pt(_to_pt(size))
        run.font.name = en_font
        run.element.rPr.rFonts.set(qn("w:eastAsia"), cn_font)
        if color is not None:
            run.font.color.rgb = RGBColor(*color)
        if highlight is not None:
            run.font.highlight_color = highlight
        if underline:
            # 需要显式写 w:u，因为 Word 的默认值不是「无」，不写会继承样式里的设置
            style = underline if isinstance(underline, str) else "single"
            run.font.underline = WD_UNDERLINE[style.upper()]

    # 旧名保留（私有方法被外部依赖，保留为公开别名）

    @staticmethod
    def set_paragraph_format(paragraph, *, alignment=None, line_spacing=1.5,
                              line_spacing_pt: float | None = None,
                              space_before=0, space_after=0,
                              first_line_indent_pt=None,
                              first_line_indent_chars=None,
                              left_indent_pt: float | None = None,
                              strict: bool = False):
        """设置段落基础格式。

        :param alignment: 对齐方式，支持中文、英文、单字母、数字和对齐枚举。
        :param paragraph: 目标段落。
        :param line_spacing: 行距倍数。
        :param line_spacing_pt: 行距**固定值**，单位 pt；给了它就覆盖 ``line_spacing``。
            从 PDF 还原版式时用它：PDF 里只有「相邻两行基线差多少 pt」这一个事实，
            换算成倍数要依赖字体自身的行高（随字体而变），直接写 pt 才是精确的。
        :param space_before: 段前空白，单位 pt。
        :param space_after: 段后空白，单位 pt。
        :param first_line_indent_pt: 首行缩进兜底值，单位 pt；None 表示不设置。
        :param first_line_indent_chars: 首行缩进字符数；写入 ``w:firstLineChars``，
                                        优先级高于 ``w:firstLine``。
        :param left_indent_pt: 整段左缩进，单位 pt；None 表示不设置。用于还原
                               「整段右移」的表单行——它既不是居中，也不是首行缩进。
        :param strict: 是否对未知对齐方式抛出 ValueError，默认否。
        """
        if alignment is not None:
            paragraph.alignment = WordFormatter.resolve_alignment(alignment, strict=strict)
        # 赋值 Pt 对象时 python-docx 会写 lineRule=exact（固定值）；赋数字则是倍数
        paragraph.paragraph_format.line_spacing = (
            Pt(line_spacing_pt) if line_spacing_pt else line_spacing)
        paragraph.paragraph_format.space_before = Pt(space_before)
        paragraph.paragraph_format.space_after = Pt(space_after)
        if left_indent_pt is not None:
            paragraph.paragraph_format.left_indent = Pt(left_indent_pt)
        WordFormatter._set_first_line_indent(
            paragraph,
            first_line_indent_pt=first_line_indent_pt,
            first_line_indent_chars=first_line_indent_chars,
        )
        return paragraph

    @staticmethod
    def _set_first_line_indent(target, *, first_line_indent_pt=None,
                               first_line_indent_chars=None):
        """设置首行缩进，字符单位优先，pt 值作为兼容兜底。"""
        if first_line_indent_pt is None and first_line_indent_chars is None:
            return target

        if first_line_indent_pt is not None:
            target.paragraph_format.first_line_indent = Pt(first_line_indent_pt)

        p_pr = target._p.get_or_add_pPr() if hasattr(target, "_p") \
            else target.element.get_or_add_pPr()
        ind = p_pr.get_or_add_ind()
        if first_line_indent_chars is None:
            # 旧接口只传 pt 时恢复绝对缩进语义，避免历史字符值继续抢占优先级。
            ind.attrib.pop(qn("w:firstLineChars"), None)
            return target

        try:
            chars_value = float(first_line_indent_chars)
        except (TypeError, ValueError) as exc:
            raise TypeError("first_line_indent_chars 必须是数字") from exc
        if chars_value < 0:
            raise ValueError("first_line_indent_chars 不能小于 0")

        raw_value = round(chars_value * 100)
        # 首行缩进与悬挂缩进互斥，避免模板中的旧属性干扰字符缩进。
        ind.attrib.pop(qn("w:hanging"), None)
        ind.attrib.pop(qn("w:hangingChars"), None)
        ind.set(qn("w:firstLineChars"), str(raw_value))
        return target


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
    def resolve_page_number_format(number_format) -> str | None:
        """将页码格式参数解析为 OOXML ``w:pgNumType/@w:fmt`` 取值。

        支持 ``"upperRoman"``/``"lowerRoman"``/``"chineseCounting"`` 等 OOXML 原值，
        也支持中文别名（``"大写罗马"``/``"小写罗马"``/``"中文数字"``/``"带圈数字"``）。

        :param number_format: 页码格式；传 None 时返回 None（表示沿用上一节格式）。
        """
        if number_format is None:
            return None
        if isinstance(number_format, str):
            key = number_format.strip()
            resolved = PAGE_NUMBER_FORMATS.get(
                key, PAGE_NUMBER_FORMATS.get(key.lower())
            )
            if resolved is None:
                raise ValueError(
                    f"未知的页码格式: {number_format!r}。可选值: "
                    + "、".join(sorted(set(PAGE_NUMBER_FORMATS)))
                )
            return resolved
        raise TypeError(f"页码格式类型不支持: {type(number_format).__name__}")

    @staticmethod
    def resolve_orientation(orientation):
        """将纸张方向参数解析为 ``(WD_ORIENT 枚举, OOXML 取值)``。

        :param orientation: ``"纵向"``/``"横向"``/``"portrait"``/``"landscape"``/``"P"``/``"L"``。
        """
        if orientation is None:
            return None
        if isinstance(orientation, WD_ORIENT):
            ooxml = "landscape" if orientation == WD_ORIENT.LANDSCAPE else "portrait"
            return orientation, ooxml
        if isinstance(orientation, str):
            key = orientation.strip()
            resolved = ORIENTATION_MAP.get(key, ORIENTATION_MAP.get(key.lower()))
            if resolved is not None:
                return resolved
        raise ValueError(
            f"未知的纸张方向: {orientation!r}。支持 纵向/横向 或 portrait/landscape。"
        )

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
        first_line_indent_chars = options.get("first_line_indent_chars")
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
        if first_line_indent is not None or first_line_indent_chars is not None:
            WordFormatter._set_first_line_indent(
                style,
                first_line_indent_pt=first_line_indent,
                first_line_indent_chars=first_line_indent_chars,
            )
        if left_indent is not None:
            style.paragraph_format.left_indent = Pt(left_indent)
        if alignment is not None:
            style.paragraph_format.alignment = WordFormatter.resolve_alignment(alignment)
        return style

    def _configure_heading_style(self, level: int, *, cn_font: str, en_font: str, size, bold: bool,
                                 color: tuple[int, int, int] | None = None, line_spacing=1,
                                 space_before=12, space_after=12, first_line_indent=28,
                                 first_line_indent_chars=2):
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
                "first_line_indent_chars": first_line_indent_chars,
                "alignment": "左对齐",
            },
        )

    def _add_heading_with_style(self, heading_text: str, *, level: int, cn_font: str, en_font: str,
                                size, bold: bool, color: tuple[int, int, int] | None = None,
                                line_spacing=1, line_spacing_pt: float | None = None,
                                left_indent_pt: float | None = None,
                                space_before=12, space_after=12,
                                first_line_indent=28, first_line_indent_chars=2,
                                alignment=None):
        self._configure_heading_style(
            level, cn_font=cn_font, en_font=en_font, size=size, bold=bold, color=color,
            line_spacing=line_spacing, space_before=space_before, space_after=space_after,
            first_line_indent=first_line_indent,
            first_line_indent_chars=first_line_indent_chars,
        )
        heading = self.doc.add_heading("", level=level)
        run = heading.add_run(heading_text)
        self.set_run_font(run, cn_font=cn_font, en_font=en_font, size=size,
                                   bold=bold, color=color)
        # 对齐与段间距只写在段落上而非样式上：同一级的多个标题可能参数不同
        # （例如居中的封面大标题与居左的正文一级标题），写进样式会互相串味。
        heading.paragraph_format.space_before = Pt(space_before)
        heading.paragraph_format.space_after = Pt(space_after)
        if line_spacing_pt:
            # 行距固定值同样只写段落：样式里存的是倍数，混用会互相覆盖
            heading.paragraph_format.line_spacing = Pt(line_spacing_pt)
        if left_indent_pt:
            heading.paragraph_format.left_indent = Pt(left_indent_pt)
        if alignment is not None:
            heading.alignment = self.resolve_alignment(alignment, strict=True)
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
                       left: float = 2.8, right: float = 2.6, gutter: float = 0):
        """一次设置文档的默认页面参数。

        :param top/bottom/left/right: 页边距，单位 cm。
        :param gutter: 装订线，单位 cm。

        正文字体、字号、行距、缩进和对齐请通过 ``body()`` 或
        ``set_default_font()`` 单独设置。
        """
        self.set_page_margins(
            top=top, bottom=bottom, left=left, right=right, gutter=gutter,
        )
        return self.doc

    def set_document_language(self, lang: str = "zh-CN"):
        """设置 DOCX 默认校对语言，避免 Word/OnlyOffice 显示 English。

        :param lang: BCP 47 语言标签，默认 ``"zh-CN"``。
        """
        styles = self.doc.styles.element
        for lang_element in styles.xpath(".//w:lang"):
            lang_element.set(qn("w:val"), lang)
            lang_element.set(qn("w:eastAsia"), lang)
            lang_element.set(qn("w:bidi"), lang)

        settings = self.doc.settings.element
        theme_font_lang = settings.find(qn("w:themeFontLang"))
        if theme_font_lang is None:
            theme_font_lang = OxmlElement("w:themeFontLang")
            settings.insert(0, theme_font_lang)
        theme_font_lang.set(qn("w:val"), lang)
        theme_font_lang.set(qn("w:eastAsia"), lang)
        theme_font_lang.set(qn("w:bidi"), lang)
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
             underline: bool | str = False,
             alignment="两端对齐", line_spacing=1.5, line_spacing_pt: float = 0,
             left_indent_pt: float = 0,
             space_before=0, space_after=0, page_break_before: bool = False):
        """添加正文段落。

        默认格式：宋体、四号、首行缩进 2 字符、两端对齐、1.5 倍行距。

        :param text: 正文文本。
        :param font_name: 中文字体名称，默认 ``"宋体"``；可传任意本机已安装字体。
        :param font_size: 字号，默认 ``"四号"``。支持数字/浮点数（单位 pt）、
                          数字字符串，以及 ``FONT_SIZE_MAP`` 中的中文字号，
                          例如 ``"三号"``、``"四号"``、``"小五"``。
        :param indent: 是否首行缩进 2 字符，默认 ``True``。缩进值按当前字号
                       乘以 2 生成 ``w:firstLine`` 兼容值，同时优先写入
                       ``w:firstLineChars=\"200\"``。
        :param bold: 是否加粗，默认 ``False``。
        :param highlight: 是否使用黄色高亮，默认 ``False``。
        :param underline: 整段是否加下划线，默认 ``False``；``True`` 为单实线，
                          也可传 ``"double"``/``"wave"`` 等 ``WD_UNDERLINE`` 取值名。
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
        :param page_break_before: 是否在本段之前分页，默认 ``False``。
            优先用它而不是单独插分页符：该属性在段落已经处于页首时自动失效，
            不会像独立分页段落那样多顶出一张空白页。
        :return: 新创建的 ``Paragraph`` 对象。
        """
        par = self.doc.add_paragraph("")
        self.set_paragraph_format(
            par,
            alignment=self.resolve_alignment(alignment),
            line_spacing=line_spacing,
            line_spacing_pt=line_spacing_pt or None,
            left_indent_pt=left_indent_pt or None,
            space_before=space_before,
            space_after=space_after,
            first_line_indent_pt=_to_pt(font_size) * 2 if indent else None,
            first_line_indent_chars=2 if indent else None,
        )
        if page_break_before:
            par.paragraph_format.page_break_before = True
        run = par.add_run(text)
        self.set_run_font(
            run, cn_font=font_name, size=font_size, color=(0, 0, 0),
            bold=bold, underline=underline,
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

    def add_page_break(self, text: str = ""):
        """插入一个分页符（可同时带一段文字）。

        模板类文档常常「一个表单起一页」，把这些分页还原出来才能对上原稿的页序；
        只靠内容自然流排的话，页数和每页内容都会错位。

        :param text: 分页符之后要写的文字，默认空（纯分页）。
        :return: 承载分页符的 ``Paragraph``。
        """
        paragraph = self.doc.add_paragraph("")
        run = paragraph.add_run()
        br = OxmlElement("w:br")
        br.set(qn("w:type"), "page")
        run._r.append(br)
        if text:
            text_run = paragraph.add_run(text)
            self.set_run_font(
                text_run, cn_font=DEFAULT_CN_FONT, en_font=DEFAULT_EN_FONT,
                size=DEFAULT_BODY_SIZE,
            )
        return paragraph

    def body_segments(self, segments, *, font_name: str = "宋体", font_size="四号",
                      indent: bool = True, alignment="两端对齐", line_spacing=1.5,
                      line_spacing_pt: float = 0, left_indent_pt: float = 0,
                      space_before=0, space_after=0, page_break_before: bool = False):
        """在一段里写入多个「片段」，每段可单独高亮或加下划线。

        用于「标签 + 填值」这类只需要对其中一部分做标记的场景：
        ``[("项目名称：", False), ("雄安新区示范工程", True)]`` 只会把填进去的值标黄，
        标签保持原样；整段高亮会把标签也涂黄，看起来分不清哪些是 AI 填的。

        下划线同理：招投标模板里「填空线」只画在填进去的值下面，占位标签不划线，
        所以下划线也必须是**逐片段**的。

        :param segments: ``[(文本, 是否高亮)]`` 序列；也可传
                         ``{"text":..., "highlight":..., "underline":...}``。
                         字典形式才支持逐片段下划线。
        :param font_name: 中文字体名称。
        :param font_size: 字号，支持中文字号字符串。
        :param indent: 是否首行缩进 2 字符。
        :param alignment: 段落对齐方式。
        :param line_spacing: 行距倍数。
        :param space_before: 段前空白，单位 pt。
        :param space_after: 段后空白，单位 pt。
        :param page_break_before: 是否在本段之前分页。
        :return: 新创建的 ``Paragraph`` 对象。
        """
        par = self.doc.add_paragraph("")
        self.set_paragraph_format(
            par,
            alignment=self.resolve_alignment(alignment),
            line_spacing=line_spacing,
            line_spacing_pt=line_spacing_pt or None,
            left_indent_pt=left_indent_pt or None,
            space_before=space_before,
            space_after=space_after,
            first_line_indent_pt=_to_pt(font_size) * 2 if indent else None,
            first_line_indent_chars=2 if indent else None,
        )
        if page_break_before:
            par.paragraph_format.page_break_before = True

        for segment in segments or []:
            if isinstance(segment, dict):
                text = segment.get("text", "")
                highlight = bool(segment.get("highlight"))
                underline = segment.get("underline", False)
            else:
                text, highlight = segment[0], bool(segment[1])
                underline = False
            if not text:
                continue
            run = par.add_run(str(text))
            self.set_run_font(
                run, cn_font=font_name, size=font_size, color=(0, 0, 0),
                underline=underline,
                highlight=WD_COLOR_INDEX.YELLOW if highlight else None,
            )
        return par

    def heading(self, text: str, *, level: int = 1,
                font_name: str = "黑体", font_size="三号",
                bold: bool = False,
                indent: bool = False, line_spacing=1.5, line_spacing_pt: float = 0,
                left_indent_pt: float = 0,
                alignment=None,
                 space_before=None, space_after=None,
                page_break_before: bool = False):
        """添加通用标题，默认一级标题格式。

        :param text: 标题文本。
        :param level: 标题层级，1 到 9，默认 1。
        :param font_name: 标题字体名称。
        :param font_size: 标题字号，支持中文字号字符串。
        :param bold: 是否加粗，默认否。
        :param indent: 是否首行缩进 2 字符，默认否。
        :param line_spacing: 行距倍数，默认 1.5。
        :param line_spacing_pt: 行距固定值，单位 pt；给了它覆盖 ``line_spacing``。
        :param alignment: 对齐方式，默认不设置（沿用标题样式的居左）。
                          传 ``"居中"`` 可做居中的大标题，支持中文、英文、单字母、数字和对齐枚举。
        :param space_before: 段前空白，单位 pt；不传时按字号的 0.5 行计算。
        :param space_after: 段后空白，单位 pt；不传时按字号的 0.5 行计算。
        :param page_break_before: 是否在本标题之前分页，默认 ``False``。
            该属性在标题已处于页首时自动失效，不会多顶出空白页。
        """
        if level < 1 or level > 9:
            raise ValueError("标题 level 必须在 1 到 9 之间")
        paragraph = self._add_heading_with_style(
            text, level=level, cn_font=font_name, en_font=DEFAULT_EN_FONT,
            size=font_size, bold=bold, line_spacing=line_spacing,
            line_spacing_pt=line_spacing_pt or None,
            left_indent_pt=left_indent_pt or None,
            space_before=_to_pt(font_size) * 0.5 if space_before is None else space_before,
            space_after=_to_pt(font_size) * 0.5 if space_after is None else space_after,
            first_line_indent=_to_pt(font_size) * 2 if indent else None,
            first_line_indent_chars=2 if indent else None, color=(0, 0, 0),
            alignment=alignment,
        )
        if page_break_before:
            paragraph.paragraph_format.page_break_before = True
        return paragraph

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
            first_line_indent_chars=2 if indent else None,
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
            first_line_indent_pt=_to_pt(font_size) * 2 if indent else None,
            first_line_indent_chars=2 if indent else None,
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

    @staticmethod
    def _image_stream(image):
        """把路径或 bytes 统一成 ``add_picture`` 能接受的输入。"""
        if isinstance(image, (bytes, bytearray)):
            return io.BytesIO(bytes(image)), True
        if hasattr(image, "read"):
            return image, False
        path = Path(image)
        if not path.is_file():
            raise FileNotFoundError(f"图片不存在: {path}")
        return str(path), False

    @staticmethod
    def _native_width_cm(image) -> float:
        """按图片自身的 DPI 计算原始宽度，单位 cm。

        扫描件裁剪出来的图通常没有「应该多宽」的先验，按原始像素 / DPI 换算最接近
        它在源文档里的物理尺寸，比拍一个默认值可靠。
        """
        stream = image if hasattr(image, "read") else open(str(image), "rb")
        try:
            probe = _DocxImage.from_file(stream)
        finally:
            if hasattr(image, "read"):
                stream.seek(0)
            else:
                stream.close()
        dpi = probe.horz_dpi or 96
        return probe.px_width / dpi * 2.54

    def insert_img(self, img_path, width=None, *, height=None, alignment="居左",
                   floating: bool = False, y_offset_pt: float = 0):
        """插入图片。

        ``width`` 与 ``height`` 都不传时，按图片自身的 DPI 使用原始尺寸，
        因此 OCR/PDF 裁切出来的图片可以直接落进来，无需先猜一个宽度。

        :param img_path: 图片文件路径（str/Path），或图片的二进制内容（bytes/bytearray）。
        :param width: 图片宽度，单位 cm，可选。
        :param height: 图片高度，单位 cm，可选。
        :param alignment: 段落对齐方式，支持中文、英文、单字母、数字和对齐枚举。
        :param floating: 是否让图片浮于文字上方。印章需要压在落款文字上时用它，
                         PDF 里裁出的图表、公式也一样。
        :param y_offset_pt: 浮动图片的垂直偏移，单位 pt，正数上移、负数下移。
        """
        if width is not None and width <= 0:
            raise ValueError(f"图片宽度必须大于 0，收到 {width!r}")
        if height is not None and height <= 0:
            raise ValueError(f"图片高度必须大于 0，收到 {height!r}")

        par = self.doc.add_paragraph("")
        par.alignment = self.resolve_alignment(alignment)
        run = par.add_run()
        source, owned = self._image_stream(img_path)
        try:
            kwargs = {}
            if width is not None:
                kwargs["width"] = Cm(width)
            if height is not None:
                kwargs["height"] = Cm(height)
            if not kwargs:
                kwargs["width"] = Cm(self._native_width_cm(source))
            run.add_picture(source, **kwargs)
        finally:
            if owned:
                source.close()
        if floating:
            self._make_run_floating(run, y_offset_pt=y_offset_pt)
        return par

    # ================================================================
    # 3.1 结算报告封面与签发页
    # ================================================================

    def fengmian_jiesuan(self, logo, project_name: str = "默认工程名称",
                         entrusting_unit: str = "默认委托单位", *,
                         compiling_unit: str = "默认编制单位",
                         report_title: str = "结算审核报告", date=None,
                         logo_width: float = 4, info_blank_lines: int = 11):
        """增加结算审核报告封面。

        :param logo: 封面 Logo 图片路径。
        :param project_name: 工程名称，默认 ``"默认工程名称"``。
        :param entrusting_unit: 委托单位名称，默认 ``"默认委托单位"``。
        :param compiling_unit: 编制单位名称，默认 ``"默认编制单位"``。
        :param report_title: 报告标题，默认“结算审核报告”。
        :param date: 封面日期；不传时使用当天，支持 date/datetime 或标准日期字符串。
        :param logo_width: Logo 宽度，单位 cm。
        :param info_blank_lines: 报告标题与落款表格之间的空行数。
        :return: 封面落款表格。
        """
        self.insert_img(str(logo), logo_width)
        self.blank_lines(1)
        self.body(
            project_name, font_name="黑体", font_size=20,
            alignment="居中", indent=False,
        )
        self.body(
            report_title, font_name="宋体", font_size=24,
            alignment="居中", bold=True, indent=False,
        )
        self.blank_lines(info_blank_lines)

        table = self.doc.add_table(rows=2, cols=3)
        self.edit_cell(table, "委托单位：", 0, 1, bold=True)
        self.edit_cell(table, entrusting_unit, 0, 2, bold=True)
        self.edit_cell(table, "编制单位：", 1, 1, bold=True)
        self.edit_cell(table, compiling_unit, 1, 2, bold=True)
        for row in table.rows:
            row.cells[0].width = Cm(1.5)
            row.cells[1].width = Cm(4)
            row.cells[2].width = Cm(10)

        self.chinese_year_month(
            date, font_size="三号", font_name="宋体", bold=True,
        )
        return table

    def qianfaye(self, logo, project_name: str = "默认工程名称",
                 entrusting_unit: str = "默认委托单位",
                 personnel: dict | None = None, *, participants=None,
                 compiling_unit: str = "默认编制单位",
                 report_title: str = "结算审核报告",
                 footer_text: str = "公司从业方针：客观、公正、严谨、专业",
                 logo_width: float = 2,
                 logo_y_offset_pt: float = 0):
        """新建一节并增加结算审核报告签发页。

        ``personnel`` 可包含“公司签发、部门核准、部门审核、小组初审、项目负责人”；
        每项可填写“姓名、部门、职务、职称”，项目负责人还可填写“联系电话”。

        :param logo: 页眉 Logo 图片路径。
        :param project_name: 工程名称，默认 ``"默认工程名称"``。
        :param entrusting_unit: 委托单位名称，默认 ``"默认委托单位"``。
        :param personnel: 各签发岗位的人员信息字典，默认空字典。
        :param participants: 项目参与者列表，每项可填写“姓名、职称”。
        :param compiling_unit: 报告编制单位，默认 ``"默认编制单位"``。
        :param report_title: 页眉中的报告标题。
        :param footer_text: 页脚文字。
        :param logo_width: 页眉 Logo 宽度，单位 cm。
        :param logo_y_offset_pt: 页眉 Logo 垂直偏移，单位 pt；正数上移，负数下移。
        :return: 签发页表格。
        """
        personnel = personnel or {}
        participants = participants or []
        section = self.insert_section(
            add_page_number=False, inherit_header=False, inherit_footer=False,
        )
        self.set_header_image(
            section, logo, width=logo_width, alignment="居左",
            y_offset_pt=logo_y_offset_pt,
        )
        self.set_header(
            section, f"{project_name}{report_title}", font_name="宋体",
            font_size="小五", bottom_border=True,
        )
        self.set_footer(
            section, footer_text, font_name="楷体_GB2312", font_size="小五",
            alignment="居左", line_length=6,
        )

        table = self.doc.add_table(rows=10 + len(participants), cols=4)
        table.cell(0, 0).merge(table.cell(0, 3))
        self.edit_cell(table, f"委托单位：{entrusting_unit}", 0, 0, bold=True)
        table.cell(1, 0).merge(table.cell(1, 3))
        self.edit_cell(table, f"报告编制单位：{compiling_unit}", 1, 0, bold=True)

        roles = (
            ("公司签发", 3), ("部门核准", 4), ("部门审核", 5),
            ("小组初审", 6), ("项目负责人", 7),
        )
        for role, row_index in roles:
            data = personnel.get(role, {})
            department_and_position = [data.get("部门", ""), data.get("职务", "")]
            if role == "公司签发":
                department_and_position[0] = ""
            self.edit_cell(table, f"{role}：", row_index, 0)
            self.edit_cell(table, data.get("姓名", ""), row_index, 1)
            self.edit_cell(table, department_and_position, row_index, 2)
            self.edit_cell(table, data.get("职称", ""), row_index, 3)

        self.edit_cell(table, "签字：", 8, 0)
        self.edit_cell(table, "盖章", 8, 2)
        table.rows[8].height = Cm(1.5)
        table.rows[8].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

        table.cell(9, 2).merge(table.cell(9, 3))
        project_manager = personnel.get("项目负责人", {})
        self.edit_cell(
            table, f"联系电话：{project_manager.get('联系电话', '')}", 9, 2,
        )

        for index, data in enumerate(participants, start=10):
            if index == 10:
                self.edit_cell(table, "项目参与者：", index, 0)
            self.edit_cell(table, data.get("姓名", ""), index, 1)
            self.edit_cell(table, data.get("职称", ""), index, 3)

        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for row in table.rows:
            row.cells[0].width = Cm(4)
            row.cells[3].width = Cm(5)
        return table

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
                 bold: bool = False, alignment="居中", line_spacing=1.25,
                 highlight: bool = False):
        """设置表格单元格内容与格式（增强版，支持加粗、对齐与高亮）。

        :param cell: docx.table._Cell 对象。
        :param text: 单元格文本。
        :param font_name: 字体名称。
        :param font_size: 字号，支持中文字号字符串。
        :param bold: 是否加粗，默认否。
        :param alignment: 对齐方式，默认居中，支持中文、英文、单字母、数字和对齐枚举。
        :param line_spacing: 行距倍数，默认 1.25。
        :param highlight: 是否黄色高亮，默认否。用于标出「由 AI 填写的临时数据」，
                          方便人工复核。
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
        WordFormatter.set_run_font(
            run, cn_font=font_name, size=font_size, bold=bold,
            highlight=WD_COLOR_INDEX.YELLOW if highlight else None,
        )
        return cell

    @staticmethod
    def edit_cell(table, context: str | list[str], row: int, col: int, *,
                  font_name: str = "宋体", font_size="小三",
                  bold: bool = False):
        """编辑表格指定单元格，支持一个单元格写入多行内容。

        :param table: 目标表格。
        :param context: 单元格内容，可以是字符串或字符串列表；列表项逐行写入。
        :param row: 行索引，从 0 开始。
        :param col: 列索引，从 0 开始。
        :param font_name: 中文字体名称。
        :param font_size: 字号，支持中文字号字符串。
        :param bold: 是否加粗，默认否。
        :return: 编辑后的单元格。
        """
        cell = table.cell(row, col)
        paragraph = cell.paragraphs[0]
        paragraph.clear()
        texts = context if isinstance(context, list) else [context]

        for index, text in enumerate(texts):
            run = paragraph.add_run(str(text))
            WordFormatter.set_run_font(
                run, cn_font=font_name, en_font="Times New Roman",
                size=font_size, bold=bold,
            )
            if index < len(texts) - 1:
                run.add_break()
        return cell

    @staticmethod
    def set_table_columns(table, col_widths, *, fixed_layout: bool = True):
        """设置表格的列宽，并让 Word **真正按这个宽度渲染**。

        只写 ``w:tcW`` 是不够的：python-docx 建表时写的是 ``w:tblW type="auto"``
        且不设 ``w:tblLayout``，Word 会按自动布局重新分配列宽，把手工算好的列宽冲掉。
        因此这里一次写全三处，缺一不可：

        * ``w:tblGrid/w:gridCol`` —— Word 计算列布局的基准网格；
        * 每个单元格的 ``w:tcW`` —— 实际列宽（合并单元格按其跨列数写合计值）；
        * ``w:tblW`` + ``w:tblLayout type="fixed"`` —— 声明固定布局，禁止自动重排。

        :param table: docx.table.Table 对象。
        :param col_widths: 每列宽度列表，单位 cm，长度等于表格列数。
        :param fixed_layout: 是否锁定为固定布局，默认是。
        """
        widths = [float(w) for w in (col_widths or [])]
        if not widths:
            raise ValueError("col_widths 不能为空")
        if any(w <= 0 for w in widths):
            raise ValueError(f"列宽必须全部大于 0，收到 {col_widths!r}")

        tbl = table._tbl
        grid = tbl.find(qn("w:tblGrid"))
        if grid is None:
            grid = OxmlElement("w:tblGrid")
            tbl.insert(list(tbl).index(tbl.tblPr) + 1, grid)
        for col in list(grid.findall(qn("w:gridCol"))):
            grid.remove(col)
        for width in widths:
            col = OxmlElement("w:gridCol")
            col.set(qn("w:w"), str(int(round(width * 566.929))))   # cm -> twips
            grid.append(col)

        # 逐单元格写 tcW：按真实 tc 元素遍历（不能用 row.cells——合并后它会重复返回
        # 同一个锚点单元格，跨列数会被重复计算），合并单元格写它跨越列的合计宽度。
        from docx.table import _Cell
        for tr in tbl.tr_lst:
            grid_col = 0
            for tc in tr.tc_lst:
                if grid_col >= len(widths):
                    break
                span = tc.grid_span
                total = sum(widths[grid_col:grid_col + span]) or widths[grid_col]
                _Cell(tc, table).width = Cm(total)
                grid_col += span

        WordFormatter.set_table_width(table, sum(widths))
        if fixed_layout:
            tbl_pr = tbl.tblPr
            layout = tbl_pr.find(qn("w:tblLayout"))
            if layout is None:
                layout = OxmlElement("w:tblLayout")
                tbl_pr.append(layout)
            layout.set(qn("w:type"), "fixed")
        return table

    @staticmethod
    def set_table_width(table, width_cm: float):
        """设置表格总宽度（``w:tblW``）。

        :param table: docx.table.Table 对象。
        :param width_cm: 表格总宽度，单位 cm。
        """
        if width_cm <= 0:
            raise ValueError(f"表格宽度必须大于 0，收到 {width_cm!r}")
        tbl_pr = table._tbl.tblPr
        tbl_w = tbl_pr.find(qn("w:tblW"))
        if tbl_w is None:
            tbl_w = OxmlElement("w:tblW")
            # tblPr 是顺序敏感的：tblW 必须排在 jc / tblBorders 等元素之前
            anchor = None
            for tag in ("w:jc", "w:tblCellSpacing", "w:tblInd", "w:tblBorders",
                        "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook"):
                found = tbl_pr.find(qn(tag))
                if found is not None:
                    anchor = found
                    break
            if anchor is not None:
                anchor.addprevious(tbl_w)
            else:
                tbl_pr.append(tbl_w)
        tbl_w.set(qn("w:w"), str(int(round(width_cm * 567))))   # cm -> twips
        tbl_w.set(qn("w:type"), "dxa")
        return table

    @staticmethod
    def _normalize_table_merges(merges, row_count: int, col_count: int):
        """把合并描述归一成 ``(r1, c1, r2, c2)`` 列表，并校验越界与重叠。

        接受 ``{"row": 0, "col": 0, "rowspan": 2, "colspan": 3}`` 形式的字典，
        也接受 ``(row, col, rowspan, colspan)`` 形式的元组/列表。
        """
        regions = []
        claimed = {}
        for item in merges or []:
            if isinstance(item, dict):
                row = int(item.get("row", 0))
                col = int(item.get("col", 0))
                rowspan = int(item.get("rowspan", 1))
                colspan = int(item.get("colspan", 1))
            else:
                values = list(item)
                if len(values) < 4:
                    values = list(values) + [1] * (4 - len(values))
                row, col, rowspan, colspan = (int(v) for v in values[:4])

            if rowspan < 1 or colspan < 1:
                raise ValueError(f"合并区域的 rowspan/colspan 必须 >= 1，收到 {item!r}")
            r2, c2 = row + rowspan - 1, col + colspan - 1
            if row < 0 or col < 0 or r2 >= row_count or c2 >= col_count:
                raise ValueError(
                    f"合并区域 {item!r} 超出表格范围（{row_count} 行 × {col_count} 列）。"
                )
            if rowspan == 1 and colspan == 1:
                continue
            for r in range(row, r2 + 1):
                for c in range(col, c2 + 1):
                    if (r, c) in claimed and claimed[(r, c)] != (row, col):
                        raise ValueError(
                            f"合并区域 {item!r} 与 {(claimed[(r, c)][0], claimed[(r, c)][1])} "
                            f"覆盖了同一个单元格 ({r}, {c})。"
                        )
                    claimed[(r, c)] = (row, col)
            regions.append((row, col, r2, c2))
        return regions

    @staticmethod
    def merge_cells(table, merges):
        """合并表格中指定的单元格区域（对已存在的表格按索引合并）。

        与 :meth:`add_table` 的 ``merges`` 语义一致：**只保留区域左上角单元格的内容**。
        python-docx 的 ``merge()`` 会把被并单元格的文本拼接进来，这里在合并后把
        多出来的段落删掉，避免出现「A\\nB\\nC」这种意外结果。

        :param table: docx.table.Table 对象。
        :param merges: 合并描述列表，如 ``[{"row": 0, "col": 0, "rowspan": 2, "colspan": 3}]``
                       或 ``[(0, 0, 2, 3)]``。
        """
        row_count, col_count = len(table.rows), len(table.columns)
        regions = WordFormatter._normalize_table_merges(merges, row_count, col_count)
        for r1, c1, r2, c2 in regions:
            anchor = table.cell(r1, c1)
            keep = len(anchor.paragraphs)
            anchor.merge(table.cell(r2, c2))
            WordFormatter._strip_extra_paragraphs(table.cell(r1, c1), keep=keep)
        return table

    @staticmethod
    def _pick_from_grid(grid, row: int, col: int, default=None):
        """从二维配置里安全取值；越界或该位置为空时回落到默认值。"""
        if not grid or row >= len(grid):
            return default
        line = grid[row]
        if line is None or col >= len(line):
            return default
        value = line[col]
        return default if value is None else value

    @staticmethod
    def _strip_extra_paragraphs(cell, keep: int = 1):
        """删除单元格中多余的空段落。

        合并单元格时，被并进来的单元格各自带着一个空段落，不清理会让单元格白高一截。
        """
        for paragraph in cell.paragraphs[keep:]:
            paragraph._p.getparent().remove(paragraph._p)
        return cell

    @staticmethod
    def _validate_table_grid(headers, rows, col_widths=None, row_heights=None):
        """校验表格参数自洽，把「列数对不上」变成可读的错误。

        这里刻意不做静默修正：``headers`` / ``rows`` / ``col_widths`` 的长度必须相互自洽，
        否则表格会整体错列。而越界时 python-docx 只会抛
        ``IndexError: tuple index out of range``——对调用方毫无指向性，
        实测 OCR 产出「表头 1 列、数据行 4 列、列宽 4 个」时排查了很久。
        """
        if not headers:
            raise ValueError("headers 不能为空，至少需要一列表头。")
        col_count = len(headers)
        if col_widths is not None and len(col_widths) > col_count:
            raise ValueError(
                f"col_widths 有 {len(col_widths)} 项，超过表头的 {col_count} 列"
                f"（表头：{list(headers)}）。请让列宽数量与表头列数一致。"
            )
        for index, row in enumerate(rows):
            if len(row) > col_count:
                raise ValueError(
                    f"第 {index + 1} 行数据有 {len(row)} 个单元格，超过表头的 {col_count} 列"
                    f"（表头：{list(headers)}）。每行的单元格数不能多于表头列数。"
                )
        row_count = 1 + len(rows)
        if row_heights is not None and len(row_heights) > row_count:
            raise ValueError(
                f"row_heights 有 {len(row_heights)} 项，超过表格的 {row_count} 行"
                f"（1 行表头 + {len(rows)} 行数据）。"
            )

    def add_table(self, headers, rows, *, col_widths=None, font_size=12,
                  header_bold: bool = True, alignment="居中",
                  font_name: str = DEFAULT_CN_FONT,
                  border_color: str = "000000", border_size: int = 4,
                  merges=None, row_heights=None, table_width_cm=None,
                  cell_font_names=None, cell_highlights=None):
        """一站式创建带表头的表格。

        ``headers`` 与 ``rows`` 组成**矩形网格**：每行的单元格数不能多于表头列数，
        ``col_widths`` 也不能长于列数；长度不齐会直接抛 ``ValueError`` 并指出是第几行，
        不做静默修正（静默修正会让表格整体错列）。

        :param headers: 表头文本列表，如 ["序号", "项目", "金额"]。
        :param rows: 二维数据列表，每个子列表为一行。
        :param col_widths: 每列宽度列表，单位 cm，可选。
        :param font_size: 表格字号，支持中文字号字符串。
        :param header_bold: 表头是否加粗，默认是。
        :param alignment: 单元格对齐方式，默认居中，支持中文、英文、单字母、数字和对齐枚举。
        :param font_name: 单元格中文字体名，默认仿宋_GB2312。
        :param border_color: 边框颜色，6 位十六进制 RGB 字符串。
        :param border_size: 边框粗细，单位 1/8 pt。
        :param merges: 合并区域列表，如 ``[{"row": 0, "col": 0, "colspan": 4}]``；
                       合并区域的内容取该区域**左上角**单元格的值。
        :param row_heights: 每行高度，单位 cm，可选；按「最小值」规则设置，内容多时自动撑高。
        :param table_width_cm: 表格总宽度，单位 cm，可选。
        :param cell_font_names: 逐单元格中文字体名的二维列表（含表头那一行），可选；
                                未覆盖的位置回落到 ``font_name``。
        :param cell_highlights: 逐单元格「是否黄色高亮」的二维布尔列表（含表头那一行），
                                可选。用来标出由 AI 填写的临时数据，便于人工复核。
        :return: docx.table.Table 对象
        """
        row_count = 1 + len(rows)
        col_count = len(headers)
        self._validate_table_grid(headers, rows, col_widths, row_heights)
        table = self.doc.add_table(rows=row_count, cols=col_count)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        self.set_table_borders(table, color=border_color, size=border_size)
        if table_width_cm is not None:
            self.set_table_width(table, table_width_cm)

        regions = self._normalize_table_merges(merges, row_count, col_count)

        # 先合并再填内容：反过来会被 python-docx 把被并单元格的文本一并拼接进合并格
        for r1, c1, r2, c2 in regions:
            table.cell(r1, c1).merge(table.cell(r2, c2))

        # 列宽放在合并之后：合并单元格的 tcW 要写成它跨越列的合计宽度，
        # 而「跨了几列」只有合并完成、gridSpan 落定之后才知道。
        if col_widths:
            self.set_table_columns(table, col_widths)

        covered = set()
        for r1, c1, r2, c2 in regions:
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    if (r, c) != (r1, c1):
                        covered.add((r, c))

        def write(row: int, col: int, value, bold: bool):
            if (row, col) in covered:
                return
            cell = table.cell(row, col)
            self.set_cell(
                cell, str(value),
                font_name=self._pick_from_grid(cell_font_names, row, col, font_name),
                font_size=font_size, bold=bold, alignment=alignment,
                highlight=bool(self._pick_from_grid(cell_highlights, row, col, False)),
            )
            if (row, col) in {(r1, c1) for r1, c1, _, _ in regions}:
                self._strip_extra_paragraphs(cell)

        for index, header in enumerate(headers):
            write(0, index, header, header_bold)
        for r_idx, row in enumerate(rows, start=1):
            for c_idx, value in enumerate(row):
                write(r_idx, c_idx, value, False)

        # 行高最后设：合并会重建 tr，先设会被丢掉
        if row_heights:
            for row, height in zip(table.rows, row_heights):
                if height is None:
                    continue
                row.height = Cm(height)
                row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        return table

    # ================================================================
    # 5. 页眉页脚（新增能力）
    # ================================================================

    @staticmethod
    def _text_width_cm(section) -> float:
        """版心宽度（页面宽 - 左右页边距），单位 cm。"""
        return (
            section.page_width.cm - section.left_margin.cm - section.right_margin.cm
        )

    @staticmethod
    def header_footer_part(section, *, target: str = "header", variant: str = "primary"):
        """取某一节的页眉/页脚部件。

        :param section: 目标节。
        :param target: ``"header"`` 或 ``"footer"``。
        :param variant: ``"primary"``（默认/奇数页）、``"first"``（首页）、``"even"``（偶数页），
                        也接受 ``"首页"``/``"偶页"``。
        """
        if target not in ("header", "footer"):
            raise ValueError(f"target 只能是 header/footer，收到 {target!r}")
        keys = _HEADER_FOOTER_VARIANTS.get(str(variant).strip().lower())
        if keys is None:
            raise ValueError(
                f"未知的页眉页脚变体: {variant!r}。支持 primary/first/even 或 首页/偶页。"
            )
        return getattr(section, keys[0] if target == "header" else keys[1])

    @staticmethod
    def set_different_first_page(section, enabled: bool = True):
        """设置本节首页是否使用独立的页眉页脚（``w:titlePg``）。

        :param section: 目标节。
        :param enabled: True 时首页可用 ``variant="first"`` 单独设置页眉页脚。
        """
        section.different_first_page_header_footer = bool(enabled)
        return section

    def set_different_odd_even(self, enabled: bool = True):
        """设置全文奇偶页是否使用不同的页眉页脚（文档级 ``w:evenAndOddHeaders``）。

        :param enabled: True 时``variant="even"`` 的页眉页脚才会生效。
        """
        settings = self.doc.settings.element
        existing = settings.find(qn("w:evenAndOddHeaders"))
        if not enabled:
            if existing is not None:
                settings.remove(existing)
            return self.doc

        if existing is None:
            existing = OxmlElement("w:evenAndOddHeaders")
            # settings.xml 是顺序敏感的 CT_Settings，插在 compat/rsids 之前才合法
            anchor = None
            for tag in ("w:compat", "w:rsids", "w:updateFields", "w:themeFontLang"):
                found = settings.find(qn(tag))
                if found is not None:
                    anchor = found
                    break
            if anchor is not None:
                anchor.addprevious(existing)
            else:
                settings.append(existing)
        existing.set(qn("w:val"), "true")
        return self.doc

    @staticmethod
    def _add_text_with_tabs(paragraph, text: str, *, font_name: str, font_size,
                            bold: bool = False):
        """把 ``\\t`` 分隔的文本写入段落，制表符落成真正的 ``<w:tab/>``。

        python-docx 的 ``add_run("a\\tb")`` 会把制表符写进 ``<w:t>``，Word 不把它当
        制表位处理，因此必须显式插入 ``w:tab`` 元素。
        """
        runs = []
        for index, segment in enumerate(str(text).split("\t")):
            if index:
                tab_run = paragraph.add_run()
                tab_run._r.append(OxmlElement("w:tab"))
                runs.append(tab_run)
            if segment:
                run = paragraph.add_run(segment)
                WordFormatter.set_run_font(
                    run, cn_font=font_name, en_font=DEFAULT_EN_FONT,
                    size=font_size, bold=bold, color=(0, 0, 0),
                )
                runs.append(run)
        return runs

    @staticmethod
    def set_header(section, text: str, *, variant: str = "primary", alignment="居中",
                   font_name: str = DEFAULT_HEADER_FOOTER_FONT,
                   font_size=DEFAULT_HEADER_FOOTER_SIZE,
                   bottom_border: bool = True, line_length=None,
                   line_alignment="居左"):
        """设置指定节的页眉文本（默认楷体小五号）。

        :param section: 目标节。
        :param text: 页眉文本；含 ``\\t`` 时按制表位分列（配合 set_header_parts 使用）。
        :param variant: ``primary``（默认/奇数页）、``first``（首页）、``even``（偶数页）。
        :param alignment: 对齐方式，默认居中，支持中文、英文、单字母、数字和对齐枚举。
        :param font_name: 字体名称。
        :param font_size: 字号，支持中文字号字符串。
        """
        return WordFormatter._fill_header_footer_part(
            WordFormatter.header_footer_part(section, target="header", variant=variant),
            text, alignment, font_name, font_size,
            bottom_border=bottom_border, section=section,
            line_length=line_length, line_alignment=line_alignment,
        )

    @staticmethod
    def set_footer(section, text: str, *, variant: str = "primary", alignment="居中",
                   font_name: str = DEFAULT_HEADER_FOOTER_FONT,
                   font_size=DEFAULT_HEADER_FOOTER_SIZE,
                   bottom_border: bool = False, top_border: bool = True,
                   line_length=None,
                   line_alignment="居左"):
        """设置指定节的页脚文本（默认楷体小五号）。

        :param section: 目标节。
        :param text: 页脚文本。已有的页码域会被保留。
        :param variant: ``primary``（默认/奇数页）、``first``（首页）、``even``（偶数页）。
        :param alignment: 对齐方式，默认居中，支持中文、英文、单字母、数字和对齐枚举。
        :param font_name: 字体名称。
        :param font_size: 字号，支持中文字号字符串。
        :note: 页码请使用 add_footer_page_number 或 set_page_number_from_section。
        """
        return WordFormatter._fill_header_footer_part(
            WordFormatter.header_footer_part(section, target="footer", variant=variant),
            text, alignment, font_name, font_size,
            bottom_border=bottom_border, section=section,
            line_length=line_length, line_alignment=line_alignment,
            top_border=top_border, separate_page_number=True,
        )

    @staticmethod
    def set_header_parts(section, *, left: str = "", center: str = "", right: str = "",
                         variant: str = "primary",
                         font_name: str = DEFAULT_HEADER_FOOTER_FONT,
                         font_size=DEFAULT_HEADER_FOOTER_SIZE,
                         bottom_border: bool = True, line_length=None,
                         line_alignment: str = "居左"):
        """在一行内分左、中、右三段设置页眉（中文报告最常见的页眉形式）。

        用制表位实现：左段贴版心左边界、中段居中、右段贴版心右边界，
        因此「文档名 + 页码」「公司名 + 报告名」这类页眉无需手工数空格。

        :param section: 目标节。
        :param left: 左段文本。
        :param center: 中段文本。
        :param right: 右段文本。
        :param variant: ``primary``（默认/奇数页）、``first``（首页）、``even``（偶数页）。
        :param font_name: 字体名称。
        :param font_size: 字号，支持中文字号字符串。
        :param bottom_border: 是否给页眉段落加下横线。
        :param line_length: 横线长度，单位 cm；省略时使用页面可用宽度。
        :param line_alignment: 横线对齐方式。
        """
        part = WordFormatter.set_header(
            section, "\t".join([left, center, right]), variant=variant,
            alignment="居左", font_name=font_name, font_size=font_size,
            bottom_border=bottom_border, line_length=line_length,
            line_alignment=line_alignment,
        )
        WordFormatter._apply_tri_section_tab_stops(part, section)
        return part

    @staticmethod
    def set_footer_parts(section, *, left: str = "", center: str = "", right: str = "",
                         variant: str = "primary",
                         font_name: str = DEFAULT_HEADER_FOOTER_FONT,
                         font_size=DEFAULT_HEADER_FOOTER_SIZE,
                         top_border: bool = True, line_length=None,
                         line_alignment: str = "居左"):
        """在一行内分左、中、右三段设置页脚；参数含义同 :meth:`set_header_parts`。

        :param top_border: 是否给页脚段落加上横线。
        """
        part = WordFormatter.set_footer(
            section, "\t".join([left, center, right]), variant=variant,
            alignment="居左", font_name=font_name, font_size=font_size,
            top_border=top_border, line_length=line_length,
            line_alignment=line_alignment,
        )
        WordFormatter._apply_tri_section_tab_stops(part, section)
        return part

    @staticmethod
    def _apply_tri_section_tab_stops(part, section):
        """给页眉/页脚首段设置「居中 + 居右」两个制表位，实现三段分列。"""
        paragraph = part.paragraphs[0]
        stops = paragraph.paragraph_format.tab_stops
        stops.clear_all()
        width = WordFormatter._text_width_cm(section)
        stops.add_tab_stop(Cm(width / 2), WD_TAB_ALIGNMENT.CENTER)
        stops.add_tab_stop(Cm(width), WD_TAB_ALIGNMENT.RIGHT)
        return paragraph

    @staticmethod
    def set_header_image(section, image_path, *, variant: str = "primary",
                         width: float | None = None,
                         height: float | None = None, alignment="居中",
                         floating: bool = True, bottom_border: bool = True,
                         line_length=None, line_alignment="居左",
                         y_offset_pt: float = 0):
        """设置页眉图片；已有文字会保留，已有图片会被替换。

        :param section: 目标节。
        :param image_path: 图片路径，必须存在。
        :param variant: ``primary``（默认/奇数页）、``first``（首页）、``even``（偶数页）。
        :param width: 图片宽度，单位 cm，可选。
        :param height: 图片高度，单位 cm，可选；width 和 height 至少传一个。
        :param alignment: 对齐方式，默认居中，支持中文、英文、单字母、数字和对齐枚举。
        """
        return WordFormatter._set_header_footer_image(
            WordFormatter.header_footer_part(section, target="header", variant=variant),
            image_path, width=width, height=height,
            alignment=alignment, floating=floating,
            bottom_border=bottom_border, section=section,
            line_length=line_length, line_alignment=line_alignment,
            y_offset_pt=y_offset_pt,
        )

    @staticmethod
    def set_footer_image(section, image_path, *, variant: str = "primary",
                         width: float | None = None,
                         height: float | None = None, alignment="居中",
                         floating: bool = True, bottom_border: bool = True,
                         line_length=None, line_alignment="居左",
                         y_offset_pt: float = 0):
        """设置页脚图片；已有文字会保留，已有图片会被替换。

        :param section: 目标节。
        :param image_path: 图片路径，必须存在。
        :param variant: ``primary``（默认/奇数页）、``first``（首页）、``even``（偶数页）。
        :param width: 图片宽度，单位 cm，可选。
        :param height: 图片高度，单位 cm，可选；width 和 height 至少传一个。
        :param alignment: 对齐方式，默认居中，支持中文、英文、单字母、数字和对齐枚举。
        """
        return WordFormatter._set_header_footer_image(
            WordFormatter.header_footer_part(section, target="footer", variant=variant),
            image_path, width=width, height=height,
            alignment=alignment, floating=floating,
            bottom_border=bottom_border, section=section,
            line_length=line_length, line_alignment=line_alignment,
            y_offset_pt=y_offset_pt,
        )

    @staticmethod
    def _set_header_footer_image(part, image_path, *, width=None, height=None,
                                 alignment="居中", floating=False,
                                 bottom_border=False, section=None,
                                 line_length=None, line_alignment="居左",
                                 y_offset_pt=0):
        path = Path(image_path)
        if not path.is_file():
            raise FileNotFoundError(f"页眉页脚图片不存在: {path}")
        if width is None and height is None:
            raise ValueError("width 和 height 至少设置一个")

        part.is_linked_to_previous = False
        paragraph = part.paragraphs[0] if part.paragraphs else part.add_paragraph()
        for extra in part.paragraphs[1:]:
            extra._element.getparent().remove(extra._element)
        # 只替换已有图片，保留同一段落中的文字，支持图片和文字共存。
        for run in list(paragraph.runs):
            if run._r.xpath(".//w:drawing") or run._r.xpath(".//w:pict"):
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
        if floating:
            WordFormatter._make_run_floating(run, y_offset_pt=y_offset_pt)
        if bottom_border:
            WordFormatter._set_bottom_border(
                paragraph, section=section, line_length=line_length,
                line_alignment=line_alignment,
            )
        return part

    @staticmethod
    def _fill_header_footer_part(part, text, alignment, font_name, font_size,
                                 *, bottom_border=False, section=None,
                                 line_length=None, line_alignment="居左",
                                 top_border=False, separate_page_number=False):
        """页眉/页脚共用填充实现，保留已有图片和页码域。"""
        part.is_linked_to_previous = False
        paragraph = part.paragraphs[0] if part.paragraphs else part.add_paragraph()
        if separate_page_number and any(WordFormatter._is_field_run(run) for run in paragraph.runs):
            new_paragraph = part.add_paragraph("")
            part._element.remove(new_paragraph._p)
            part._element.insert(list(part._element).index(paragraph._p), new_paragraph._p)
            paragraph = new_paragraph
        # 只替换普通文字，保留图片和 PAGE 域，支持文字、图片、页码共存。
        for run in list(paragraph.runs):
            if (
                not run._r.xpath(".//w:drawing")
                and not run._r.xpath(".//w:pict")
                and not WordFormatter._is_field_run(run)
            ):
                run._r.getparent().remove(run._r)
        paragraph.alignment = WordFormatter.resolve_alignment(alignment, strict=True)
        paragraph.paragraph_format.line_spacing = 1
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        if text:
            written_runs = WordFormatter._add_text_with_tabs(
                paragraph, text, font_name=font_name, font_size=font_size,
            )
            # 页码已经存在时，把新文字放到 PAGE 域之前。
            field_run = next(
                (item for item in paragraph.runs if WordFormatter._is_field_run(item)),
                None,
            )
            if field_run is not None:
                for written in written_runs:
                    paragraph._p.remove(written._r)
                field_index = list(paragraph._p).index(field_run._r)
                for offset, written in enumerate(written_runs):
                    paragraph._p.insert(field_index + offset, written._r)
        if bottom_border:
            WordFormatter._set_bottom_border(
                paragraph, section=section, line_length=line_length,
                line_alignment=line_alignment,
            )
        if top_border:
            WordFormatter._set_top_border(
                paragraph, section=section, line_length=line_length,
                line_alignment=line_alignment,
            )
        return part

    @staticmethod
    def _set_top_border(paragraph, *, section=None, line_length=None,
                        line_alignment="居左"):
        """给页眉/页脚段落增加顶部横线。"""
        p_pr = paragraph._p.get_or_add_pPr()
        p_bdr = p_pr.find(qn("w:pBdr"))
        if p_bdr is None:
            p_bdr = OxmlElement("w:pBdr")
            p_pr.append(p_bdr)
        top = p_bdr.find(qn("w:top"))
        if top is None:
            top = OxmlElement("w:top")
            p_bdr.append(top)
        top.set(qn("w:val"), "single")
        top.set(qn("w:sz"), "6")
        top.set(qn("w:space"), "1")
        top.set(qn("w:color"), "auto")
        if section is not None:
            available = section.page_width.cm - section.left_margin.cm - section.right_margin.cm
            length = available if line_length is None else max(0, min(float(line_length), available))
            remaining = available - length
            if line_alignment in ("居右", "右对齐", "right", "R"):
                paragraph.paragraph_format.left_indent = Cm(remaining)
                paragraph.paragraph_format.right_indent = Cm(0)
            elif line_alignment in ("居中", "居中对齐", "center", "C"):
                paragraph.paragraph_format.left_indent = Cm(remaining / 2)
                paragraph.paragraph_format.right_indent = Cm(remaining / 2)
            else:
                paragraph.paragraph_format.left_indent = Cm(0)
                paragraph.paragraph_format.right_indent = Cm(remaining)

    @staticmethod
    def _is_field_run(run):
        return bool(run._r.xpath(".//w:fldChar | .//w:instrText"))

    @staticmethod
    def _set_bottom_border(paragraph, *, section=None, line_length=None,
                           line_alignment="居左", border_color="auto"):
        """给页眉/页脚段落增加底部横线。"""
        p_pr = paragraph._p.get_or_add_pPr()
        p_bdr = p_pr.find(qn("w:pBdr"))
        if p_bdr is None:
            p_bdr = OxmlElement("w:pBdr")
            p_pr.append(p_bdr)
        bottom = p_bdr.find(qn("w:bottom"))
        if bottom is None:
            bottom = OxmlElement("w:bottom")
            p_bdr.append(bottom)
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "1")
        if isinstance(border_color, tuple):
            border_color = "%02X%02X%02X" % border_color
        bottom.set(qn("w:color"), border_color)
        if section is not None:
            page_width = section.page_width.cm
            available_width = page_width - section.left_margin.cm - section.right_margin.cm
            length = available_width if line_length is None else max(0, min(float(line_length), available_width))
            remaining = available_width - length
            if line_alignment in ("居右", "右对齐", "right", "R"):
                paragraph.paragraph_format.left_indent = Cm(remaining)
                paragraph.paragraph_format.right_indent = Cm(0)
            elif line_alignment in ("居中", "居中对齐", "center", "C"):
                paragraph.paragraph_format.left_indent = Cm(remaining / 2)
                paragraph.paragraph_format.right_indent = Cm(remaining / 2)
            else:
                paragraph.paragraph_format.left_indent = Cm(0)
                paragraph.paragraph_format.right_indent = Cm(remaining)

    @staticmethod
    def _make_run_floating(run, *, y_offset_pt: float = 0):
        """将图片 run 转为 Word 兼容的 VML 浮动对象。"""
        inline = run._r.xpath(".//wp:inline")
        if not inline:
            return
        inline = inline[0]
        extent = inline.find(qn("wp:extent"))
        if extent is None:
            return
        drawing = inline.getparent()
        blip = inline.xpath(".//a:blip")
        if drawing is None or not blip:
            return
        image_rel_id = blip[0].get(qn("r:embed"))
        if not image_rel_id:
            return

        # Word for Mac 会拒绝本模块此前手工拼接的 wp:anchor。VML 同样是
        # Word 支持的绝对定位图片格式，且在页眉、页脚中兼容性更好。
        image_width = int(extent.get("cx", "0")) / 12700
        image_height = int(extent.get("cy", "0")) if extent is not None else 0
        image_height_pt = image_height / 12700
        top_offset = -(image_height_pt / 2) - float(y_offset_pt)

        # python-docx 默认没有注册 VML 的 v/o 前缀，因此此处完整声明命名空间。
        pict = parse_xml(
            '<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:v="urn:schemas-microsoft-com:vml" '
            'xmlns:o="urn:schemas-microsoft-com:office:office" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<v:shape id="FloatingPicture{run.part.next_id}" type="#_x0000_t75" '
            'style="position:absolute;'
            'margin-left:-5pt;'
            f'margin-top:{top_offset:.2f}pt;'
            f'width:{image_width:.2f}pt;'
            f'height:{image_height_pt:.2f}pt;'
            'z-index:251658240;'
            'mso-wrap-style:none;'
            'mso-position-horizontal:left;'
            'mso-position-horizontal-relative:margin;'
            'mso-position-vertical-relative:line" '
            'o:allowincell="f">'
            f'<v:imagedata r:id="{image_rel_id}" o:title=""/>'
            '</v:shape>'
            '</w:pict>'
        )
        drawing.getparent().replace(drawing, pict)

    @staticmethod
    def clear_footer(section, *, variant: str = "primary"):
        """清空指定节的页脚内容。

        :param section: 目标节。
        :param variant: ``primary``（默认/奇数页）、``first``（首页）、``even``（偶数页）。
        """
        part = WordFormatter.header_footer_part(section, target="footer", variant=variant)
        part.is_linked_to_previous = False
        for paragraph in part.paragraphs:
            for run in list(paragraph.runs):
                run._r.getparent().remove(run._r)
        return section

    @staticmethod
    def clear_header(section, *, variant: str = "primary"):
        """清空指定节的页眉内容。

        :param section: 目标节。
        :param variant: ``primary``（默认/奇数页）、``first``（首页）、``even``（偶数页）。
        """
        part = WordFormatter.header_footer_part(section, target="header", variant=variant)
        part.is_linked_to_previous = False
        for paragraph in part.paragraphs:
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
                       inherit_header: bool = True,
                       inherit_footer: bool = True,
                       show_total_pages: bool = False,
                       total_pages_separator: str = " / ",
                       prefix: str = "", suffix: str = "",
                       alignment="居中", font_name: str | None = None,
                       font_size=None, number_format=None,
                       orientation=None, page_width: float | None = None,
                       page_height: float | None = None,
                       top: float | None = None, bottom: float | None = None,
                       left: float | None = None, right: float | None = None,
                       gutter: float | None = None):
        """插入新节。

        默认不添加页码；``add_page_number=True`` 时默认跟随上一节。
        ``restart_page_number=True`` 时从 ``start_page_number`` 重新编号。

        :param start_type: 分节方式，默认 NEW_PAGE；已有分页符时自动避免重复换页。
        :param add_page_number: 是否设置页码，默认否。
        :param restart_page_number: 是否重新开始编号，默认否。
        :param start_page_number: 重新编号的起始页码，默认 1。
        :param inherit_header: 是否继承上一节页眉，默认是。
        :param inherit_footer: 是否继承上一节页脚，默认是；重新开始页码时会由本节页码覆盖。
        :param prefix: 页码前缀。
        :param suffix: 页码后缀。
        :param alignment: 页码对齐方式，默认居中，支持中文、英文、单字母、数字和对齐枚举。
        :param font_name: 页码字体名称。
        :param font_size: 页码字号，支持中文字号字符串。
        :param number_format: 页码编号格式，如 ``"upperRoman"``/``"大写罗马"``；省略时沿用原格式。
        :param orientation: 本节纸张方向，``"纵向"``/``"横向"``；与上一节不同时自动交换宽高。
        :param page_width/page_height: 本节纸张宽度/高度，单位 cm，可选。
        :param top/bottom/left/right: 本节页边距，单位 cm；只设置传入的项，
                                      未传的项沿用 python-docx 新建节的默认值。
        :param gutter: 本节装订线宽度，单位 cm，可选。
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
        if inherit_header:
            section.header.is_linked_to_previous = True
        else:
            self.clear_header(section)

        # 纸张方向/尺寸/边距必须按节设置：报告里横向插页与纵向正文混排是常态
        if any(value is not None for value in (
            orientation, page_width, page_height, top, bottom, left, right, gutter,
        )):
            self.set_section_page(
                section, orientation=orientation, page_width=page_width,
                page_height=page_height, top=top, bottom=bottom, left=left,
                right=right, gutter=gutter,
            )

        if add_page_number:
            if restart_page_number:
                self.restart_page_numbering(
                    section, start=start_page_number, add_footer_number=True,
                    prefix=prefix, suffix=suffix, alignment=alignment,
                    font_name=font_name, font_size=font_size,
                    number_format=number_format,
                    show_total_pages=show_total_pages,
                    total_pages_separator=total_pages_separator,
                )
            elif inherit_footer:
                section.footer.is_linked_to_previous = True
            else:
                # 不继承页脚时，仍在本节添加连续编号的页码。
                self.add_footer_page_number(
                    section, prefix=prefix, suffix=suffix, alignment=alignment,
                    font_name=font_name, font_size=font_size,
                    show_total_pages=show_total_pages,
                    total_pages_separator=total_pages_separator,
                )
        elif inherit_footer:
            section.footer.is_linked_to_previous = True
        else:
            self.clear_footer(section)

        if number_format is not None and not restart_page_number:
            WordFormatter.set_page_number_format(section, number_format)
        return section

    @staticmethod
    def set_page_number_start(section, start: int = 1, *, number_format=None):
        """设置指定节页码起始值与编号格式。

        :param section: 目标节。
        :param start: 起始页码。
        :param number_format: 编号格式，如 ``"upperRoman"``/``"大写罗马"``、
                              ``"lowerRoman"``/``"小写罗马"``、``"chineseCounting"``；
                              省略时保留原有格式。
        """
        fmt = WordFormatter.resolve_page_number_format(number_format)
        sect_pr = section._sectPr
        pg_num_type = sect_pr.find(qn("w:pgNumType"))
        if pg_num_type is None:
            pg_num_type = OxmlElement("w:pgNumType")
            sect_pr.append(pg_num_type)
        pg_num_type.set(qn("w:start"), str(start))
        if fmt is not None:
            pg_num_type.set(qn("w:fmt"), fmt)
        return section

    @staticmethod
    def set_page_number_format(section, number_format, *, start: int | None = None):
        """设置某节的页码编号格式（可选同时改起始页码）。

        :param section: 目标节。
        :param number_format: 编号格式，支持 ``upperRoman``/``lowerRoman``/
                              ``chineseCounting``/``decimal``/``decimalEnclosedCircle`` 等，
                              也接受 ``大写罗马``/``小写罗马``/``中文数字``/``带圈数字``。
        :param start: 同时指定起始页码，可选；省略时保留原有起始值。
        """
        fmt = WordFormatter.resolve_page_number_format(number_format)
        if fmt is None:
            raise ValueError("number_format 不能为空")
        sect_pr = section._sectPr
        pg_num_type = sect_pr.find(qn("w:pgNumType"))
        if pg_num_type is None:
            pg_num_type = OxmlElement("w:pgNumType")
            sect_pr.append(pg_num_type)
        pg_num_type.set(qn("w:fmt"), fmt)
        if start is not None:
            pg_num_type.set(qn("w:start"), str(start))
        return section

    def set_section_page(self, section, *, orientation=None,
                         page_width: float | None = None,
                         page_height: float | None = None,
                         top: float | None = None, bottom: float | None = None,
                         left: float | None = None, right: float | None = None,
                         gutter: float | None = None,
                         horizontal_alignment=None):
        """单独设置某一节的纸张方向、尺寸与页边距。

        报告里横向插页（宽表格、图纸）与纵向正文混排时，必须按节设置，
        :meth:`set_page_margins` 只适合全文统一边距的场景。

        :param section: 目标节。
        :param orientation: ``"纵向"``/``"横向"`` 或 ``portrait``/``landscape``；
                            与当前方向不同时会自动交换宽高。
        :param page_width: 页面宽度，单位 cm，可选。
        :param page_height: 页面高度，单位 cm，可选。
        :param top/bottom/left/right: 页边距，单位 cm，只设置传入的项。
        :param gutter: 装订线宽度，单位 cm，可选。
        :param horizontal_alignment: 页面水平对齐，支持 L/C/R 或 左对齐/居中/右对齐。
        """
        resolved = WordFormatter.resolve_orientation(orientation)
        width_cm = page_width
        height_cm = page_height

        if resolved is not None:
            target_orient, ooxml_orient = resolved
            current_is_landscape = section.page_width > section.page_height
            target_is_landscape = target_orient == WD_ORIENT.LANDSCAPE
            available = section.page_width.cm, section.page_height.cm
            if width_cm is None and height_cm is None and \
                    current_is_landscape != target_is_landscape:
                width_cm, height_cm = available[1], available[0]
            section.orientation = target_orient
            sect_pr = section._sectPr
            pg_sz = sect_pr.find(qn("w:pgSz"))
            if pg_sz is not None:
                pg_sz.set(qn("w:orient"), ooxml_orient)

        if width_cm is not None:
            section.page_width = Cm(width_cm)
        if height_cm is not None:
            section.page_height = Cm(height_cm)

        if top is not None:
            section.top_margin = Cm(top)
        if bottom is not None:
            section.bottom_margin = Cm(bottom)
        if left is not None:
            section.left_margin = Cm(left)
        if right is not None:
            section.right_margin = Cm(right)
        if gutter is not None:
            section.gutter = Cm(gutter)

        if horizontal_alignment is not None:
            jc_value = self.resolve_page_justification(horizontal_alignment)
            sect_pr = section._sectPr
            jc = sect_pr.find(qn("w:jc"))
            if jc is None:
                jc = OxmlElement("w:jc")
                sect_pr.append(jc)
            jc.set(qn("w:val"), jc_value)
        return section

    @staticmethod
    def add_page_number(paragraph, *, prefix: str = "", suffix: str = "",
                       alignment="居中", font_name: str | None = None,
                       font_size=None, show_total_pages: bool = False,
                       total_pages_separator: str = " / "):
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
        if show_total_pages:
            separator_run = paragraph.add_run(total_pages_separator)
            WordFormatter.set_run_font(separator_run, cn_font=cn_font, size=size)
            total_runs = WordFormatter._append_field_run(paragraph, "NUMPAGES")
            for run in total_runs:
                WordFormatter.set_run_font(run, cn_font=cn_font, size=size)
        return paragraph

    @staticmethod
    def add_footer_page_number(section, *, prefix: str = "", suffix: str = "",
                                alignment="居中", font_name: str | None = None,
                                font_size=None, show_total_pages: bool = False,
                                total_pages_separator: str = " / "):
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
        # 保留已有 PAGE 域，避免重新设置页脚文字时页码消失。
        for run in list(paragraph.runs):
            if not WordFormatter._is_field_run(run):
                run._r.getparent().remove(run._r)
        paragraph.paragraph_format.line_spacing = 1
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        return WordFormatter.add_page_number(
            paragraph, prefix=prefix, suffix=suffix, alignment=alignment,
            font_name=font_name, font_size=font_size,
            show_total_pages=show_total_pages,
            total_pages_separator=total_pages_separator,
        )

    @staticmethod
    def restart_page_numbering(section, *, start: int = 1, add_footer_number: bool = True,
                               prefix: str = "", suffix: str = "", alignment="居中",
                               font_name: str | None = None, font_size=None,
                               number_format=None,
                               show_total_pages: bool = False,
                               total_pages_separator: str = " / "):
        """重启指定节的页码编号。

        :param section: 目标节。
        :param start: 起始页码，默认 1。
        :param add_footer_number: 是否同时添加页脚页码，默认是。
        :param prefix: 页码前文本。
        :param suffix: 页码后文本。
        :param alignment: 对齐方式，默认居中，支持中文、英文、单字母、数字和对齐枚举。
        :param font_name: 页码字体名称。
        :param font_size: 页码字号，支持中文字号字符串。
        :param number_format: 页码编号格式，如 ``"upperRoman"``/``"大写罗马"``、
                              ``"chineseCounting"``；省略时沿用原有格式。
        """
        section.header.is_linked_to_previous = False
        section.footer.is_linked_to_previous = False
        WordFormatter.set_page_number_start(
            section, start=start, number_format=number_format,
        )
        if add_footer_number:
            WordFormatter.add_footer_page_number(
                section, prefix=prefix, suffix=suffix, alignment=alignment,
                font_name=font_name, font_size=font_size,
                show_total_pages=show_total_pages,
                total_pages_separator=total_pages_separator,
            )
        return section

    def insert_section_with_page_numbering(self, *, start_type=WD_SECTION_START.NEW_PAGE,
                                           start_page_number: int = 1,
                                           add_footer_number: bool = True,
                                           prefix: str = "", suffix: str = "",
                                           alignment="居中", font_name: str | None = None,
                                           font_size=None, number_format=None):
        """插入新节并设置页码；默认从指定数字重新开始。

        :param start_type: 分节方式，默认 NEW_PAGE。
        :param start_page_number: 新节起始页码，默认 1。
        :param add_footer_number: 是否添加页脚页码，默认是。
        :param prefix/suffix: 页码前后文本。
        :param alignment: 页码对齐方式，支持中文、英文、单字母、数字和对齐枚举。
        :param font_name: 页码字体名称。
        :param font_size: 页码字号，支持中文字号字符串。
        :param number_format: 页码编号格式，如 ``"upperRoman"``/``"大写罗马"``。
        """
        return self.insert_section(
            start_type=start_type, add_page_number=add_footer_number,
            restart_page_number=True, start_page_number=start_page_number,
            prefix=prefix, suffix=suffix, alignment=alignment,
            font_name=font_name, font_size=font_size, number_format=number_format,
        )

    def set_page_number_from_section(self, start_section_idx: int, *, start: int = 1,
                                    prefix: str = "", suffix: str = "",
                                    alignment="居中", font_name: str | None = None,
                                    font_size=None, number_format=None):
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
        :param number_format: 页码编号格式，如 ``"upperRoman"``/``"大写罗马"``；
                              封面/目录与正文用不同编号体系时用它区分。
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
            font_name=font_name, font_size=font_size, number_format=number_format,
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
                            first_line_indent=None, first_line_indent_chars=None,
                            left_indent=None, alignment=None):
        """设置某一级 TOC 样式。

        :param level: 目录层级。
        :param font_name: 字体名称，可选。
        :param font_size: 字号，支持中文字号字符串，可选。
        :param bold: 是否加粗，可选。
        :param color: RGB 颜色元组，可选。
        :param line_spacing: 行距倍数，可选。
        :param space_before/space_after: 段前/段后空白，单位 pt，可选。
        :param first_line_indent: 首行缩进，单位 pt，可选。
        :param first_line_indent_chars: 首行缩进字符数，可选；设置后优先于 pt 值。
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
                "first_line_indent": first_line_indent,
                "first_line_indent_chars": first_line_indent_chars,
                "left_indent": left_indent,
                "alignment": alignment,
            },
        )

    def set_paragraph_style(self, style_name: str, *, base_style_name: str = "Normal",
                            font_name: str | None = None, font_size=None,
                            bold: bool | None = None,
                            color: tuple[int, int, int] | None = None,
                            line_spacing=None, space_before=None, space_after=None,
                            first_line_indent=None, first_line_indent_chars=None,
                            left_indent=None, alignment=None):
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
        :param first_line_indent_chars: 首行缩进字符数，可选；设置后优先于 pt 值。
        :param alignment: 对齐方式，可选。
        """
        style = self._get_or_create_style(style_name, base_style_name)
        return self._apply_style_options(
            style,
            {
                "font_name": font_name, "font_size": font_size, "bold": bold,
                "color": color, "line_spacing": line_spacing,
                "space_before": space_before, "space_after": space_after,
                "first_line_indent": first_line_indent,
                "first_line_indent_chars": first_line_indent_chars,
                "left_indent": left_indent,
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
                title_style: str = "Normal",
                title_font_name: str = "黑体", title_font_size=16,
                title_bold: bool = True, title_alignment="居中",
                title_color: tuple[int, int, int] | None = None,
                title_bottom_border: bool = False,
                title_border_color: tuple[int, int, int] | None = None,
                title_border_length: float | None = None,
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
        :param title_style: 目录标题使用的段落样式，默认 ``"Normal"``，避免模板 ``Title`` 样式自带下划线。
        :param title_font_name: 标题字体名称。
        :param title_font_size: 标题字号，支持中文字号字符串。
        :param title_bold: 目录标题是否加粗。
        :param title_alignment: 目录标题对齐方式。
        :param title_color: 目录标题颜色，RGB 元组，如 ``(31, 78, 121)``。
        :param title_bottom_border: 是否为目录标题添加底部横线。
        :param title_border_color: 横线颜色，RGB 元组。
        :param title_border_length: 横线长度，单位 cm；默认使用页面可用宽度。
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
            title_par = self.doc.add_paragraph(title, style=title_style)
            self.set_paragraph_format(
                title_par,
                alignment=self.resolve_alignment(title_alignment),
                space_before=12, space_after=12,
            )
            if title_par.runs:
                self.set_run_font(
                    title_par.runs[0], cn_font=title_font_name,
                    size=title_font_size, bold=title_bold, color=title_color,
                )
            if title_bottom_border:
                self._set_bottom_border(
                    title_par, section=self.doc.sections[0],
                    line_length=title_border_length, line_alignment="居中",
                    border_color=title_border_color or "auto",
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

    # ================================================================
    # 8. 模板填充（在已有内容上原位改写）
    # ================================================================
    #
    # 前面各节都是「从空白文档按顺序追加」，适合生成报告；但招投标这类模板是
    # 「拿一份现成的 .docx，把里面的空填上」，追加型接口会把内容写到文档末尾。
    # 本节补的就是这条缺失的路：先看清已有内容，再就地改写，且格式沿用原处。

    @staticmethod
    def _style_chain(style) -> list:
        """段落样式的继承链：自身 -> 基样式 -> … -> Normal。"""
        chain = []
        seen = set()
        current = style
        while current is not None and id(current) not in seen:
            seen.add(id(current))
            chain.append(current)
            current = current.base_style
        return chain

    @staticmethod
    def _document_default_font(paragraph):
        """从 styles.xml 的 ``docDefaults`` 取默认中文字体与字号。"""
        part = getattr(paragraph, "part", None)
        styles = getattr(part, "styles", None)
        if styles is None:
            return None, None
        for r_pr in styles.element.xpath("./w:docDefaults/w:rPrDefault/w:rPr"):
            fonts = r_pr.find(qn("w:rFonts"))
            name = None
            if fonts is not None:
                name = fonts.get(qn("w:eastAsia")) or fonts.get(qn("w:ascii"))
            size = None
            sz = r_pr.find(qn("w:sz"))
            if sz is not None:
                try:
                    size = Pt(float(sz.get(qn("w:val"))) / 2)
                except (TypeError, ValueError):
                    size = None
            return name, size
        return None, None

    @staticmethod
    def _effective_font(paragraph, run=None):
        """解析**实际生效**的 (中文字体名, 字号)，run 可为 None。

        模板里的 run 常常一个字体属性都不写，全靠样式继承；只读 run 会得到一片 None，
        AI 就会以为「原文没有字体」而自己指定一个，反而破坏一致性。这里按
        run -> 段落样式链 -> docDefaults 的顺序逐层回落。

        ``run`` 传 None 表示段落本身还没有 run（模板里的空格子），
        此时全部从样式链取，用来给新写入的内容一个和邻居一致的字体。
        """
        name = None
        size = None
        if run is not None:
            r_pr = run._r.rPr
            fonts = r_pr.rFonts if r_pr is not None else None
            if fonts is not None:
                name = fonts.get(qn("w:eastAsia"))
            size = run.font.size

        style = getattr(paragraph, "style", None)
        for level in WordFormatter._style_chain(style):
            level_r_pr = level.element.find(qn("w:rPr"))
            if name is None and level_r_pr is not None:
                level_fonts = level_r_pr.find(qn("w:rFonts"))
                if level_fonts is not None:
                    name = level_fonts.get(qn("w:eastAsia"))
            if size is None:
                size = level.font.size
            if name is not None and size is not None:
                break

        if name is None or size is None:
            default_name, default_size = WordFormatter._document_default_font(paragraph)
            if name is None:
                name = default_name
            if size is None:
                size = default_size
        return name, size

    @staticmethod
    def _run_font_info(paragraph, run) -> dict:
        """读取 run 实际生效的字体（中文字体名、字号 pt、是否加粗）。"""
        name, size = WordFormatter._effective_font(paragraph, run)
        return {
            "text": run.text,
            "font_name": name,
            "font_size_pt": round(size.pt, 2) if size is not None else None,
            "bold": run.font.bold is True,
        }

    @staticmethod
    def _truncate(text: str, limit: int) -> str:
        """按 limit 截断文本，超长时补省略号；limit 为 0 表示不截断。"""
        if limit and len(text) > limit:
            return text[:limit] + "…"
        return text

    @classmethod
    def _table_block(cls, table, index: int, max_text_chars: int) -> dict:
        """表格块摘要：单元格文字网格 + 每格首个 run 的实际字体。"""
        grid = []
        fonts = []
        for row in table.rows:
            grid.append([
                cls._truncate(cell.text.strip(), max_text_chars) for cell in row.cells
            ])
            row_fonts = []
            for cell in row.cells:
                paragraphs = cell.paragraphs
                if not paragraphs:
                    row_fonts.append(None)
                    continue
                # 空格子没有 run，但仍能按样式解析出「填进去会是什么字体」
                name, size = cls._effective_font(paragraphs[0])
                row_fonts.append({
                    "font_name": name,
                    "font_size_pt": round(size.pt, 2) if size is not None else None,
                })
            fonts.append(row_fonts)
        return {
            "type": "table",
            "index": index,
            "rows": len(table.rows),
            "columns": len(table.columns),
            "grid": grid,
            "cell_fonts": fonts,
        }

    def describe_blocks(self, *, include_empty: bool = False,
                        max_text_chars: int = 200) -> list:
        """按文档顺序列出正文块（段落与表格），用来定位要改写的位置。

        生成报告时不需要它——按顺序 ``add`` 就行；但填模板必须先把已有文档看清楚。

        段落块给出它在 ``doc.paragraphs`` 里的下标，可直接用于 :meth:`replace_text`
        对应的工具；表格块给出 ``doc.tables`` 的下标与单元格网格，
        用来决定往哪个格子填值或贴图。

        :param include_empty: 是否返回空段落，默认否（模板里的留白通常不用处理）。
        :param max_text_chars: 单段/单格文本的截断长度，默认 200 字符，
                               避免长文档一次性回传过大。
        :return: 块列表，段落块形如
                 ``{"type": "paragraph", "index": 0, "text": "...", "runs": [...]}``，
                 表格块形如 ``{"type": "table", "index": 0, "grid": [[...]]}``。
        """
        blocks = []
        paragraph_index = 0
        table_index = 0
        for item in self.doc.iter_inner_content():
            if isinstance(item, Table):
                blocks.append(self._table_block(item, table_index, max_text_chars))
                table_index += 1
                continue
            index = paragraph_index
            paragraph_index += 1
            if not include_empty and not item.text.strip():
                continue
            blocks.append({
                "type": "paragraph",
                "index": index,
                "text": self._truncate(item.text, max_text_chars),
                "style": item.style.name if item.style is not None else None,
                "runs": [self._run_font_info(item, run) for run in item.runs],
            })
        return blocks

    @staticmethod
    def _run_text_spans(paragraph):
        """把段落的 run 拼成「字符区间 -> run」的映射，用于按字符定位替换范围。"""
        spans = []
        offset = 0
        for run in paragraph.runs:
            text = run.text
            spans.append((offset, offset + len(text), run, text))
            offset += len(text)
        return spans

    @staticmethod
    def _locate_offset(spans, position: int):
        """找出字符位置落在哪个 run，以及它在 run 内的偏移。"""
        for index, (start, end, _run, _text) in enumerate(spans):
            if start <= position <= end:
                return index, position - start
        last = len(spans) - 1
        return last, len(spans[last][3])

    @staticmethod
    def _set_run_text(run, text: str):
        """只改写 run 里的 ``w:t``，保留字体等其余属性。

        直接给 ``run.text`` 赋值会重建整个 ``w:r``，把制表符、换行等子元素一并丢掉；
        这里只替换文本节点，``w:tab`` / ``w:br`` 之类的结构得以保留。
        """
        element = run._r
        for existing in element.findall(qn("w:t")):
            element.remove(existing)
        node = OxmlElement("w:t")
        node.set(qn("xml:space"), "preserve")
        node.text = text
        element.append(node)
        return run

    @staticmethod
    def _highlight_run(run, enabled: bool):
        """按需把 run 标黄；``enabled`` 为假时不动原有高亮设置。"""
        if enabled:
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        return run

    @staticmethod
    def replace_text(paragraph, old: str, new: str, *, count: int = 1,
                     highlight: bool = False) -> int:
        """在已有段落里把 ``old`` 原地替换为 ``new``，**字体字号沿用原处**。

        替换过程不新建 run，只改写原 run 的文本节点，因此填进去的内容自带原位置的
        字体、字号、加粗等格式——这正是「填完和原文一致」所要求的。Word 常把一句话
        拆成好几个 run，匹配跨 run 时新文本落在**第一个** run 上，其余 run 里被匹配到
        的部分清空。

        ``old`` 传空串表示在段尾追加 ``new``，用于「委托代理人身份证号码：____」
        这类只有标签、值要补在标签后面的填空；模板里的空格子（没有任何 run）
        也走这条路，此时字体按段落样式解析出来，和邻居保持一致。

        :param paragraph: 目标段落（``docx.text.paragraph.Paragraph``）。
        :param old: 要被替换掉的原文；空串表示追加到段尾。
        :param new: 新写入的文本。
        :param count: 最多替换几处，默认 1；传 0 表示全部替换。
        :param highlight: 是否把写入 ``new`` 的那个 run 标黄，便于人工复核，默认否。
        :return: 实际替换的处数。
        :raises ValueError: 一处都没匹配上（含对空段落替换非空原文）。
        """
        spans = WordFormatter._run_text_spans(paragraph)
        if not spans:
            if old:
                raise ValueError(
                    f"该段落是空的，没有可匹配的文本 {old!r}。"
                    "往空段落/空格子里填内容请把 old 传空串。"
                )
            # 空格子没有 run 可继承格式，退回按段落样式解析出的字体写一个
            name, size = WordFormatter._effective_font(paragraph)
            run = paragraph.add_run(new)
            WordFormatter.set_run_font(
                run, cn_font=name or DEFAULT_CN_FONT,
                size=size.pt if size is not None else DEFAULT_BODY_SIZE,
            )
            WordFormatter._highlight_run(run, highlight)
            return 1

        if not old:
            index = len(spans) - 1
            run, text = spans[index][2], spans[index][3]
            WordFormatter._set_run_text(run, text + new)
            WordFormatter._highlight_run(run, highlight)
            return 1

        limit = count if count and count > 0 else None
        replaced = 0
        while limit is None or replaced < limit:
            spans = WordFormatter._run_text_spans(paragraph)
            full_text = "".join(span[3] for span in spans)
            start = full_text.find(old)
            if start < 0:
                break
            end = start + len(old)
            first_index, first_offset = WordFormatter._locate_offset(spans, start)
            last_index, last_offset = WordFormatter._locate_offset(spans, end)
            if first_index == last_index:
                text = spans[first_index][3]
                WordFormatter._set_run_text(
                    spans[first_index][2],
                    text[:first_offset] + new + text[last_offset:],
                )
            else:
                WordFormatter._set_run_text(
                    spans[first_index][2],
                    spans[first_index][3][:first_offset] + new,
                )
                for middle in range(first_index + 1, last_index):
                    WordFormatter._set_run_text(spans[middle][2], "")
                WordFormatter._set_run_text(
                    spans[last_index][2], spans[last_index][3][last_offset:],
                )
            WordFormatter._highlight_run(spans[first_index][2], highlight)
            replaced += 1

        if not replaced:
            raise ValueError(f"段落里没有找到要替换的文本 {old!r}。")
        return replaced

    @staticmethod
    def insert_cell_image(cell, image_path, *, width=None, height=None,
                          alignment="居中", keep_text: bool = False):
        """把图片放进表格单元格——「此格附身份证正面」这类版式的正确工具。

        :meth:`insert_img` 只能往正文追加图片，落到不了一个已有的格子里；而身份证、
        营业执照、印章在表单里都是要「贴进格子」的。

        ``width`` 与 ``height`` 都省略时按图片自身 DPI 使用原始尺寸。

        :param cell: ``docx.table._Cell`` 对象。
        :param image_path: 图片路径（str/Path），或图片二进制内容（bytes/bytearray）。
        :param width: 图片宽度，单位 cm，可选。
        :param height: 图片高度，单位 cm，可选。
        :param alignment: 单元格内段落对齐方式，默认居中。
        :param keep_text: 是否保留单元格已有文字，默认否（清空后再贴图，避免挤成一行）。
        :return: 该单元格。
        """
        if width is not None and width <= 0:
            raise ValueError(f"图片宽度必须大于 0，收到 {width!r}")
        if height is not None and height <= 0:
            raise ValueError(f"图片高度必须大于 0，收到 {height!r}")

        if not keep_text:
            for paragraph in cell.paragraphs:
                for run in list(paragraph.runs):
                    run._r.getparent().remove(run._r)

        paragraph = cell.paragraphs[0]
        paragraph.alignment = WordFormatter.resolve_alignment(alignment, strict=True)
        paragraph.paragraph_format.line_spacing = 1
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)

        run = paragraph.add_run()
        source, owned = WordFormatter._image_stream(image_path)
        try:
            kwargs = {}
            if width is not None:
                kwargs["width"] = Cm(width)
            if height is not None:
                kwargs["height"] = Cm(height)
            if not kwargs:
                kwargs["width"] = Cm(WordFormatter._native_width_cm(source))
            run.add_picture(source, **kwargs)
        finally:
            if owned:
                source.close()
        # 被并进来的空段落会白高一截，贴完图整理掉
        WordFormatter._strip_extra_paragraphs(cell, keep=1)
        return cell

    '''REMOVED_COMPATIBILITY_BLOCK'''


about_word = WordFormatter
