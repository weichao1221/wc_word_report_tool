from __future__ import annotations

import datetime as _datetime

from docx.enum.style import WD_STYLE_TYPE
from docx.enum.section import WD_SECTION_START
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX, WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


class WordFormatter:
    _ALIGNMENT_MAP = {
        "左对齐": WD_PARAGRAPH_ALIGNMENT.LEFT,
        "居中": WD_PARAGRAPH_ALIGNMENT.CENTER,
        "右对齐": WD_PARAGRAPH_ALIGNMENT.RIGHT,
    }

    @staticmethod
    def check_item(item) -> bool:
        try:
            return float(item) != 0
        except (TypeError, ValueError):
            return False

    @staticmethod
    def _set_run_font(run, *, cn_font: str, en_font: str = "Times New Roman", size: int = 14,
                      bold: bool = False, color: tuple[int, int, int] | None = None,
                      highlight=None) -> None:
        run.font.bold = bold
        run.font.size = Pt(size)
        run.font.name = en_font
        run.element.rPr.rFonts.set(qn("w:eastAsia"), cn_font)
        if color is not None:
            run.font.color.rgb = RGBColor(*color)
        if highlight is not None:
            run.font.highlight_color = highlight

    @staticmethod
    def _set_paragraph_basic_format(paragraph, *, alignment=None, line_spacing=1.5, space_before=0, space_after=0,
                                    first_line_indent_pt=None):
        if alignment is not None:
            paragraph.alignment = alignment
        paragraph.paragraph_format.line_spacing = line_spacing
        paragraph.paragraph_format.space_before = Pt(space_before)
        paragraph.paragraph_format.space_after = Pt(space_after)
        if first_line_indent_pt is not None:
            paragraph.paragraph_format.first_line_indent = Pt(first_line_indent_pt)
        return paragraph

    @staticmethod
    def _resolve_alignment(alignment):
        if isinstance(alignment, str):
            return WordFormatter._ALIGNMENT_MAP.get(alignment, WD_PARAGRAPH_ALIGNMENT.LEFT)
        return alignment

    @staticmethod
    def _append_field_run(paragraph, instruction: str):
        begin = OxmlElement("w:fldChar")
        begin.set(qn("w:fldCharType"), "begin")

        instr = OxmlElement("w:instrText")
        instr.set(qn("xml:space"), "preserve")
        instr.text = instruction

        separate = OxmlElement("w:fldChar")
        separate.set(qn("w:fldCharType"), "separate")

        end = OxmlElement("w:fldChar")
        end.set(qn("w:fldCharType"), "end")

        paragraph.add_run()._r.append(begin)
        paragraph.add_run()._r.append(instr)
        paragraph.add_run()._r.append(separate)
        paragraph.add_run()._r.append(end)
        return paragraph

    @staticmethod
    def _ensure_update_fields_on_open(doc):
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
            style.font.name = "Times New Roman"
            style.element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
        if size is not None:
            style.font.size = Pt(size)
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
            style.paragraph_format.alignment = WordFormatter._resolve_alignment(alignment)
        return style

    @staticmethod
    def _configure_heading_style(doc, level: int, *, cn_font: str, size: int, bold: bool,
                                 color: tuple[int, int, int] | None = None, line_spacing=1,
                                 space_before=12, space_after=12, first_line_indent=28):
        style = doc.styles[f"Heading {level}"]
        return WordFormatter._apply_style_options(
            style,
            {
                "font_name": cn_font,
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
    def _add_heading_with_style(doc, heading_text: str, *, level: int, cn_font: str, size: int, bold: bool,
                                color: tuple[int, int, int] | None = None, line_spacing=1,
                                space_before=12, space_after=12, first_line_indent=28):
        WordFormatter._configure_heading_style(
            doc,
            level,
            cn_font=cn_font,
            size=size,
            bold=bold,
            color=color,
            line_spacing=line_spacing,
            space_before=space_before,
            space_after=space_after,
            first_line_indent=first_line_indent,
        )
        heading = doc.add_heading("", level=level)
        run = heading.add_run(heading_text)
        WordFormatter._set_run_font(run, cn_font=cn_font, size=size, bold=bold, color=color)
        return heading

    @staticmethod
    def fengmian_doc1(doc, text_content: str, font_size: int = 22):
        par = doc.add_paragraph("")
        WordFormatter._set_paragraph_basic_format(par, alignment=WD_PARAGRAPH_ALIGNMENT.CENTER)
        run = par.add_run(text_content)
        WordFormatter._set_run_font(run, cn_font="宋体", size=font_size, bold=True)
        return par

    @staticmethod
    def fengmian_doc2(doc, text_content: str):
        par = doc.add_paragraph("")
        WordFormatter._set_paragraph_basic_format(par, alignment=WD_PARAGRAPH_ALIGNMENT.CENTER)
        run = par.add_run(text_content)
        WordFormatter._set_run_font(run, cn_font="方正小标宋简体", size=16)
        return par

    @staticmethod
    def fengmian_doc3(doc, text_content: str):
        par = doc.add_paragraph("")
        WordFormatter._set_paragraph_basic_format(par, alignment=WD_PARAGRAPH_ALIGNMENT.CENTER)
        run = par.add_run(text_content)
        WordFormatter._set_run_font(run, cn_font="黑体", size=22, bold=True)
        return par

    @staticmethod
    def Heading_1(doc, heading_text: str):
        return WordFormatter._add_heading_with_style(
            doc,
            heading_text,
            level=1,
            cn_font="黑体",
            size=16,
            bold=False,
            line_spacing=1.5,
            space_before=12,
            space_after=12,
            first_line_indent=28,
        )

    @staticmethod
    def Heading_2(doc, heading_text: str):
        return WordFormatter._add_heading_with_style(
            doc, heading_text, level=2, cn_font="仿宋_GB2312", size=14, bold=True, color=(0, 0, 0)
        )

    @staticmethod
    def Heading_3(doc, heading_text: str):
        return WordFormatter._add_heading_with_style(
            doc, heading_text, level=3, cn_font="仿宋_GB2312", size=14, bold=False, color=(0, 0, 0)
        )

    @staticmethod
    def Heading_union(doc, text_content: str, layout: str, font_name: str, font_size: int):
        heading = doc.add_heading("", level=2)
        heading.alignment = WordFormatter._resolve_alignment(layout) or WD_PARAGRAPH_ALIGNMENT.RIGHT
        heading.paragraph_format.line_spacing = 1
        heading.paragraph_format.space_before = Pt(12)
        heading.paragraph_format.space_after = Pt(12)
        run = heading.add_run(text_content)
        WordFormatter._set_run_font(run, cn_font=font_name, size=font_size, bold=True, color=(0, 0, 0))
        return heading

    @staticmethod
    def Normal_doc(doc, text_content: str):
        par = doc.add_paragraph("")
        WordFormatter._set_paragraph_basic_format(
            par, alignment=WD_PARAGRAPH_ALIGNMENT.JUSTIFY, first_line_indent_pt=28
        )
        run = par.add_run(text_content)
        WordFormatter._set_run_font(run, cn_font="仿宋_GB2312", size=14)
        return par

    @staticmethod
    def Normal_doc_Highlight(doc, text_content: str):
        par = WordFormatter.Normal_doc(doc, text_content)
        par.runs[-1].font.highlight_color = WD_COLOR_INDEX.YELLOW
        return par

    @staticmethod
    def Normal_doc_仿宋三号加粗(doc, text_content: str):
        par = doc.add_paragraph("")
        WordFormatter._set_paragraph_basic_format(par, alignment=WD_PARAGRAPH_ALIGNMENT.CENTER)
        run = par.add_run(text_content)
        WordFormatter._set_run_font(run, cn_font="仿宋_GB2312", size=16, bold=True)
        return par

    @staticmethod
    def set_cell_format(cell, text_content: str):
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        paragraph.paragraph_format.line_spacing = 1.5
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        run = paragraph.add_run(text_content)
        WordFormatter._set_run_font(run, cn_font="仿宋_GB2312", size=12)
        return cell

    @staticmethod
    def company_name(doc, company_name: str):
        par = doc.add_paragraph("")
        WordFormatter._set_paragraph_basic_format(
            par, alignment=WD_PARAGRAPH_ALIGNMENT.RIGHT, first_line_indent_pt=28
        )
        run = par.add_run(company_name)
        WordFormatter._set_run_font(run, cn_font="仿宋_GB2312", size=16)
        return par

    @staticmethod
    def created_time(doc, value: _datetime.date | _datetime.datetime | str | None = None):
        if value is None:
            text_content = _datetime.datetime.now().strftime("%Y年%m月%d日")
        elif isinstance(value, (_datetime.date, _datetime.datetime)):
            text_content = value.strftime("%Y年%m月%d日")
        else:
            text_content = str(value)

        par = doc.add_paragraph("")
        WordFormatter._set_paragraph_basic_format(
            par, alignment=WD_PARAGRAPH_ALIGNMENT.RIGHT, first_line_indent_pt=28
        )
        run = par.add_run(text_content)
        WordFormatter._set_run_font(run, cn_font="仿宋_GB2312", size=14)
        return par

    @staticmethod
    def insert_new_section(doc, start_type=WD_SECTION_START.NEW_PAGE):
        return doc.add_section(start_type=start_type)

    @staticmethod
    def set_page_number_start(section, start: int = 1):
        sect_pr = section._sectPr
        pg_num_type = sect_pr.find(qn("w:pgNumType"))
        if pg_num_type is None:
            pg_num_type = OxmlElement("w:pgNumType")
            sect_pr.append(pg_num_type)
        pg_num_type.set(qn("w:start"), str(start))
        return section

    @staticmethod
    def add_page_number(paragraph, *, prefix: str = "", suffix: str = "", alignment="居中"):
        paragraph.alignment = WordFormatter._resolve_alignment(alignment)
        if prefix:
            paragraph.add_run(prefix)
        WordFormatter._append_field_run(paragraph, "PAGE")
        if suffix:
            paragraph.add_run(suffix)
        return paragraph

    @staticmethod
    def add_footer_page_number(section, *, prefix: str = "", suffix: str = "", alignment="居中"):
        section.footer.is_linked_to_previous = False
        footer = section.footer
        paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        if paragraph.text:
            paragraph = footer.add_paragraph()
        return WordFormatter.add_page_number(paragraph, prefix=prefix, suffix=suffix, alignment=alignment)

    @staticmethod
    def restart_page_numbering(section, *, start: int = 1, add_footer_number: bool = True,
                               prefix: str = "", suffix: str = "", alignment="居中"):
        section.header.is_linked_to_previous = False
        section.footer.is_linked_to_previous = False
        WordFormatter.set_page_number_start(section, start=start)
        if add_footer_number:
            WordFormatter.add_footer_page_number(
                section, prefix=prefix, suffix=suffix, alignment=alignment
            )
        return section

    @staticmethod
    def insert_section_with_page_numbering(doc, *, start_type=WD_SECTION_START.NEW_PAGE, start_page_number: int = 1,
                                           add_footer_number: bool = True, prefix: str = "", suffix: str = "",
                                           alignment="居中"):
        section = WordFormatter.insert_new_section(doc, start_type=start_type)
        return WordFormatter.restart_page_numbering(
            section,
            start=start_page_number,
            add_footer_number=add_footer_number,
            prefix=prefix,
            suffix=suffix,
            alignment=alignment,
        )

    @staticmethod
    def insert_img(doc, img_path, width: float):
        par = doc.add_paragraph("")
        par.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        run = par.add_run()
        run.add_picture(img_path, width=Cm(width))
        return par

    @staticmethod
    def set_all_layout(doc, *, top: float = 2.54, bottom: float = 2.54, left: float = 3.17,
                       right: float = 3.17, gutter: float = 0):
        for section in doc.sections:
            section.top_margin = Cm(top)
            section.bottom_margin = Cm(bottom)
            section.left_margin = Cm(left)
            section.right_margin = Cm(right)
            section.gutter = Cm(gutter)
        return doc

    @staticmethod
    def set_toc_level_style(doc, level: int, *, font_name: str | None = None, font_size: int | None = None,
                            bold: bool | None = None, color: tuple[int, int, int] | None = None,
                            line_spacing=None, space_before: int | None = None, space_after: int | None = None,
                            first_line_indent: int | None = None, left_indent: int | None = None,
                            alignment=None):
        style_name = f"TOC {level}"
        style = WordFormatter._get_or_create_style(doc, style_name, "Normal")
        return WordFormatter._apply_style_options(
            style,
            {
                "font_name": font_name,
                "font_size": font_size,
                "bold": bold,
                "color": color,
                "line_spacing": line_spacing,
                "space_before": space_before,
                "space_after": space_after,
                "first_line_indent": first_line_indent,
                "left_indent": left_indent,
                "alignment": alignment,
            },
        )

    @staticmethod
    def set_paragraph_style(doc, style_name: str, *, base_style_name: str = "Normal", font_name: str | None = None,
                            font_size: int | None = None, bold: bool | None = None,
                            color: tuple[int, int, int] | None = None, line_spacing=None,
                            space_before: int | None = None, space_after: int | None = None,
                            first_line_indent: int | None = None, left_indent: int | None = None,
                            alignment=None):
        style = WordFormatter._get_or_create_style(doc, style_name, base_style_name)
        return WordFormatter._apply_style_options(
            style,
            {
                "font_name": font_name,
                "font_size": font_size,
                "bold": bold,
                "color": color,
                "line_spacing": line_spacing,
                "space_before": space_before,
                "space_after": space_after,
                "first_line_indent": first_line_indent,
                "left_indent": left_indent,
                "alignment": alignment,
            },
        )

    @staticmethod
    def add_custom_heading(doc, text_content: str, *, style_name: str, level: int | None = None):
        par = doc.add_paragraph(text_content, style=style_name)
        if level is not None:
            p_pr = par._p.get_or_add_pPr()
            outline_lvl = p_pr.find(qn("w:outlineLvl"))
            if outline_lvl is None:
                outline_lvl = OxmlElement("w:outlineLvl")
                p_pr.append(outline_lvl)
            outline_lvl.set(qn("w:val"), str(max(level - 1, 0)))
        return par

    @staticmethod
    def add_toc(doc, *, title: str = "目录", levels: tuple[int, int] = (1, 3), title_font_name: str = "黑体",
                title_font_size: int = 16, title_bold: bool = True, title_alignment="居中",
                toc_level_styles: dict[int, dict] | None = None, use_hyperlinks: bool = True,
                hide_page_numbers_in_web: bool = True, use_outline_levels: bool = False,
                custom_style_levels: dict[str, int] | None = None, page_number_separator: str | None = None,
                right_align_page_numbers: bool | None = None):
        min_level, max_level = levels
        if min_level < 1 or max_level < min_level:
            raise ValueError("levels 参数不合法，例如 (1, 3)")

        WordFormatter._ensure_update_fields_on_open(doc)

        if title:
            title_par = doc.add_paragraph("")
            WordFormatter._set_paragraph_basic_format(
                title_par, alignment=WordFormatter._resolve_alignment(title_alignment), space_before=12, space_after=12
            )
            title_run = title_par.add_run(title)
            WordFormatter._set_run_font(
                title_run, cn_font=title_font_name, size=title_font_size, bold=title_bold
            )

        if toc_level_styles:
            for level, options in toc_level_styles.items():
                WordFormatter.set_toc_level_style(doc, level, **options)

        switches = [f'\\o "{min_level}-{max_level}"', "\\h" if use_hyperlinks else "", "\\z" if hide_page_numbers_in_web else ""]
        if use_outline_levels:
            switches.append("\\u")
        if custom_style_levels:
            style_level_text = ",".join(f"{style_name},{level}" for style_name, level in custom_style_levels.items())
            switches.append(f'\\t "{style_level_text}"')
        if page_number_separator is not None:
            switches.append(f'\\p "{page_number_separator}"')
        if right_align_page_numbers is not None:
            pass
        instruction = f'TOC {" ".join(item for item in switches if item)}'

        toc_par = doc.add_paragraph("")
        WordFormatter._append_field_run(toc_par, instruction)
        return toc_par


about_word = WordFormatter
