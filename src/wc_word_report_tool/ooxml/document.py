"""Direct OOXML document model used by the formatter.

The classes intentionally implement only the public object surface used by
``wc_word_report_tool``. They do not proxy python-docx.
"""

from __future__ import annotations

import posixpath
import re
from copy import deepcopy
from datetime import datetime, timezone
from typing import Iterator

from .enums import (
    CELL_VERTICAL_ALIGNMENT_XML,
    PARAGRAPH_ALIGNMENT_XML,
    ROW_HEIGHT_RULE_XML,
    SECTION_START_REVERSE,
    SECTION_START_XML,
    TAB_ALIGNMENT_XML,
    TABLE_ALIGNMENT_XML,
    WD_CELL_VERTICAL_ALIGNMENT,
    WD_ORIENT,
    WD_PARAGRAPH_ALIGNMENT,
    WD_ROW_HEIGHT_RULE,
    WD_SECTION_START,
    WD_STYLE_TYPE,
    WD_TAB_ALIGNMENT,
    WD_TABLE_ALIGNMENT,
    WD_TABLE_DIRECTION,
    WD_UNDERLINE,
)
from .image import Image, image_stream
from .namespaces import local_name, qn
from .package import (
    COMMENTS_REL,
    CORE_PROPERTIES_REL,
    FOOTER_REL,
    HEADER_REL,
    IMAGE_REL,
    DocumentPackage,
    PackageNotFoundError,
    Part,
)
from .units import Cm, Length, Pt, RGBColor
from .xml import OxmlElement, insert_in_order

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CONTENT_TYPE_HEADER = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
)
CONTENT_TYPE_FOOTER = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
)
CONTENT_TYPE_COMMENTS = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
)
COMMENTS_PARTNAME = "/word/comments.xml"

# Accepted `w:lineRule` values. The high level formatter API documents the same
# strings ("exact", "atLeast"); "auto" is the proportional default.
LINE_SPACING_RULES = {
    "auto": "auto",
    "exact": "exact",
    "atleast": "atLeast",
}


def _twips(value: Length) -> str:
    return str(int(round(value.pt * 20)))


def _cm_from_twips(value: str | None, default: float = 0.0) -> Length:
    if value is None:
        return Cm(default)
    try:
        return Cm(int(value) / 566.9291338582677)
    except (TypeError, ValueError):
        return Cm(default)


def _set_on_off(parent, tag: str, value) -> None:
    """Write an on/off property: True enables, False disables, None removes.

    ``False`` is written as ``w:val="0"`` rather than by dropping the element,
    so a paragraph can override a ``True`` inherited from its style. Omitting
    the element would silently fall back to the inherited value.
    """
    node = parent.find(qn(tag))
    if value is None:
        if node is not None:
            parent.remove(node)
        return
    if node is None:
        node = OxmlElement(tag)
        insert_in_order(parent, node)
    if value:
        node.attrib.pop(qn("w:val"), None)
    else:
        node.set(qn("w:val"), "0")


def _bool_val(element) -> bool:
    if element is None:
        return False
    value = element.get(qn("w:val"))
    return value is None or value.lower() not in {"0", "false", "off", "none"}


def _comment_timestamp() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _parse_timestamp(value: str | None):
    """Parse an OOXML timestamp, returning an aware datetime or None."""
    if not value:
        return None
    for fmt in ("%Y-%m-%dT%H:%M:%SZ", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(value, fmt).replace(tzinfo=timezone.utc)
        except ValueError:
            continue
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None
    return parsed if parsed.tzinfo else parsed.replace(tzinfo=timezone.utc)


def _display_style_name(name: str | None) -> str | None:
    if not name:
        return name
    aliases = {
        "normal": "Normal",
        "header": "Header",
        "footer": "Footer",
        "title": "Title",
        "subtitle": "Subtitle",
    }
    lowered = name.lower()
    if lowered in aliases:
        return aliases[lowered]
    match = re.fullmatch(r"heading ([1-9])", lowered)
    if match:
        return f"Heading {match.group(1)}"
    match = re.fullmatch(r"toc ([1-9])", lowered)
    if match:
        return f"TOC {match.group(1)}"
    return name


def _copy_text(paragraph_element, text: str) -> None:
    run = OxmlElement("w:r")
    node = OxmlElement("w:t")
    if text.startswith(" ") or text.endswith(" "):
        node.set(qn("xml:space"), "preserve")
    node.text = text
    run.append(node)
    paragraph_element.append(run)


class Document:
    def __init__(self, source=None):
        if isinstance(source, DocumentPackage):
            package = source
        elif source is None:
            package = DocumentPackage.default()
        elif isinstance(source, (bytes, bytearray)):
            package = DocumentPackage.from_bytes(bytes(source))
        else:
            package = DocumentPackage.from_path(source)
        self._package = package
        self._part = package.main_document_part
        self._element = self._part.element
        self._body = self._element.find(qn("w:body"))
        if self._body is None:
            raise PackageNotFoundError("DOCX 主文档缺少 w:body")

    @property
    def part(self) -> Part:
        return self._part

    @property
    def package(self) -> DocumentPackage:
        return self._package

    @property
    def element(self):
        return self._element

    @property
    def paragraphs(self) -> list["Paragraph"]:
        return [
            Paragraph(child, self._part, self)
            for child in self._body
            if child.tag == qn("w:p")
        ]

    @property
    def tables(self) -> list["Table"]:
        return [
            Table(child, self._part, self)
            for child in self._body
            if child.tag == qn("w:tbl")
        ]

    @property
    def sections(self) -> list["Section"]:
        sect_prs = []
        for paragraph in self._body.findall(qn("w:p")):
            sect = paragraph.find(qn("w:pPr") + "/" + qn("w:sectPr"))
            if sect is not None:
                sect_prs.append(sect)
        final = self._body.find(qn("w:sectPr"))
        if final is not None:
            sect_prs.append(final)
        if not sect_prs:
            final = OxmlElement("w:sectPr")
            self._body.append(final)
            sect_prs.append(final)
        return [Section(item, self) for item in sect_prs]

    @property
    def styles(self) -> "Styles":
        return Styles(self)

    @property
    def settings(self) -> "Settings":
        return Settings(self)

    @property
    def core_properties(self) -> "CoreProperties":
        return CoreProperties(self)

    @property
    def comments(self) -> "Comments":
        return Comments(self)

    def iter_inner_content(self) -> Iterator["Paragraph | Table"]:
        for child in self._body:
            if child.tag == qn("w:p"):
                yield Paragraph(child, self._part, self)
            elif child.tag == qn("w:tbl"):
                yield Table(child, self._part, self)

    def add_paragraph(self, text: str = "", style: str | None = None) -> "Paragraph":
        element = OxmlElement("w:p")
        sect_pr = self._body.find(qn("w:sectPr"))
        if sect_pr is None:
            self._body.append(element)
        else:
            sect_pr.addprevious(element)
        paragraph = Paragraph(element, self._part, self)
        if style:
            paragraph.style = self.styles[style]
        if text:
            paragraph.add_run(text)
        return paragraph

    def add_heading(self, text: str = "", level: int = 1) -> "Paragraph":
        try:
            paragraph = self.add_paragraph("", style=f"Heading {level}")
        except KeyError:
            paragraph = self.add_paragraph("")
        if text:
            paragraph.add_run(text)
        return paragraph

    def add_page_break(self) -> "Paragraph":
        paragraph = self.add_paragraph("")
        run = paragraph.add_run()
        br = OxmlElement("w:br")
        br.set(qn("w:type"), "page")
        run._r.append(br)
        return paragraph

    def add_table(self, rows: int, cols: int, style: str | None = None) -> "Table":
        element = _new_table_element(rows, cols)
        sect_pr = self._body.find(qn("w:sectPr"))
        if sect_pr is None:
            self._body.append(element)
        else:
            sect_pr.addprevious(element)
        table = Table(element, self._part, self)
        if style:
            table.style = self.styles[style]
        return table

    def add_section(self, start_type=WD_SECTION_START.NEW_PAGE) -> "Section":
        final = self._body.find(qn("w:sectPr"))
        if final is None:
            final = OxmlElement("w:sectPr")
            self._body.append(final)
        break_paragraph = OxmlElement("w:p")
        p_pr = OxmlElement("w:pPr")
        p_pr.append(deepcopy(final))
        break_paragraph.append(p_pr)
        final.addprevious(break_paragraph)
        type_node = final.find(qn("w:type"))
        if type_node is None:
            type_node = OxmlElement("w:type")
            insert_in_order(final, type_node)
        type_node.set(
            qn("w:val"),
            SECTION_START_XML.get(
                WD_SECTION_START(start_type), "nextPage",
            ),
        )
        return Section(final, self)

    def add_picture(self, image, width=None, height=None) -> "InlineShape":
        """Append a picture paragraph and return the new inline shape."""
        paragraph = self.add_paragraph("")
        return paragraph.add_run().add_picture(image, width, height)

    def add_comment(self, runs, text: str | None = "", author: str = "",
                    initials: str | None = "") -> "Comment":
        """Add a comment anchored to the given run or run sequence."""
        run_list = [runs] if isinstance(runs, Run) else list(runs)
        if not run_list:
            raise ValueError("add_comment 需要至少一个 run")
        comment = self.comments.add_comment(
            text=text, author=author, initials=initials,
        )
        run_list[0].mark_comment_range(run_list[-1], comment.comment_id)
        return comment

    @property
    def inline_shapes(self) -> "InlineShapes":
        return InlineShapes(
            (
                drawing
                for drawing in self._body.iter(qn("w:drawing"))
                if drawing.find(".//" + qn("wp:inline")) is not None
            ),
            self._part,
        )

    def save(self, target):
        """Save to a filesystem path or to any writable binary stream."""
        self._part.element = self._element
        if hasattr(target, "write"):
            target.write(self._package.to_bytes())
            return
        self._package.save(target)

    def _add_image(self, part: Part, data: bytes, image: Image) -> tuple[str, int]:
        media = self._package.add_binary_part(
            "word/media", "image", image.ext, data, image.content_type,
        )
        base = posixpath.dirname(part.partname.lstrip("/"))
        target = posixpath.relpath(media.partname.lstrip("/"), start=base)
        rid = self._package.add_relationship(part, IMAGE_REL, target)
        return rid, part.next_id


class Paragraph:
    def __init__(self, element, part: Part, document: Document):
        self._p = element
        self._part = part
        self._document = document

    @property
    def part(self) -> Part:
        return self._part

    @property
    def text(self) -> str:
        chunks = []
        for child in self._p:
            if child.tag == qn("w:r"):
                chunks.append(_run_text(child))
            elif child.tag in (qn("w:ins"), qn("w:hyperlink"), qn("w:smartTag")):
                for run in child.findall(qn("w:r")):
                    chunks.append(_run_text(run))
        return "".join(chunks)

    @property
    def runs(self) -> list["Run"]:
        return [
            Run(child, self) for child in self._p if child.tag == qn("w:r")
        ]

    def add_run(self, text: str | None = None) -> "Run":
        run_element = OxmlElement("w:r")
        self._p.append(run_element)
        run = Run(run_element, self)
        if text:
            run.add_text(text)
        return run

    def insert_paragraph_before(self, text: str | None = None,
                                style=None) -> "Paragraph":
        """Insert a new paragraph directly before this one."""
        element = OxmlElement("w:p")
        self._p.addprevious(element)
        paragraph = Paragraph(element, self._part, self._document)
        if style is not None:
            paragraph.style = (
                self._document.styles[style] if isinstance(style, str) else style
            )
        if text:
            paragraph.add_run(text)
        return paragraph

    @property
    def contains_page_break(self) -> bool:
        """True when a rendered page-break (w:lastRenderedPageBreak) occurs here."""
        return bool(self._p.findall(".//" + qn("w:lastRenderedPageBreak")))

    def iter_inner_content(self):
        """Yield the runs of this paragraph in document order."""
        for child in self._p:
            if child.tag == qn("w:r"):
                yield Run(child, self)
            elif child.tag in (qn("w:hyperlink"), qn("w:ins"), qn("w:smartTag")):
                for run in child.findall(qn("w:r")):
                    yield Run(run, self)

    def clear(self):
        for child in list(self._p):
            if child.tag != qn("w:pPr"):
                self._p.remove(child)

    @property
    def alignment(self):
        jc = self._p.find(qn("w:pPr") + "/" + qn("w:jc"))
        if jc is None:
            return None
        reverse = {
            "left": WD_PARAGRAPH_ALIGNMENT.LEFT,
            "start": WD_PARAGRAPH_ALIGNMENT.LEFT,
            "center": WD_PARAGRAPH_ALIGNMENT.CENTER,
            "right": WD_PARAGRAPH_ALIGNMENT.RIGHT,
            "end": WD_PARAGRAPH_ALIGNMENT.RIGHT,
            "both": WD_PARAGRAPH_ALIGNMENT.JUSTIFY,
            "distribute": WD_PARAGRAPH_ALIGNMENT.DISTRIBUTE,
        }
        return reverse.get(jc.get(qn("w:val")), WD_PARAGRAPH_ALIGNMENT.LEFT)

    @alignment.setter
    def alignment(self, value):
        if value is None:
            p_pr = self._p.find(qn("w:pPr"))
            if p_pr is not None:
                jc = p_pr.find(qn("w:jc"))
                if jc is not None:
                    p_pr.remove(jc)
            return
        try:
            xml_value = PARAGRAPH_ALIGNMENT_XML[WD_PARAGRAPH_ALIGNMENT(value)]
        except (ValueError, KeyError) as exc:
            raise ValueError(f"未知段落对齐方式: {value!r}") from exc
        p_pr = self._p.get_or_add_pPr()
        jc = p_pr.find(qn("w:jc"))
        if jc is None:
            jc = OxmlElement("w:jc")
            insert_in_order(p_pr, jc)
        jc.set(qn("w:val"), xml_value)

    @property
    def paragraph_format(self) -> "ParagraphFormat":
        return ParagraphFormat(self)

    @property
    def style(self):
        style_id = None
        p_style = self._p.find(qn("w:pPr") + "/" + qn("w:pStyle"))
        if p_style is not None:
            style_id = p_style.get(qn("w:val"))
        return self._document.styles.get_by_id(style_id) if style_id else self._document.styles.get_by_id("Normal")

    @style.setter
    def style(self, style):
        style_id = style.style_id if isinstance(style, Style) else str(style)
        p_pr = self._p.get_or_add_pPr()
        node = p_pr.find(qn("w:pStyle"))
        if node is None:
            node = OxmlElement("w:pStyle")
            insert_in_order(p_pr, node)
        node.set(qn("w:val"), style_id)


class Run:
    def __init__(self, element, paragraph: Paragraph):
        self._r = element
        self._paragraph = paragraph

    @property
    def element(self):
        return self._r

    @property
    def _element(self):
        return self._r

    @property
    def part(self) -> Part:
        return self._paragraph.part

    @property
    def text(self) -> str:
        return _run_text(self._r)

    @text.setter
    def text(self, value: str):
        for child in list(self._r):
            if child.tag in (qn("w:t"), qn("w:tab"), qn("w:br"), qn("w:cr")):
                self._r.remove(child)
        if value:
            self.add_text(value)

    def add_text(self, text: str) -> None:
        node = OxmlElement("w:t")
        if text.startswith(" ") or text.endswith(" "):
            node.set(qn("xml:space"), "preserve")
        node.text = text
        self._r.append(node)

    def add_tab(self) -> None:
        self._r.append(OxmlElement("w:tab"))

    def mark_comment_range(self, last_run: "Run", comment_id: int) -> None:
        """Delimit a comment range from this run through last_run."""
        start = OxmlElement("w:commentRangeStart")
        start.set(qn("w:id"), str(comment_id))
        self._r.addprevious(start)

        end = OxmlElement("w:commentRangeEnd")
        end.set(qn("w:id"), str(comment_id))
        last_run._r.addnext(end)

        reference_run = OxmlElement("w:r")
        reference = OxmlElement("w:commentReference")
        reference.set(qn("w:id"), str(comment_id))
        reference_run.append(reference)
        end.addnext(reference_run)

    def clear(self) -> None:
        """Remove all content from this run, keeping character formatting."""
        for child in list(self._r):
            if child.tag != qn("w:rPr"):
                self._r.remove(child)

    @property
    def contains_page_break(self) -> bool:
        return bool(self._r.findall(".//" + qn("w:lastRenderedPageBreak")))

    def iter_inner_content(self):
        """Yield plain-text chunks and inline shapes of this run in order."""
        for child in self._r:
            if child.tag == qn("w:t"):
                yield child.text or ""
            elif child.tag == qn("w:tab"):
                yield "\t"
            elif child.tag in (qn("w:br"), qn("w:cr")):
                yield "\n"
            elif child.tag == qn("w:noBreakHyphen"):
                yield "-"
            elif child.tag == qn("w:drawing"):
                yield InlineShape(child)

    def add_break(self, break_type=None):
        node = OxmlElement("w:br")
        if break_type is not None:
            value = getattr(break_type, "value", break_type)
            if isinstance(value, str):
                node.set(qn("w:type"), value)
        self._r.append(node)

    def add_picture(self, image, width=None, height=None):
        stream, owned, data = image_stream(image)
        try:
            probe = Image.from_blob(data)
        finally:
            if owned:
                stream.close()
        if width is None and height is None:
            width = Cm(probe.px_width / (probe.horz_dpi or 96) * 2.54)
        if width is not None and height is None:
            height = Length(round(int(width) * probe.px_height / probe.px_width))
        elif height is not None and width is None:
            width = Length(round(int(height) * probe.px_width / probe.px_height))
        rid, doc_pr_id = self._paragraph._document._add_image(
            self.part, data, probe,
        )
        drawing = _drawing_xml(rid, int(width), int(height), doc_pr_id, probe)
        self._r.append(drawing)
        return InlineShape(drawing)

    @property
    def font(self) -> "Font":
        return Font(self._r, self)

    @property
    def style(self):
        r_style = self._r.find(qn("w:rStyle"))
        if r_style is None:
            return None
        return self._paragraph._document.styles.get_by_id(r_style.get(qn("w:val")))

    @property
    def underline(self) -> bool:
        return bool(self.font.underline)

    @underline.setter
    def underline(self, value):
        self.font.underline = value

    @property
    def bold(self) -> bool:
        return bool(self.font.bold)

    @bold.setter
    def bold(self, value):
        self.font.bold = value

    @property
    def italic(self) -> bool:
        return bool(self.font.italic)

    @italic.setter
    def italic(self, value):
        self.font.italic = value


class InlineShape:
    """An inline drawing; width/height are EMU Length values."""

    def __init__(self, element):
        self._drawing = element
        self._inline = element.find(".//" + qn("wp:inline"))

    def _extent(self, attribute: str):
        if self._inline is None:
            return None
        extent = self._inline.find(qn("wp:extent"))
        if extent is None:
            return None
        value = extent.get(attribute)
        return Length(int(value)) if value is not None else None

    @property
    def width(self):
        return self._extent("cx")

    @property
    def height(self):
        return self._extent("cy")


class InlineShapes:
    """Collection matching the python-docx access pattern."""

    def __init__(self, drawings, part=None):
        self._items = [InlineShape(element) for element in drawings]
        self._part = part

    @property
    def part(self):
        return self._part

    def __len__(self) -> int:
        return len(self._items)

    def __iter__(self):
        return iter(self._items)

    def __getitem__(self, index):
        return self._items[index]


class ColorFormat:
    def __init__(self, r_pr):
        self._r_pr = r_pr

    @property
    def rgb(self):
        color = self._r_pr.find(qn("w:color"))
        value = color.get(qn("w:val")) if color is not None else None
        if not value or value.lower() == "auto":
            return None
        try:
            return RGBColor(
                int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16),
            )
        except (ValueError, IndexError):
            return None

    @rgb.setter
    def rgb(self, value):
        if value is None:
            color = self._r_pr.find(qn("w:color"))
            if color is not None:
                self._r_pr.remove(color)
            return
        color = self._r_pr.find(qn("w:color"))
        if color is None:
            color = OxmlElement("w:color")
            insert_in_order(self._r_pr, color)
        color.set(qn("w:val"), str(value))


class Font:
    def __init__(self, container, owner):
        self._container = container
        self._owner = owner

    @property
    def _r_pr(self):
        if local_name(self._container.tag) == "rPr":
            return self._container
        return self._container.get_or_add_rPr()

    @property
    def _r_fonts(self):
        r_pr = self._r_pr
        r_fonts = r_pr.find(qn("w:rFonts"))
        if r_fonts is None:
            r_fonts = OxmlElement("w:rFonts")
            insert_in_order(r_pr, r_fonts)
        return r_fonts

    @property
    def name(self):
        value = self._r_pr.find(qn("w:rFonts"))
        return value.get(qn("w:ascii")) if value is not None else None

    @name.setter
    def name(self, value):
        if value is None:
            r_pr = self._r_pr
            node = r_pr.find(qn("w:rFonts"))
            if node is not None:
                r_pr.remove(node)
            return
        fonts = self._r_fonts
        fonts.set(qn("w:ascii"), value)
        fonts.set(qn("w:hAnsi"), value)

    @property
    def size(self):
        node = self._r_pr.find(qn("w:sz"))
        if node is None:
            return None
        try:
            return Pt(int(node.get(qn("w:val"))) / 2)
        except (TypeError, ValueError):
            return None

    @size.setter
    def size(self, value):
        if value is None:
            for tag in ("w:sz", "w:szCs"):
                node = self._r_pr.find(qn(tag))
                if node is not None:
                    self._r_pr.remove(node)
            return
        points = value.pt if isinstance(value, Length) else float(value)
        half_points = str(int(round(points * 2)))
        for tag in ("w:sz", "w:szCs"):
            node = self._r_pr.find(qn(tag))
            if node is None:
                node = OxmlElement(tag)
                insert_in_order(self._r_pr, node)
            node.set(qn("w:val"), half_points)

    def _get_bool(self, tag: str):
        return _bool_val(self._r_pr.find(qn(tag)))

    def _set_bool(self, tag: str, value):
        """True -> bare element, False -> w:val="0", None -> removed.

        This is the same on/off convention used by paragraph properties and by
        python-docx, so a False still overrides an inherited True.
        """
        _set_on_off(self._r_pr, tag, value)

    @property
    def bold(self):
        return self._get_bool("w:b")

    @bold.setter
    def bold(self, value):
        self._set_bool("w:b", value)
        self._set_bool("w:bCs", value)

    @property
    def italic(self):
        return self._get_bool("w:i")

    @italic.setter
    def italic(self, value):
        self._set_bool("w:i", value)
        self._set_bool("w:iCs", value)

    @property
    def all_caps(self):
        return self._get_bool("w:caps")

    @all_caps.setter
    def all_caps(self, value):
        self._set_bool("w:caps", value)

    @property
    def element(self):
        return self._r_pr

    @property
    def part(self):
        return getattr(self._owner, "part", None)

    @property
    def strike(self):
        return self._get_bool("w:strike")

    @strike.setter
    def strike(self, value):
        self._set_bool("w:strike", value)

    @property
    def double_strike(self):
        return self._get_bool("w:dstrike")

    @double_strike.setter
    def double_strike(self, value):
        self._set_bool("w:dstrike", value)

    @property
    def outline(self):
        return self._get_bool("w:outline")

    @outline.setter
    def outline(self, value):
        self._set_bool("w:outline", value)

    @property
    def shadow(self):
        return self._get_bool("w:shadow")

    @shadow.setter
    def shadow(self, value):
        self._set_bool("w:shadow", value)

    @property
    def emboss(self):
        return self._get_bool("w:emboss")

    @emboss.setter
    def emboss(self, value):
        self._set_bool("w:emboss", value)

    @property
    def imprint(self):
        return self._get_bool("w:imprint")

    @imprint.setter
    def imprint(self, value):
        self._set_bool("w:imprint", value)

    @property
    def hidden(self):
        return self._get_bool("w:vanish")

    @hidden.setter
    def hidden(self, value):
        self._set_bool("w:vanish", value)

    @property
    def small_caps(self):
        return self._get_bool("w:smallCaps")

    @small_caps.setter
    def small_caps(self, value):
        self._set_bool("w:smallCaps", value)

    @property
    def snap_to_grid(self):
        return self._get_bool("w:snapToGrid")

    @snap_to_grid.setter
    def snap_to_grid(self, value):
        self._set_bool("w:snapToGrid", value)

    @property
    def spec_vanish(self):
        return self._get_bool("w:specVanish")

    @spec_vanish.setter
    def spec_vanish(self, value):
        self._set_bool("w:specVanish", value)

    @property
    def web_hidden(self):
        return self._get_bool("w:webHidden")

    @web_hidden.setter
    def web_hidden(self, value):
        self._set_bool("w:webHidden", value)

    @property
    def rtl(self):
        return self._get_bool("w:rtl")

    @rtl.setter
    def rtl(self, value):
        self._set_bool("w:rtl", value)

    @property
    def no_proof(self):
        return self._get_bool("w:noProof")

    @no_proof.setter
    def no_proof(self, value):
        self._set_bool("w:noProof", value)

    @property
    def math(self):
        return self._get_bool("w:oMath")

    @math.setter
    def math(self, value):
        self._set_bool("w:oMath", value)

    @property
    def cs_bold(self):
        return self._get_bool("w:bCs")

    @cs_bold.setter
    def cs_bold(self, value):
        self._set_bool("w:bCs", value)

    @property
    def cs_italic(self):
        return self._get_bool("w:iCs")

    @cs_italic.setter
    def cs_italic(self, value):
        self._set_bool("w:iCs", value)

    @property
    def complex_script(self):
        return self._get_bool("w:cs")

    @complex_script.setter
    def complex_script(self, value):
        self._set_bool("w:cs", value)

    def _get_vert_align(self):
        node = self._r_pr.find(qn("w:vertAlign"))
        return node.get(qn("w:val")) if node is not None else None

    def _set_vert_align(self, value: str, enable) -> None:
        node = self._r_pr.find(qn("w:vertAlign"))
        current = node.get(qn("w:val")) if node is not None else None
        if enable:
            if node is None:
                node = OxmlElement("w:vertAlign")
                insert_in_order(self._r_pr, node)
            node.set(qn("w:val"), value)
        elif current == value and node is not None:
            self._r_pr.remove(node)

    @property
    def subscript(self):
        return self._get_vert_align() == "subscript"

    @subscript.setter
    def subscript(self, value):
        self._set_vert_align("subscript", value)

    @property
    def superscript(self):
        return self._get_vert_align() == "superscript"

    @superscript.setter
    def superscript(self, value):
        self._set_vert_align("superscript", value)

    @property
    def color(self):
        return ColorFormat(self._r_pr)

    @property
    def highlight_color(self):
        node = self._r_pr.find(qn("w:highlight"))
        return node is not None

    @highlight_color.setter
    def highlight_color(self, value):
        node = self._r_pr.find(qn("w:highlight"))
        if value is None:
            if node is not None:
                self._r_pr.remove(node)
            return
        if node is None:
            node = OxmlElement("w:highlight")
            insert_in_order(self._r_pr, node)
        node.set(qn("w:val"), getattr(value, "value", str(value)))

    @property
    def underline(self):
        node = self._r_pr.find(qn("w:u"))
        if node is None:
            return None
        value = node.get(qn("w:val"), "single")
        return False if value == "none" else True

    @underline.setter
    def underline(self, value):
        node = self._r_pr.find(qn("w:u"))
        if value in (None, False):
            if value is False:
                if node is None:
                    node = OxmlElement("w:u")
                    insert_in_order(self._r_pr, node)
                node.set(qn("w:val"), "none")
            elif node is not None:
                self._r_pr.remove(node)
            return
        if node is None:
            node = OxmlElement("w:u")
            insert_in_order(self._r_pr, node)
        if value is True:
            xml_value = "single"
        elif isinstance(value, WD_UNDERLINE):
            xml_value = value.value
        else:
            xml_value = str(value)
        node.set(qn("w:val"), xml_value)


class ParagraphFormat:
    def __init__(self, paragraph: Paragraph):
        self._paragraph = paragraph

    @property
    def _p_pr(self):
        return self._paragraph._p.get_or_add_pPr()

    @property
    def element(self):
        return self._p_pr

    @property
    def part(self):
        return self._paragraph.part

    def _get_spacing(self):
        return self._p_pr.find(qn("w:spacing"))

    def _ensure_spacing(self):
        spacing = self._get_spacing()
        if spacing is None:
            spacing = OxmlElement("w:spacing")
            insert_in_order(self._p_pr, spacing)
        return spacing

    @property
    def alignment(self):
        jc = self._p_pr.find(qn("w:jc"))
        if jc is None:
            return None
        reverse = {
            "left": WD_PARAGRAPH_ALIGNMENT.LEFT,
            "start": WD_PARAGRAPH_ALIGNMENT.LEFT,
            "center": WD_PARAGRAPH_ALIGNMENT.CENTER,
            "right": WD_PARAGRAPH_ALIGNMENT.RIGHT,
            "end": WD_PARAGRAPH_ALIGNMENT.RIGHT,
            "both": WD_PARAGRAPH_ALIGNMENT.JUSTIFY,
            "distribute": WD_PARAGRAPH_ALIGNMENT.DISTRIBUTE,
        }
        return reverse.get(jc.get(qn("w:val")), WD_PARAGRAPH_ALIGNMENT.LEFT)

    @alignment.setter
    def alignment(self, value):
        if value is None:
            jc = self._p_pr.find(qn("w:jc"))
            if jc is not None:
                self._p_pr.remove(jc)
            return
        jc = self._p_pr.find(qn("w:jc"))
        if jc is None:
            jc = OxmlElement("w:jc")
            insert_in_order(self._p_pr, jc)
        jc.set(qn("w:val"), PARAGRAPH_ALIGNMENT_XML[WD_PARAGRAPH_ALIGNMENT(value)])

    @property
    def line_spacing(self):
        spacing = self._get_spacing()
        if spacing is None:
            return None
        line = spacing.get(qn("w:line"))
        if line is None:
            return None
        try:
            number = int(line)
        except ValueError:
            return None
        if spacing.get(qn("w:lineRule")) in {"exact", "atLeast"}:
            return Pt(number / 20)
        return number / 240

    @line_spacing.setter
    def line_spacing(self, value):
        if value is None:
            spacing = self._get_spacing()
            if spacing is not None:
                spacing.attrib.pop(qn("w:line"), None)
                spacing.attrib.pop(qn("w:lineRule"), None)
            return
        spacing = self._ensure_spacing()
        if isinstance(value, Length):
            spacing.set(qn("w:line"), str(int(round(value.pt * 20))))
            spacing.set(qn("w:lineRule"), "exact")
        else:
            spacing.set(qn("w:line"), str(int(round(float(value) * 240))))
            spacing.set(qn("w:lineRule"), "auto")

    @property
    def line_spacing_rule(self):
        spacing = self._get_spacing()
        if spacing is None:
            return None
        return spacing.get(qn("w:lineRule"))

    @line_spacing_rule.setter
    def line_spacing_rule(self, value):
        if value is None:
            spacing = self._get_spacing()
            if spacing is not None:
                spacing.attrib.pop(qn("w:lineRule"), None)
            return
        raw = getattr(value, "value", value)
        rule = LINE_SPACING_RULES.get(str(raw).strip().lower())
        if rule is None:
            raise ValueError(
                "unsupported line spacing rule %r; expected "
                "'auto', 'exact' or 'atLeast'" % (raw,)
            )
        self._ensure_spacing().set(qn("w:lineRule"), rule)

    def _get_length(self, tag: str):
        node = self._p_pr.find(qn(tag))
        if node is None:
            return None
        return _cm_from_twips(node.get(qn("w:w")))

    def _set_length(self, tag: str, value):
        if value is None:
            node = self._p_pr.find(qn(tag))
            if node is not None:
                self._p_pr.remove(node)
            return
        node = self._p_pr.find(qn(tag))
        if node is None:
            node = OxmlElement(tag)
            insert_in_order(self._p_pr, node)
        node.set(qn("w:w"), _twips(value))
        if tag == "w:ind":
            node.set(qn("w:firstLineChars"), "0")

    @property
    def left_indent(self):
        return self._get_length("w:ind")

    @left_indent.setter
    def left_indent(self, value):
        ind = self._p_pr.find(qn("w:ind"))
        if value is None:
            if ind is not None:
                ind.attrib.pop(qn("w:left"), None)
            return
        if ind is None:
            ind = OxmlElement("w:ind")
            insert_in_order(self._p_pr, ind)
        ind.set(qn("w:left"), _twips(value))

    @property
    def right_indent(self):
        ind = self._p_pr.find(qn("w:ind"))
        return _cm_from_twips(ind.get(qn("w:right"))) if ind is not None and ind.get(qn("w:right")) else None

    @right_indent.setter
    def right_indent(self, value):
        ind = self._p_pr.find(qn("w:ind"))
        if value is None:
            if ind is not None:
                ind.attrib.pop(qn("w:right"), None)
            return
        if ind is None:
            ind = OxmlElement("w:ind")
            insert_in_order(self._p_pr, ind)
        ind.set(qn("w:right"), _twips(value))

    @property
    def first_line_indent(self):
        ind = self._p_pr.find(qn("w:ind"))
        value = ind.get(qn("w:firstLine")) if ind is not None else None
        return _cm_from_twips(value) if value else None

    @first_line_indent.setter
    def first_line_indent(self, value):
        ind = self._p_pr.find(qn("w:ind"))
        if value is None:
            if ind is not None:
                ind.attrib.pop(qn("w:firstLine"), None)
            return
        if ind is None:
            ind = OxmlElement("w:ind")
            insert_in_order(self._p_pr, ind)
        ind.set(qn("w:firstLine"), _twips(value))

    def _set_bool(self, tag, value):
        _set_on_off(self._p_pr, tag, value)

    @property
    def page_break_before(self):
        return _bool_val(self._p_pr.find(qn("w:pageBreakBefore")))

    @page_break_before.setter
    def page_break_before(self, value):
        self._set_bool("w:pageBreakBefore", value)

    @property
    def keep_with_next(self):
        return _bool_val(self._p_pr.find(qn("w:keepNext")))

    @keep_with_next.setter
    def keep_with_next(self, value):
        self._set_bool("w:keepNext", value)

    @property
    def keep_together(self):
        return _bool_val(self._p_pr.find(qn("w:keepLines")))

    @keep_together.setter
    def keep_together(self, value):
        self._set_bool("w:keepLines", value)

    @property
    def widow_control(self):
        return _bool_val(self._p_pr.find(qn("w:widowControl")))

    @widow_control.setter
    def widow_control(self, value):
        self._set_bool("w:widowControl", value)

    @property
    def space_before(self):
        spacing = self._get_spacing()
        value = spacing.get(qn("w:before")) if spacing is not None else None
        return Pt(int(value) / 20) if value not in (None, "0") else Pt(0)

    @space_before.setter
    def space_before(self, value):
        if value is None:
            return
        spacing = self._ensure_spacing()
        spacing.set(qn("w:before"), str(int(round(value.pt * 20))))

    @property
    def space_after(self):
        spacing = self._get_spacing()
        value = spacing.get(qn("w:after")) if spacing is not None else None
        return Pt(int(value) / 20) if value not in (None, "0") else Pt(0)

    @space_after.setter
    def space_after(self, value):
        if value is None:
            return
        spacing = self._ensure_spacing()
        spacing.set(qn("w:after"), str(int(round(value.pt * 20))))

    @property
    def tab_stops(self):
        return TabStops(self)


class TabStops:
    def __init__(self, paragraph_format: ParagraphFormat):
        self._format = paragraph_format

    @property
    def _tabs(self):
        p_pr = self._format._p_pr
        tabs = p_pr.find(qn("w:tabs"))
        if tabs is None:
            tabs = OxmlElement("w:tabs")
            insert_in_order(p_pr, tabs)
        return tabs

    def clear_all(self):
        p_pr = self._format._p_pr
        tabs = p_pr.find(qn("w:tabs"))
        if tabs is not None:
            p_pr.remove(tabs)

    def add_tab_stop(self, position, alignment=WD_TAB_ALIGNMENT.LEFT, leader=None):
        tabs = self._tabs
        node = OxmlElement("w:tab")
        node.set(qn("w:val"), TAB_ALIGNMENT_XML.get(alignment, "left"))
        node.set(qn("w:pos"), _twips(position))
        if leader is not None:
            node.set(qn("w:leader"), getattr(leader, "value", str(leader)))
        tabs.append(node)


class Style:
    def __init__(self, element, document: Document):
        self.element = element
        self._document = document

    @property
    def part(self):
        return self._document.styles.part

    @property
    def style_id(self):
        return self.element.get(qn("w:styleId"))

    @property
    def name(self):
        node = self.element.find(qn("w:name"))
        raw = node.get(qn("w:val")) if node is not None else self.style_id
        return _display_style_name(raw)

    @property
    def type(self):
        raw = self.element.get(qn("w:type"), "paragraph")
        return {
            "paragraph": WD_STYLE_TYPE.PARAGRAPH,
            "character": WD_STYLE_TYPE.CHARACTER,
            "table": WD_STYLE_TYPE.TABLE,
            "numbering": WD_STYLE_TYPE.LIST,
        }.get(raw, WD_STYLE_TYPE.PARAGRAPH)

    @property
    def base_style(self):
        node = self.element.find(qn("w:basedOn"))
        if node is None:
            return None
        return self._document.styles.get_by_id(node.get(qn("w:val")))

    @base_style.setter
    def base_style(self, style):
        node = self.element.find(qn("w:basedOn"))
        if node is None:
            node = OxmlElement("w:basedOn")
            insert_in_order(self.element, node)
        node.set(qn("w:val"), style.style_id if isinstance(style, Style) else str(style))

    @property
    def next_paragraph_style(self):
        node = self.element.find(qn("w:next"))
        if node is None:
            return None
        return self._document.styles.get_by_id(node.get(qn("w:val")))

    @property
    def font(self):
        r_pr = self.element.find(qn("w:rPr"))
        if r_pr is None:
            r_pr = OxmlElement("w:rPr")
            insert_in_order(self.element, r_pr)
        return Font(r_pr, self)

    @property
    def paragraph_format(self):
        p_pr = self.element.find(qn("w:pPr"))
        if p_pr is None:
            p_pr = OxmlElement("w:pPr")
            insert_in_order(self.element, p_pr)
        return StyleParagraphFormat(self, p_pr)


class StyleParagraphFormat(ParagraphFormat):
    def __init__(self, style: Style, p_pr):
        self._style = style
        self._p_pr_element = p_pr

    @property
    def _p_pr(self):
        return self._p_pr_element


class Styles:
    def __init__(self, document: Document):
        self._document = document
        part = document.package.part_for_name("/word/styles.xml")
        if part is None:
            raise PackageNotFoundError("DOCX 缺少 styles.xml")
        self._part = part
        self.element = part.element

    @property
    def part(self):
        return self._part

    def _styles(self):
        return [
            Style(node, self._document)
            for node in self.element.findall(qn("w:style"))
        ]

    def __iter__(self):
        return iter(self._styles())

    def __len__(self):
        return len(self.element.findall(qn("w:style")))

    def __getitem__(self, key):
        if isinstance(key, int):
            return self._styles()[key]
        for style in self._styles():
            if (
                style.name == key
                or style.name.lower() == str(key).lower()
                or style.style_id == key
            ):
                return style
        raise KeyError(f"样式不存在: {key}")

    def get_by_id(self, style_id: str | None):
        if not style_id:
            for style in self._styles():
                if style.name == "Normal":
                    return style
            return self._styles()[0] if self._styles() else None
        for style in self._styles():
            if style.style_id == style_id:
                return style
        return None

    def add_style(self, name: str, style_type):
        existing = next((style for style in self._styles() if style.name == name), None)
        if existing is not None:
            return existing
        style_id = re.sub(r"[^\w\u4e00-\u9fff-]", "", name) or "Style"
        used = {style.style_id for style in self._styles()}
        base = style_id
        index = 1
        while style_id in used:
            style_id = f"{base}{index}"
            index += 1
        node = OxmlElement("w:style")
        node.set(qn("w:type"), {
            WD_STYLE_TYPE.PARAGRAPH: "paragraph",
            WD_STYLE_TYPE.CHARACTER: "character",
            WD_STYLE_TYPE.TABLE: "table",
            WD_STYLE_TYPE.LIST: "numbering",
        }.get(style_type, "paragraph"))
        node.set(qn("w:styleId"), style_id)
        name_node = OxmlElement("w:name")
        name_node.set(qn("w:val"), name)
        node.append(name_node)
        self.element.append(node)
        return Style(node, self._document)


class Settings:
    def __init__(self, document: Document):
        part = document.package.part_for_name("/word/settings.xml")
        if part is None:
            raise PackageNotFoundError("DOCX 缺少 settings.xml")
        self._part = part
        self.element = part.element


class CoreProperties:
    def __init__(self, document: Document):
        self._document = document
        part = document.package.part_for_name("/docProps/core.xml")
        if part is None:
            root = OxmlElement("cp:coreProperties")
            part = document.package.add_xml_part(
                "docProps", "core", root,
                "application/vnd.openxmlformats-package.core-properties+xml",
            )
            document.package.add_package_relationship(
                CORE_PROPERTIES_REL, "docProps/core.xml",
            )
        self._part = part
        self.element = part.element

    def _get(self, tag: str) -> str:
        node = self.element.find(qn(tag))
        return node.text or "" if node is not None else ""

    def _set(self, tag: str, value: str):
        node = self.element.find(qn(tag))
        if node is None:
            node = OxmlElement(tag)
            self.element.append(node)
        node.text = value

    @property
    def title(self):
        return self._get("dc:title")

    @title.setter
    def title(self, value):
        self._set("dc:title", value)

    @property
    def author(self):
        return self._get("dc:creator")

    @author.setter
    def author(self, value):
        self._set("dc:creator", value)


class Comment:
    """A single comment entry stored in the comments part."""

    def __init__(self, element, document: "Document"):
        self._element = element
        self._document = document

    @property
    def element(self):
        return self._element

    @property
    def comment_id(self):
        value = self._element.get(qn("w:id"))
        return int(value) if value is not None else None

    @property
    def author(self) -> str:
        return self._element.get(qn("w:author")) or ""

    @property
    def initials(self):
        return self._element.get(qn("w:initials"))

    @property
    def timestamp(self):
        return _parse_timestamp(self._element.get(qn("w:date")))

    @property
    def paragraphs(self) -> list["Paragraph"]:
        return [
            Paragraph(child, self._document.part, self._document)
            for child in self._element
            if child.tag == qn("w:p")
        ]

    def add_paragraph(self, text: str = "") -> "Paragraph":
        element = OxmlElement("w:p")
        self._element.append(element)
        paragraph = Paragraph(element, self._document.part, self._document)
        if text:
            paragraph.add_run(text)
        return paragraph

    @property
    def text(self) -> str:
        return "\n".join(paragraph.text for paragraph in self.paragraphs)


class Comments:
    """The document comment collection; its part is created on first use."""

    def __init__(self, document: "Document"):
        self._document = document

    def _comments_relationship(self):
        package = self._document.package
        for rel in package.get_relationships(self._document.part).items.values():
            if rel.reltype == COMMENTS_REL:
                return rel
        return None

    def _existing_part(self):
        """Find the comments part; the relationship wins over its file name.

        Documents produced by merge tools may store comments under a
        non-canonical name such as comments1.xml, so the part is resolved
        through the relationship and only then via the default name.
        """
        package = self._document.package
        rel = self._comments_relationship()
        if rel is not None:
            target = package.resolve_target(self._document.part, rel)
            part = package.part_for_name(target) if target else None
            if part is not None:
                return part
        return package.part_for_name(COMMENTS_PARTNAME)

    def _ensure_part(self):
        part = self._existing_part()
        if part is not None:
            return part

        package = self._document.package
        rel = self._comments_relationship()
        if rel is not None:
            # The relationship exists but its part does not: recreate the part
            # under the referenced name instead of adding a second relationship.
            target = package.resolve_target(self._document.part, rel) or COMMENTS_PARTNAME
            partname = "/" + target.lstrip("/")
            part = Part(package, partname, CONTENT_TYPE_COMMENTS, b"")
            part.element = OxmlElement("w:comments")
            package.parts[partname] = part
            package.content_types[partname] = CONTENT_TYPE_COMMENTS
            return part

        root = OxmlElement("w:comments")
        part = package.add_xml_part("word", "comments", root, CONTENT_TYPE_COMMENTS)
        package.add_relationship(
            self._document.part,
            COMMENTS_REL,
            posixpath.basename(part.partname),
        )
        return part

    def _elements(self):
        part = self._existing_part()
        if part is None:
            return []
        return part.element.findall(qn("w:comment"))

    def __len__(self) -> int:
        return len(self._elements())

    def __iter__(self):
        return iter(
            [Comment(element, self._document) for element in self._elements()]
        )

    def __getitem__(self, index):
        return list(self)[index]

    def get(self, comment_id: int):
        for comment in self:
            if comment.comment_id == comment_id:
                return comment
        return None

    def add_comment(self, text: str = "", author: str = "",
                    initials: str | None = "") -> "Comment":
        part = self._ensure_part()
        element = OxmlElement("w:comment")
        element.set(qn("w:id"), str(self._next_id()))
        element.set(qn("w:author"), author)
        if initials is not None:
            element.set(qn("w:initials"), initials)
        element.set(qn("w:date"), _comment_timestamp())
        element.append(OxmlElement("w:p"))

        comment = Comment(element, self._document)
        part.element.append(element)

        if text:
            lines = str(text).split("\n")
            comment.paragraphs[0].add_run(lines[0])
            for line in lines[1:]:
                comment.add_paragraph(line)
        return comment

    def _next_id(self) -> int:
        ids = []
        for element in self._elements():
            value = element.get(qn("w:id"))
            if value is not None:
                try:
                    ids.append(int(value))
                except ValueError:
                    pass
        return max(ids, default=-1) + 1


class Section:
    def __init__(self, sect_pr, document: Document):
        self._sectPr = sect_pr
        self._document = document

    def _pg_size(self):
        node = self._sectPr.find(qn("w:pgSz"))
        if node is None:
            node = OxmlElement("w:pgSz")
            insert_in_order(self._sectPr, node)
        return node

    def _pg_mar(self):
        node = self._sectPr.find(qn("w:pgMar"))
        if node is None:
            node = OxmlElement("w:pgMar")
            insert_in_order(self._sectPr, node)
        return node

    def _get_size(self, attr: str, default_cm: float):
        value = self._pg_size().get(qn(attr))
        return _cm_from_twips(value, default_cm)

    def _set_size(self, attr: str, value):
        self._pg_size().set(qn(attr), _twips(value))

    @property
    def part(self):
        return self._document.part

    @property
    def start_type(self):
        node = self._sectPr.find(qn("w:type"))
        if node is None:
            return None
        return SECTION_START_REVERSE.get(node.get(qn("w:val")))

    @start_type.setter
    def start_type(self, value):
        node = self._sectPr.find(qn("w:type"))
        if value is None:
            if node is not None:
                self._sectPr.remove(node)
            return
        if node is None:
            node = OxmlElement("w:type")
            insert_in_order(self._sectPr, node)
        node.set(qn("w:val"), SECTION_START_XML[WD_SECTION_START(value)])

    @property
    def page_width(self):
        return self._get_size("w:w", 21.0)

    @page_width.setter
    def page_width(self, value):
        self._set_size("w:w", value)

    @property
    def page_height(self):
        return self._get_size("w:h", 29.7)

    @page_height.setter
    def page_height(self, value):
        self._set_size("w:h", value)

    @property
    def orientation(self):
        return (
            WD_ORIENT.LANDSCAPE
            if self._pg_size().get(qn("w:orient")) == "landscape"
            else WD_ORIENT.PORTRAIT
        )

    @orientation.setter
    def orientation(self, value):
        self._pg_size().set(
            qn("w:orient"),
            "landscape" if WD_ORIENT(value) == WD_ORIENT.LANDSCAPE else "portrait",
        )

    def _get_margin(self, name: str, default: float):
        return _cm_from_twips(self._pg_mar().get(qn(name)), default)

    def _set_margin(self, name: str, value):
        self._pg_mar().set(qn(name), _twips(value))

    @property
    def left_margin(self):
        return self._get_margin("w:left", 3.18)

    @left_margin.setter
    def left_margin(self, value):
        self._set_margin("w:left", value)

    @property
    def right_margin(self):
        return self._get_margin("w:right", 3.18)

    @right_margin.setter
    def right_margin(self, value):
        self._set_margin("w:right", value)

    @property
    def top_margin(self):
        return self._get_margin("w:top", 2.54)

    @top_margin.setter
    def top_margin(self, value):
        self._set_margin("w:top", value)

    @property
    def bottom_margin(self):
        return self._get_margin("w:bottom", 2.54)

    @bottom_margin.setter
    def bottom_margin(self, value):
        self._set_margin("w:bottom", value)

    @property
    def gutter(self):
        return self._get_margin("w:gutter", 0)

    @gutter.setter
    def gutter(self, value):
        self._set_margin("w:gutter", value)

    @property
    def header_distance(self):
        return self._get_margin("w:header", 1.25)

    @header_distance.setter
    def header_distance(self, value):
        self._set_margin("w:header", value)

    @property
    def footer_distance(self):
        return self._get_margin("w:footer", 1.25)

    @footer_distance.setter
    def footer_distance(self, value):
        self._set_margin("w:footer", value)

    @property
    def different_first_page_header_footer(self):
        return self._sectPr.find(qn("w:titlePg")) is not None

    @different_first_page_header_footer.setter
    def different_first_page_header_footer(self, value):
        node = self._sectPr.find(qn("w:titlePg"))
        if value and node is None:
            node = OxmlElement("w:titlePg")
            insert_in_order(self._sectPr, node)
        elif not value and node is not None:
            self._sectPr.remove(node)

    @property
    def header(self):
        return HeaderFooter(self, "header", "default")

    @property
    def footer(self):
        return HeaderFooter(self, "footer", "default")

    @property
    def first_page_header(self):
        return HeaderFooter(self, "header", "first")

    @property
    def first_page_footer(self):
        return HeaderFooter(self, "footer", "first")

    @property
    def even_page_header(self):
        return HeaderFooter(self, "header", "even")

    @property
    def even_page_footer(self):
        return HeaderFooter(self, "footer", "even")

    def iter_inner_content(self):
        body = self._document._body
        blocks = [
            child for child in body
            if child.tag in (qn("w:p"), qn("w:tbl"))
        ]
        end = len(blocks)
        current = self._sectPr
        parent = current.getparent()
        if parent is not None and parent.tag == qn("w:pPr"):
            paragraph = parent.getparent()
            if paragraph in blocks:
                end = blocks.index(paragraph) + 1
        start = 0
        all_sections = self._document.sections
        for previous in all_sections:
            if previous is self:
                break
            anchor = previous._sectPr.getparent()
            if anchor is not None and anchor.tag == qn("w:pPr"):
                p = anchor.getparent()
                if p in blocks:
                    start = blocks.index(p) + 1
        for child in blocks[start:end]:
            if child.tag == qn("w:p"):
                yield Paragraph(child, self._document.part, self._document)
            else:
                yield Table(child, self._document.part, self._document)


class HeaderFooter:
    def __init__(self, section: Section, target: str, variant: str):
        self._section = section
        self._target = target
        self._variant = variant
        self._part = None
        self._linked = True

    def _reference(self):
        tag = qn("w:headerReference") if self._target == "header" else qn("w:footerReference")
        return next(
            (
                node for node in self._section._sectPr.findall(tag)
                if node.get(qn("w:type"), "default") == self._variant
            ),
            None,
        )

    def _ensure_part(self):
        reference = self._reference()
        if reference is not None:
            rid = reference.get(qn("r:id"))
            rel = self._section._document.part.rels.get(rid)
            if rel is not None:
                target = self._section._document.package.resolve_target(
                    self._section._document.part, rel,
                )
                if target:
                    self._part = self._section._document.package.part_for_name(target)
                    self._linked = False
                    return self._part
        root = OxmlElement("w:hdr" if self._target == "header" else "w:ftr")
        root.append(OxmlElement("w:p"))
        self._part = self._section._document.package.add_xml_part(
            "word", self._target, root,
            CONTENT_TYPE_HEADER if self._target == "header" else CONTENT_TYPE_FOOTER,
        )
        rid = self._section._document.package.add_relationship(
            self._section._document.part,
            HEADER_REL if self._target == "header" else FOOTER_REL,
            posixpath.basename(self._part.partname),
        )
        reference = OxmlElement(
            "w:headerReference" if self._target == "header" else "w:footerReference",
        )
        reference.set(qn("w:type"), self._variant)
        reference.set(qn("r:id"), rid)
        insert_in_order(self._section._sectPr, reference)
        self._linked = False
        return self._part

    @property
    def is_linked_to_previous(self) -> bool:
        return self._reference() is None

    @is_linked_to_previous.setter
    def is_linked_to_previous(self, value: bool):
        if value:
            reference = self._reference()
            if reference is not None:
                self._section._sectPr.remove(reference)
            self._linked = True
        else:
            self._ensure_part()

    @property
    def part(self):
        if self._part is None:
            self._ensure_part()
        return self._part

    @property
    def _element(self):
        return self.part.element

    @property
    def paragraphs(self):
        return [
            Paragraph(node, self.part, self._section._document)
            for node in self.part.element
            if node.tag == qn("w:p")
        ]

    def add_paragraph(self, text: str = "") -> Paragraph:
        node = OxmlElement("w:p")
        self.part.element.append(node)
        paragraph = Paragraph(node, self.part, self._section._document)
        if text:
            paragraph.add_run(text)
        return paragraph


class Table:
    def __init__(self, element, part: Part, document: Document):
        self._tbl = element
        self._part = part
        self._document = document

    @property
    def part(self):
        return self._part

    @property
    def table(self):
        return self

    @property
    def table_direction(self):
        node = self._tbl.tblPr.find(qn("w:bidiVisual"))
        if node is None:
            return None
        return WD_TABLE_DIRECTION.RTL if _bool_val(node) else WD_TABLE_DIRECTION.LTR

    @table_direction.setter
    def table_direction(self, value):
        tbl_pr = self._tbl.tblPr
        node = tbl_pr.find(qn("w:bidiVisual"))
        if value is None:
            if node is not None:
                tbl_pr.remove(node)
            return
        if node is None:
            node = OxmlElement("w:bidiVisual")
            insert_in_order(tbl_pr, node)
        if WD_TABLE_DIRECTION(value) == WD_TABLE_DIRECTION.RTL:
            node.attrib.pop(qn("w:val"), None)
        else:
            node.set(qn("w:val"), "0")

    @property
    def rows(self):
        return [Row(node, self) for node in self._tbl.findall(qn("w:tr"))]

    @property
    def columns(self):
        grid = self._tbl.find(qn("w:tblGrid"))
        count = len(grid.findall(qn("w:gridCol"))) if grid is not None else 0
        return [Column(self, index) for index in range(count)]

    def cell(self, row: int, col: int):
        matrix = _table_grid(self)
        if row < 0:
            row += len(matrix)
        if col < 0:
            col += len(self.columns)
        if row < 0 or row >= len(matrix) or col < 0 or col >= len(matrix[row]):
            raise IndexError(f"表格坐标越界: ({row}, {col})")
        cell, source_row, source_col = matrix[row][col]
        # Follow a vertical merge back to its restart cell.
        while _vmerge_state(cell) == "continue" and source_row > 0:
            source_row -= 1
            cell, _, _ = matrix[source_row][source_col]
        return Cell(cell, self)

    @property
    def alignment(self):
        jc = self._tbl.tblPr.find(qn("w:jc"))
        return {
            "center": WD_TABLE_ALIGNMENT.CENTER,
            "right": WD_TABLE_ALIGNMENT.RIGHT,
            "left": WD_TABLE_ALIGNMENT.LEFT,
        }.get(jc.get(qn("w:val")) if jc is not None else None, WD_TABLE_ALIGNMENT.LEFT)

    @alignment.setter
    def alignment(self, value):
        jc = self._tbl.tblPr.find(qn("w:jc"))
        if jc is None:
            jc = OxmlElement("w:jc")
            insert_in_order(self._tbl.tblPr, jc)
        jc.set(qn("w:val"), TABLE_ALIGNMENT_XML[WD_TABLE_ALIGNMENT(value)])

    @property
    def style(self):
        node = self._tbl.tblPr.find(qn("w:tblStyle"))
        if node is None:
            return None
        return self._document.styles.get_by_id(node.get(qn("w:val")))

    @style.setter
    def style(self, value):
        if isinstance(value, Style):
            style_id = value.style_id
        else:
            try:
                style_id = self._document.styles[str(value)].style_id
            except KeyError:
                style_id = str(value)
        node = self._tbl.tblPr.find(qn("w:tblStyle"))
        if node is None:
            node = OxmlElement("w:tblStyle")
            insert_in_order(self._tbl.tblPr, node)
        node.set(qn("w:val"), style_id)

    @property
    def autofit(self):
        layout = self._tbl.tblPr.find(qn("w:tblLayout"))
        return layout is None or layout.get(qn("w:type")) != "fixed"

    @autofit.setter
    def autofit(self, value):
        layout = self._tbl.tblPr.find(qn("w:tblLayout"))
        if value:
            if layout is not None:
                self._tbl.tblPr.remove(layout)
            return
        if layout is None:
            layout = OxmlElement("w:tblLayout")
            insert_in_order(self._tbl.tblPr, layout)
        layout.set(qn("w:type"), "fixed")

    def add_row(self):
        template = self._tbl.find(qn("w:tr"))
        row = OxmlElement("w:tr")
        if template is not None:
            for _ in template.findall(qn("w:tc")):
                cell = OxmlElement("w:tc")
                cell.append(OxmlElement("w:tcPr"))
                cell.append(OxmlElement("w:p"))
                row.append(cell)
        self._tbl.append(row)
        return Row(row, self)


class Column:
    def __init__(self, table: Table, index: int):
        self.table = table
        self.index = index


class Row:
    def __init__(self, element, table: Table):
        self._tr = element
        self.table = table

    @property
    def cells(self):
        row_index = list(self.table._tbl.findall(qn("w:tr"))).index(self._tr)
        matrix = _table_grid(self.table)
        return [
            Cell(matrix[row_index][col][0], self.table)
            for col in range(len(matrix[row_index]))
        ]

    @property
    def height(self):
        node = self._tr.get_or_add_trPr().find(qn("w:trHeight"))
        if node is None:
            return None
        return _cm_from_twips(node.get(qn("w:val")))

    @height.setter
    def height(self, value):
        node = self._tr.get_or_add_trPr().find(qn("w:trHeight"))
        if node is None:
            node = OxmlElement("w:trHeight")
            insert_in_order(self._tr.get_or_add_trPr(), node)
        node.set(qn("w:val"), _twips(value))

    @property
    def height_rule(self):
        node = self._tr.get_or_add_trPr().find(qn("w:trHeight"))
        raw = node.get(qn("w:hRule")) if node is not None else None
        return {
            "atLeast": WD_ROW_HEIGHT_RULE.AT_LEAST,
            "exact": WD_ROW_HEIGHT_RULE.EXACTLY,
        }.get(raw, WD_ROW_HEIGHT_RULE.AUTO)

    @height_rule.setter
    def height_rule(self, value):
        node = self._tr.get_or_add_trPr().find(qn("w:trHeight"))
        if node is None:
            node = OxmlElement("w:trHeight")
            insert_in_order(self._tr.get_or_add_trPr(), node)
        node.set(qn("w:hRule"), ROW_HEIGHT_RULE_XML[WD_ROW_HEIGHT_RULE(value)])


class Cell:
    def __init__(self, element, table: Table):
        self._tc = element
        self.table = table

    @property
    def part(self):
        return self.table._part

    @property
    def grid_span(self) -> int:
        node = self._tc.find(qn("w:tcPr") + "/" + qn("w:gridSpan"))
        if node is None:
            return 1
        try:
            return int(node.get(qn("w:val")))
        except (TypeError, ValueError):
            return 1

    def iter_inner_content(self):
        for child in self._tc:
            if child.tag == qn("w:p"):
                yield Paragraph(child, self.table._part, self.table._document)
            elif child.tag == qn("w:tbl"):
                yield Table(child, self.table._part, self.table._document)

    @property
    def paragraphs(self):
        return [
            Paragraph(node, self.table._part, self.table._document)
            for node in self._tc.findall(qn("w:p"))
        ]

    @property
    def text(self) -> str:
        return "\n".join(paragraph.text for paragraph in self.paragraphs)

    @text.setter
    def text(self, value: str):
        for child in list(self._tc):
            if child.tag != qn("w:tcPr"):
                self._tc.remove(child)
        paragraph = OxmlElement("w:p")
        self._tc.append(paragraph)
        if value:
            _copy_text(paragraph, value)

    @property
    def vertical_alignment(self):
        node = self._tc.find(qn("w:tcPr") + "/" + qn("w:vAlign"))
        value = node.get(qn("w:val")) if node is not None else None
        return {
            "top": WD_CELL_VERTICAL_ALIGNMENT.TOP,
            "center": WD_CELL_VERTICAL_ALIGNMENT.CENTER,
            "bottom": WD_CELL_VERTICAL_ALIGNMENT.BOTTOM,
        }.get(value)

    @vertical_alignment.setter
    def vertical_alignment(self, value):
        tc_pr = self._tc.get_or_add_tcPr()
        node = tc_pr.find(qn("w:vAlign"))
        if value is None:
            if node is not None:
                tc_pr.remove(node)
            return
        if node is None:
            node = OxmlElement("w:vAlign")
            insert_in_order(tc_pr, node)
        node.set(
            qn("w:val"),
            CELL_VERTICAL_ALIGNMENT_XML[WD_CELL_VERTICAL_ALIGNMENT(value)],
        )

    def add_paragraph(self, text: str = ""):
        node = OxmlElement("w:p")
        self._tc.append(node)
        paragraph = Paragraph(node, self.table._part, self.table._document)
        if text:
            paragraph.add_run(text)
        return paragraph

    @property
    def width(self):
        tc_w = self._tc.get_or_add_tcPr().find(qn("w:tcW"))
        return _cm_from_twips(tc_w.get(qn("w:w"))) if tc_w is not None else None

    @width.setter
    def width(self, value):
        tc_pr = self._tc.get_or_add_tcPr()
        tc_w = tc_pr.find(qn("w:tcW"))
        if tc_w is None:
            tc_w = OxmlElement("w:tcW")
            insert_in_order(tc_pr, tc_w)
        tc_w.set(qn("w:w"), _twips(value))
        tc_w.set(qn("w:type"), "dxa")

    def merge(self, other: "Cell"):
        current = _cell_coords(self.table, self._tc)
        target = _cell_coords(self.table, other._tc)
        if current is None or target is None:
            raise ValueError("待合并单元格不属于当前表格")
        r1, c1 = current
        r2, c2 = target
        r1, r2 = sorted((r1, r2))
        c1, c2 = sorted((c1, c2))
        span = c2 - c1 + 1
        row_count = r2 - r1 + 1
        trs = self.table._tbl.findall(qn("w:tr"))
        anchor_tc = None
        for row_index in range(r1, r2 + 1):
            tr = trs[row_index]
            grid_col = 0
            cells = []
            for tc in tr.findall(qn("w:tc")):
                cells.append((grid_col, tc))
                grid_col += tc.grid_span
            in_region = [(start, tc) for start, tc in cells if c1 <= start <= c2]
            if not in_region:
                continue
            anchor = in_region[0][1]
            if row_index == r1:
                anchor_tc = anchor
            for _, tc in in_region[1:]:
                tr.remove(tc)
            tc_pr = anchor.get_or_add_tcPr()
            grid_span = tc_pr.find(qn("w:gridSpan"))
            if span > 1:
                if grid_span is None:
                    grid_span = OxmlElement("w:gridSpan")
                    insert_in_order(tc_pr, grid_span)
                grid_span.set(qn("w:val"), str(span))
            elif grid_span is not None:
                tc_pr.remove(grid_span)
            vmerge = tc_pr.find(qn("w:vMerge"))
            if row_count > 1:
                if vmerge is None:
                    vmerge = OxmlElement("w:vMerge")
                    insert_in_order(tc_pr, vmerge)
                vmerge.set(qn("w:val"), "restart" if row_index == r1 else "continue")
            elif vmerge is not None:
                tc_pr.remove(vmerge)
        return Cell(anchor_tc, self.table)


def _run_text(run) -> str:
    chunks = []
    for child in run:
        if child.tag == qn("w:t"):
            chunks.append(child.text or "")
        elif child.tag == qn("w:tab"):
            chunks.append("\t")
        elif child.tag in (qn("w:br"), qn("w:cr")):
            chunks.append("\n")
        elif child.tag == qn("w:noBreakHyphen"):
            chunks.append("-")
    return "".join(chunks)


def _new_table_element(rows: int, cols: int):
    tbl = OxmlElement("w:tbl")
    tbl_pr = OxmlElement("w:tblPr")
    tbl_w = OxmlElement("w:tblW")
    tbl_w.set(qn("w:w"), "0")
    tbl_w.set(qn("w:type"), "auto")
    tbl_pr.append(tbl_w)
    tbl.append(tbl_pr)
    grid = OxmlElement("w:tblGrid")
    for _ in range(cols):
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(int(round(15.0 / max(cols, 1) * 566.929))))
        grid.append(col)
    tbl.append(grid)
    for row_index in range(rows):
        tr = OxmlElement("w:tr")
        if row_index == 0:
            tr_pr = OxmlElement("w:trPr")
            tr.append(tr_pr)
        for _ in range(cols):
            tc = OxmlElement("w:tc")
            tc_pr = OxmlElement("w:tcPr")
            tc.append(tc_pr)
            tc.append(OxmlElement("w:p"))
            tr.append(tc)
        tbl.append(tr)
    return tbl


def _drawing_xml(rid: str, width_emu: int, height_emu: int,
                 doc_pr_id: int, image: Image):
    drawing = OxmlElement("w:drawing")
    inline = OxmlElement("wp:inline")
    for key, value in {
        "distT": "0", "distB": "0", "distL": "0", "distR": "0",
    }.items():
        inline.set(key, value)
    extent = OxmlElement("wp:extent")
    extent.set("cx", str(width_emu))
    extent.set("cy", str(height_emu))
    inline.append(extent)
    effect = OxmlElement("wp:effectExtent")
    for key in ("l", "t", "r", "b"):
        effect.set(key, "0")
    inline.append(effect)
    doc_pr = OxmlElement("wp:docPr")
    doc_pr.set("id", str(doc_pr_id))
    doc_pr.set("name", f"Picture {doc_pr_id}")
    inline.append(doc_pr)
    frame_pr = OxmlElement("wp:cNvGraphicFramePr")
    locks = OxmlElement("a:graphicFrameLocks")
    locks.set("noChangeAspect", "1")
    frame_pr.append(locks)
    inline.append(frame_pr)
    graphic = OxmlElement("a:graphic")
    graphic_data = OxmlElement("a:graphicData")
    graphic_data.set("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture")
    pic = OxmlElement("pic:pic")
    nv_pic = OxmlElement("pic:nvPicPr")
    c_nv_pr = OxmlElement("pic:cNvPr")
    c_nv_pr.set("id", str(doc_pr_id))
    c_nv_pr.set("name", f"Picture {doc_pr_id}")
    nv_pic.append(c_nv_pr)
    nv_pic.append(OxmlElement("pic:cNvPicPr"))
    pic.append(nv_pic)
    blip_fill = OxmlElement("pic:blipFill")
    blip = OxmlElement("a:blip")
    blip.set(qn("r:embed"), rid)
    blip_fill.append(blip)
    stretch = OxmlElement("a:stretch")
    stretch.append(OxmlElement("a:fillRect"))
    blip_fill.append(stretch)
    pic.append(blip_fill)
    sp_pr = OxmlElement("pic:spPr")
    xfrm = OxmlElement("a:xfrm")
    off = OxmlElement("a:off")
    off.set("x", "0")
    off.set("y", "0")
    xfrm.append(off)
    ext = OxmlElement("a:ext")
    ext.set("cx", str(width_emu))
    ext.set("cy", str(height_emu))
    xfrm.append(ext)
    sp_pr.append(xfrm)
    geom = OxmlElement("a:prstGeom")
    geom.set("prst", "rect")
    geom.append(OxmlElement("a:avLst"))
    sp_pr.append(geom)
    pic.append(sp_pr)
    graphic_data.append(pic)
    graphic.append(graphic_data)
    inline.append(graphic)
    drawing.append(inline)
    return drawing


def _vmerge_state(tc) -> str | None:
    node = tc.find(qn("w:tcPr") + "/" + qn("w:vMerge"))
    if node is None:
        return None
    return node.get(qn("w:val"), "continue")


def _table_grid(table: Table):
    rows = table._tbl.findall(qn("w:tr"))
    col_count = len(table.columns)
    matrix = [[None for _ in range(col_count)] for _ in rows]
    for row_index, tr in enumerate(rows):
        col = 0
        for tc in tr.findall(qn("w:tc")):
            span = tc.grid_span
            for offset in range(span):
                if col + offset < col_count:
                    matrix[row_index][col + offset] = (tc, row_index, col)
            col += span
        for index in range(col_count):
            if matrix[row_index][index] is None:
                matrix[row_index][index] = (tr.findall(qn("w:tc"))[-1], row_index, index)
    for row_index in range(len(matrix)):
        for col in range(col_count):
            tc, _, _ = matrix[row_index][col]
            if _vmerge_state(tc) == "continue":
                source_row = row_index - 1
                while source_row >= 0:
                    candidate, _, _ = matrix[source_row][col]
                    if _vmerge_state(candidate) != "continue":
                        break
                    source_row -= 1
                if source_row >= 0:
                    matrix[row_index][col] = (
                        matrix[source_row][col][0], row_index, col,
                    )
    return matrix


def _cell_coords(table: Table, tc):
    for row_index, tr in enumerate(table._tbl.findall(qn("w:tr"))):
        col = 0
        for item in tr.findall(qn("w:tc")):
            if item is tc:
                return row_index, col
            col += item.grid_span
    return None

