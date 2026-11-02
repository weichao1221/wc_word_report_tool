"""lxml helpers that keep namespace prefixes and unknown XML intact."""

from __future__ import annotations

from copy import deepcopy
from typing import Iterable

from lxml import etree

from .namespaces import NSMAP, local_name, qn

_PARSER = etree.XMLParser(
    resolve_entities=False,
    no_network=True,
    recover=False,
    remove_blank_text=False,
    huge_tree=False,
)


def _child_sequence(parent_tag: str) -> list[str]:
    """Return the OOXML schema sequence for ``parent_tag``'s children."""
    local = local_name(parent_tag)
    orders = {
        "pPr": [
            "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
            "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd",
            "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
            "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
            "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
            "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
            "textDirection", "textAlignment", "textboxTightWrap",
            "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
        ],
        "rPr": [
            "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
            "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
            "noProof", "snapToGrid", "vanish", "webHidden", "color",
            "spacing", "w", "kern", "position", "sz", "szCs", "highlight",
            "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl",
            "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath",
            "rPrChange",
        ],
        "tblPr": [
            "tblStyle", "tblpPr", "tblOverlap", "bidiVisual",
            "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
            "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout",
            "tblCellMar", "tblLook", "tblCaption", "tblDescription",
        ],
        "tcPr": [
            "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders",
            "shd", "noWrap", "tcMar", "textDirection", "tcFitText",
            "vAlign", "hideMark",
        ],
        "trPr": [
            "cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore",
            "wAfter", "cantSplit", "trHeight", "tblHeader", "tblCellSpacing",
            "jc", "hidden",
        ],
        "sectPr": [
            "headerReference", "footerReference", "footnotePr", "endnotePr",
            "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType",
            "pgNumType", "cols", "formProt", "vAlign", "noEndnote",
            "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid",
            "printerSettings", "sectPrChange",
        ],
        "style": [
            "name", "aliases", "basedOn", "next", "link", "autoRedefine",
            "hidden", "uiPriority", "semiHidden", "unhideWhenUsed",
            "qFormat", "locked", "personal", "personalCompose",
            "personalReply", "rsid", "pPr", "rPr", "tblPr", "trPr",
            "tcPr", "tblStylePr",
        ],
    }
    return orders.get(local, [])


def insert_in_order(parent, child) -> None:
    """Insert an OOXML child according to the common schema sequence.

    The insertion point is the first existing sibling that follows ``child`` in
    the schema, so properties never land after an element they must precede.
    Unknown containers or unknown children fall back to appending.
    """
    order = _child_sequence(parent.tag)
    local = local_name(child.tag)
    # Schema children all live in the w: namespace; anything else (or an
    # unknown local name) is appended rather than guessed into the sequence.
    if local not in order or child.tag != qn("w:" + local):
        parent.append(child)
        return
    later = {qn("w:" + tag) for tag in order[order.index(local) + 1 :]}
    for index, existing in enumerate(parent):
        if existing.tag in later:
            parent.insert(index, child)
            return
    parent.append(child)


class OxmlElementBase(etree.ElementBase):
    """Element class providing python-docx-compatible convenience properties."""

    @property
    def xml(self) -> str:
        return etree.tostring(
            self, encoding="unicode", with_tail=False,
        )

    def xpath(self, path, namespaces=None, **kwargs):
        merged = dict(NSMAP)
        if namespaces:
            merged.update(namespaces)
        return super().xpath(path, namespaces=merged, **kwargs)

    def _get_or_add(self, tag: str):
        child = self.find(qn(tag))
        if child is None:
            child = OxmlElement(tag)
            if local_name(self.tag) in {"p", "r", "tbl", "tr", "tc"} and local_name(qn(tag)) in {
                "pPr", "rPr", "tblPr", "trPr", "tcPr",
            }:
                self.insert(0, child)
            else:
                insert_in_order(self, child)
        return child

    def get_or_add_pPr(self):
        return self._get_or_add("w:pPr")

    def get_or_add_rPr(self):
        return self._get_or_add("w:rPr")

    def get_or_add_tblPr(self):
        return self._get_or_add("w:tblPr")

    def get_or_add_trPr(self):
        return self._get_or_add("w:trPr")

    def get_or_add_tcPr(self):
        return self._get_or_add("w:tcPr")

    def get_or_add_ind(self):
        p_pr = self if local_name(self.tag) == "pPr" else self.get_or_add_pPr()
        ind = p_pr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind")
            insert_in_order(p_pr, ind)
        return ind

    @property
    def pPr(self):
        return self.find(qn("w:pPr"))

    @property
    def rPr(self):
        return self.find(qn("w:rPr"))

    @property
    def tblPr(self):
        return self.find(qn("w:tblPr"))

    @property
    def rFonts(self):
        return self.find(qn("w:rFonts"))

    @property
    def ind(self):
        return self.find(qn("w:ind"))

    @property
    def tr_lst(self):
        return self.findall(qn("w:tr"))

    @property
    def tc_lst(self):
        return self.findall(qn("w:tc"))

    @property
    def grid_span(self) -> int:
        span = self.find(".//" + qn("w:gridSpan"))
        if span is None:
            return 1
        try:
            return max(1, int(span.get(qn("w:val"), "1")))
        except ValueError:
            return 1

    def addnext(self, element):
        super().addnext(element)


_LOOKUP = etree.ElementDefaultClassLookup(element=OxmlElementBase)
_PARSER.set_element_class_lookup(_LOOKUP)


def OxmlElement(tag: str):
    if ":" in tag and not tag.startswith("{"):
        prefix = tag.split(":", 1)[0]
        nsmap = {prefix: NSMAP[prefix]}
    else:
        nsmap = None
    return _PARSER.makeelement(qn(tag), nsmap=nsmap)


def parse_xml(xml: str | bytes):
    if isinstance(xml, str):
        xml = xml.encode("utf-8")
    return etree.fromstring(xml, parser=_PARSER)


def clone(element):
    return deepcopy(element)


def serialize(element, *, declaration: bool = True) -> bytes:
    return etree.tostring(
        element,
        xml_declaration=declaration,
        encoding="UTF-8",
        standalone=True,
    )


def remove_children(element, keep: Iterable[int] | None = None) -> None:
    keep_set = set(keep or [])
    for index in range(len(element) - 1, -1, -1):
        if index not in keep_set:
            del element[index]


serialize_for_reading = lambda element: etree.tostring(  # noqa: E731
    element, encoding="unicode", with_tail=False,
)

