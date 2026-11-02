"""Direct OOXML backend used by wc_word_report_tool.

The exported surface mirrors the small python-docx-facing API used by the
formatter and its tests, while all reads and writes are performed directly on
OOXML parts with lxml and zipfile.
"""

from .document import (
    Cell,
    Comment,
    Comments,
    CoreProperties,
    Document,
    Font,
    HeaderFooter,
    InlineShape,
    InlineShapes,
    Paragraph,
    ParagraphFormat,
    Row,
    Run,
    Section,
    Settings,
    Style,
    Styles,
    Table,
)
from .enums import (
    WD_ALIGN_PARAGRAPH,
    WD_CELL_VERTICAL_ALIGNMENT,
    WD_COLOR_INDEX,
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
from .image import Image
from .namespaces import NSMAP, qn
from .package import DocumentPackage, PackageNotFoundError, Part, Relationship
from .units import Cm, Inches, Length, Mm, Pt, RGBColor
from .xml import OxmlElement, parse_xml, serialize_for_reading

__all__ = [
    "Cell",
    "Cm",
    "Comment",
    "Comments",
    "CoreProperties",
    "Document",
    "DocumentPackage",
    "Font",
    "HeaderFooter",
    "Image",
    "Inches",
    "InlineShape",
    "InlineShapes",
    "Length",
    "Mm",
    "NSMAP",
    "OxmlElement",
    "PackageNotFoundError",
    "Paragraph",
    "ParagraphFormat",
    "Part",
    "Pt",
    "RGBColor",
    "Relationship",
    "Row",
    "Run",
    "Section",
    "Settings",
    "Style",
    "Styles",
    "Table",
    "WD_COLOR_INDEX",
    "WD_ALIGN_PARAGRAPH",
    "WD_CELL_VERTICAL_ALIGNMENT",
    "WD_ORIENT",
    "WD_PARAGRAPH_ALIGNMENT",
    "WD_ROW_HEIGHT_RULE",
    "WD_SECTION_START",
    "WD_STYLE_TYPE",
    "WD_TAB_ALIGNMENT",
    "WD_TABLE_ALIGNMENT",
    "WD_TABLE_DIRECTION",
    "WD_UNDERLINE",
    "parse_xml",
    "qn",
    "serialize_for_reading",
]

