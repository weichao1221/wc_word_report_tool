"""The subset of python-docx-facing enums used by the public API."""

from __future__ import annotations

from enum import Enum, IntEnum


class WD_PARAGRAPH_ALIGNMENT(IntEnum):
    LEFT = 0
    CENTER = 1
    RIGHT = 2
    JUSTIFY = 3
    DISTRIBUTE = 4


class WD_TABLE_ALIGNMENT(IntEnum):
    LEFT = 0
    CENTER = 1
    RIGHT = 2


WD_ALIGN_PARAGRAPH = WD_PARAGRAPH_ALIGNMENT


class WD_CELL_VERTICAL_ALIGNMENT(IntEnum):
    TOP = 0
    CENTER = 1
    BOTTOM = 2


CELL_VERTICAL_ALIGNMENT_XML = {
    WD_CELL_VERTICAL_ALIGNMENT.TOP: "top",
    WD_CELL_VERTICAL_ALIGNMENT.CENTER: "center",
    WD_CELL_VERTICAL_ALIGNMENT.BOTTOM: "bottom",
}


class WD_ROW_HEIGHT_RULE(IntEnum):
    AUTO = 0
    AT_LEAST = 1
    EXACTLY = 2


class WD_ORIENT(IntEnum):
    PORTRAIT = 0
    LANDSCAPE = 1


class WD_TABLE_DIRECTION(IntEnum):
    LTR = 0
    RTL = 1


class WD_SECTION_START(IntEnum):
    CONTINUOUS = 0
    NEW_COLUMN = 1
    NEW_PAGE = 2
    EVEN_PAGE = 3
    ODD_PAGE = 4


class WD_STYLE_TYPE(IntEnum):
    PARAGRAPH = 1
    CHARACTER = 2
    TABLE = 3
    LIST = 4


class WD_TAB_ALIGNMENT(IntEnum):
    LEFT = 0
    CENTER = 1
    RIGHT = 2
    DECIMAL = 3
    BAR = 4
    CLEAR = 101
    END = 102
    NUM = 103
    START = 104


class WD_UNDERLINE(str, Enum):
    SINGLE = "single"
    DOUBLE = "double"
    THICK = "thick"
    DOTTED = "dotted"
    DASH = "dash"
    DOT_DASH = "dotDash"
    DOT_DOT_DASH = "dotDotDash"
    WAVE = "wave"
    DOTTED_HEAVY = "dottedHeavy"
    DASHED_HEAVY = "dashedHeavy"
    DASH_LONG = "dashLong"
    DASH_LONG_HEAVY = "dashLongHeavy"
    WAVY_HEAVY = "wavyHeavy"
    WAVY_DOUBLE = "wavyDouble"
    WORDS = "words"
    NONE = "none"


class WD_COLOR_INDEX(str, Enum):
    BLACK = "black"
    BLUE = "blue"
    CYAN = "cyan"
    GREEN = "green"
    MAGENTA = "magenta"
    RED = "red"
    YELLOW = "yellow"
    WHITE = "white"


PARAGRAPH_ALIGNMENT_XML = {
    WD_PARAGRAPH_ALIGNMENT.LEFT: "left",
    WD_PARAGRAPH_ALIGNMENT.CENTER: "center",
    WD_PARAGRAPH_ALIGNMENT.RIGHT: "right",
    WD_PARAGRAPH_ALIGNMENT.JUSTIFY: "both",
    WD_PARAGRAPH_ALIGNMENT.DISTRIBUTE: "distribute",
}

TABLE_ALIGNMENT_XML = {
    WD_TABLE_ALIGNMENT.LEFT: "left",
    WD_TABLE_ALIGNMENT.CENTER: "center",
    WD_TABLE_ALIGNMENT.RIGHT: "right",
}

SECTION_START_XML = {
    WD_SECTION_START.CONTINUOUS: "continuous",
    WD_SECTION_START.NEW_COLUMN: "nextColumn",
    WD_SECTION_START.NEW_PAGE: "nextPage",
    WD_SECTION_START.EVEN_PAGE: "evenPage",
    WD_SECTION_START.ODD_PAGE: "oddPage",
}

SECTION_START_REVERSE = {
    value: key for key, value in SECTION_START_XML.items()
}

ROW_HEIGHT_RULE_XML = {
    WD_ROW_HEIGHT_RULE.AUTO: "auto",
    WD_ROW_HEIGHT_RULE.AT_LEAST: "atLeast",
    WD_ROW_HEIGHT_RULE.EXACTLY: "exact",
}

TAB_ALIGNMENT_XML = {
    WD_TAB_ALIGNMENT.LEFT: "left",
    WD_TAB_ALIGNMENT.CENTER: "center",
    WD_TAB_ALIGNMENT.RIGHT: "right",
    WD_TAB_ALIGNMENT.DECIMAL: "decimal",
    WD_TAB_ALIGNMENT.BAR: "bar",
    WD_TAB_ALIGNMENT.CLEAR: "clear",
    WD_TAB_ALIGNMENT.END: "end",
    WD_TAB_ALIGNMENT.NUM: "num",
    WD_TAB_ALIGNMENT.START: "start",
}

