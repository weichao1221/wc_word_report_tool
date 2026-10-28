from .numbers import NumberConverter, about_numbers, number_to_chinese_upper
from .word import (
    DEFAULT_BODY_SIZE,
    DEFAULT_CN_FONT,
    DEFAULT_EN_FONT,
    DEFAULT_HEADER_FOOTER_FONT,
    DEFAULT_HEADER_FOOTER_SIZE,
    FONT_SIZE_MAP,
    ORIENTATION_MAP,
    PAGE_NUMBER_FORMATS,
    PT_TO_FONT_SIZE_MAP,
    WordFormatter,
    about_word,
    describe_font_size,
    pt_to_font_size,
)

__version__ = "0.7.6"

__all__ = [
    "__version__",
    # 数字转大写
    "NumberConverter",
    "about_numbers",
    "number_to_chinese_upper",
    # Word 格式化
    "WordFormatter",
    "about_word",
    # 字号换算
    "pt_to_font_size",
    "describe_font_size",
    # 常量
    "FONT_SIZE_MAP",
    "PT_TO_FONT_SIZE_MAP",
    "PAGE_NUMBER_FORMATS",
    "ORIENTATION_MAP",
    "DEFAULT_CN_FONT",
    "DEFAULT_EN_FONT",
    "DEFAULT_BODY_SIZE",
    "DEFAULT_HEADER_FOOTER_FONT",
    "DEFAULT_HEADER_FOOTER_SIZE",
]
