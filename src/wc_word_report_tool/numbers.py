from __future__ import annotations

from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Union

NumberLike = Union[int, float, str, Decimal]

_DIGITS = "零壹贰叁肆伍陆柒捌玖"
_SMALL_UNITS = ("仟", "佰", "拾", "")
_SECTION_UNITS = (
    "",
    "万",
    "亿",
    "兆",
    "京",
    "垓",
    "秭",
    "穰",
    "沟",
    "涧",
    "正",
    "载",
)


def _to_decimal(value: NumberLike) -> Decimal:
    try:
        return Decimal(str(value).replace(",", "").strip())
    except (InvalidOperation, AttributeError):
        raise ValueError(f"无法转换为数字: {value!r}") from None


def _convert_four_digits(section: str) -> str:
    result: list[str] = []
    zero_pending = False

    for index, char in enumerate(section.zfill(4)):
        digit = int(char)
        unit = _SMALL_UNITS[index]
        if digit == 0:
            zero_pending = bool(result)
            continue
        if zero_pending:
            result.append(_DIGITS[0])
            zero_pending = False
        result.append(_DIGITS[digit])
        result.append(unit)

    return "".join(result)


def _convert_integer(integer_part: str) -> str:
    if integer_part == "0":
        return _DIGITS[0]

    groups: list[str] = []
    while integer_part:
        groups.insert(0, integer_part[-4:])
        integer_part = integer_part[:-4]

    result: list[str] = []
    zero_between_groups = False

    for index, raw_group in enumerate(groups):
        group_value = int(raw_group)
        converted = _convert_four_digits(raw_group)
        section_unit = _SECTION_UNITS[len(groups) - index - 1]

        if converted:
            if zero_between_groups or (result and group_value < 1000):
                result.append(_DIGITS[0])
                zero_between_groups = False
            result.append(converted)
            result.append(section_unit)
        elif result:
            zero_between_groups = True

    text = "".join(result).rstrip(_DIGITS[0])
    return text or _DIGITS[0]


def number_to_chinese_upper(value: NumberLike) -> str:
    amount = _to_decimal(value).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    negative = amount < 0
    amount = abs(amount)

    integer_part = int(amount)
    decimal_part = int((amount - Decimal(integer_part)) * 100)
    jiao = decimal_part // 10
    fen = decimal_part % 10

    integer_text = _convert_integer(str(integer_part))
    result = f"{integer_text}元"

    if jiao == 0 and fen == 0:
        result += "整"
    else:
        if jiao > 0:
            result += f"{_DIGITS[jiao]}角"
        elif integer_part > 0 and fen > 0:
            result += "零"
        if fen > 0:
            result += f"{_DIGITS[fen]}分"

    if negative:
        result = f"负{result}"
    return result


class NumberConverter:
    @staticmethod
    def number_to_chinese(num: NumberLike) -> str:
        return number_to_chinese_upper(num)


about_numbers = NumberConverter
