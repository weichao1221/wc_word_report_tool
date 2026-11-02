"""Small replacements for the python-docx unit and colour helpers."""

from __future__ import annotations

from dataclasses import dataclass

EMU_PER_CM = 360_000
EMU_PER_PT = 12_700
EMU_PER_INCH = 914_400
TWIPS_PER_CM = 566.9291338582677
TWIPS_PER_PT = 20


class Length(int):
    """Integer EMU value with the conversion properties used by the project."""

    @property
    def emu(self) -> int:
        return int(self)

    @property
    def cm(self) -> float:
        return int(self) / EMU_PER_CM

    @property
    def pt(self) -> float:
        return int(self) / EMU_PER_PT

    @property
    def inches(self) -> float:
        return int(self) / EMU_PER_INCH


def Cm(value: float) -> Length:
    return Length(round(float(value) * EMU_PER_CM))


def Pt(value: float) -> Length:
    return Length(round(float(value) * EMU_PER_PT))


def Mm(value: float) -> Length:
    return Length(round(float(value) / 10 * EMU_PER_CM))


def Inches(value: float) -> Length:
    return Length(round(float(value) * EMU_PER_INCH))


@dataclass(frozen=True)
class RGBColor:
    r: int
    g: int
    b: int

    def __post_init__(self):
        for value in (self.r, self.g, self.b):
            if not 0 <= int(value) <= 255:
                raise ValueError(f"RGB 分量必须在 0~255，收到 {value!r}")

    def __iter__(self):
        yield self.r
        yield self.g
        yield self.b

    def __str__(self):
        return f"{self.r:02X}{self.g:02X}{self.b:02X}"

