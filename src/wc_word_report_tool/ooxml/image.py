"""Image metadata and media-type helpers without Pillow or python-docx."""

from __future__ import annotations

import io
import struct
from pathlib import Path


class Image:
    """Minimal image probe exposing pixel size, DPI and content type."""

    def __init__(self, *, px_width: int, px_height: int, horz_dpi: float,
                 vert_dpi: float, content_type: str, extension: str):
        self.px_width = px_width
        self.px_height = px_height
        self.horz_dpi = horz_dpi
        self.vert_dpi = vert_dpi
        self.content_type = content_type
        self.ext = extension

    @classmethod
    def from_file(cls, stream):
        if isinstance(stream, (str, Path)):
            with open(stream, "rb") as handle:
                data = handle.read()
        elif isinstance(stream, (bytes, bytearray)):
            data = bytes(stream)
        else:
            position = stream.tell()
            data = stream.read()
            stream.seek(position)
        return cls.from_blob(data)

    @classmethod
    def from_blob(cls, data: bytes):
        if data.startswith(b"\x89PNG\r\n\x1a\n"):
            return _png(data)
        if data.startswith(b"\xff\xd8"):
            return _jpeg(data)
        if data[:6] in (b"GIF87a", b"GIF89a"):
            return _gif(data)
        if data.startswith(b"BM"):
            return _bmp(data)
        if data[:4] in (b"II*\x00", b"MM\x00*"):
            return _tiff(data)
        raise ValueError("无法识别图片格式，仅支持 PNG/JPEG/GIF/BMP/TIFF")


def _png(data: bytes) -> Image:
    if len(data) < 24:
        raise ValueError("PNG 文件不完整")
    width, height = struct.unpack(">II", data[16:24])
    x_dpi = y_dpi = 96.0
    offset = 8
    while offset + 12 <= len(data):
        size = struct.unpack(">I", data[offset:offset + 4])[0]
        kind = data[offset + 4:offset + 8]
        payload = data[offset + 8:offset + 8 + size]
        if kind == b"pHYs" and len(payload) >= 9 and payload[8] == 1:
            ppm_x, ppm_y = struct.unpack(">II", payload[:8])
            x_dpi = ppm_x * 0.0254 if ppm_x else 96.0
            y_dpi = ppm_y * 0.0254 if ppm_y else 96.0
        offset += 12 + size
        if kind == b"IEND":
            break
    return Image(
        px_width=width, px_height=height, horz_dpi=x_dpi, vert_dpi=y_dpi,
        content_type="image/png", extension="png",
    )


def _jpeg(data: bytes) -> Image:
    offset = 2
    x_dpi = y_dpi = 96.0
    width = height = 0
    while offset + 4 <= len(data):
        if data[offset] != 0xFF:
            offset += 1
            continue
        marker = data[offset + 1]
        offset += 2
        if marker in (0xD8, 0xD9):
            continue
        if offset + 2 > len(data):
            break
        length = struct.unpack(">H", data[offset:offset + 2])[0]
        payload = data[offset + 2:offset + length]
        if marker == 0xE0 and payload.startswith(b"JFIF\x00") and len(payload) >= 12:
            units = payload[7]
            xd = struct.unpack(">H", payload[8:10])[0]
            yd = struct.unpack(">H", payload[10:12])[0]
            if units == 1:
                x_dpi, y_dpi = float(xd or 96), float(yd or 96)
            elif units == 2:
                x_dpi, y_dpi = float(xd or 96) * 2.54, float(yd or 96) * 2.54
        if marker in {
            0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7,
            0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF,
        } and len(payload) >= 5:
            height, width = struct.unpack(">HH", payload[1:5])
            break
        offset += length
    if not width or not height:
        raise ValueError("JPEG 文件缺少尺寸信息")
    return Image(
        px_width=width, px_height=height, horz_dpi=x_dpi, vert_dpi=y_dpi,
        content_type="image/jpeg", extension="jpg",
    )


def _gif(data: bytes) -> Image:
    if len(data) < 10:
        raise ValueError("GIF 文件不完整")
    width, height = struct.unpack("<HH", data[6:10])
    return Image(
        px_width=width, px_height=height, horz_dpi=96, vert_dpi=96,
        content_type="image/gif", extension="gif",
    )


def _bmp(data: bytes) -> Image:
    if len(data) < 26:
        raise ValueError("BMP 文件不完整")
    width, height = struct.unpack("<ii", data[18:26])
    return Image(
        px_width=abs(width), px_height=abs(height), horz_dpi=96, vert_dpi=96,
        content_type="image/bmp", extension="bmp",
    )


def _tiff(data: bytes) -> Image:
    endian = "<" if data[:2] == b"II" else ">"
    if len(data) < 8:
        raise ValueError("TIFF 文件不完整")
    ifd_offset = struct.unpack(endian + "I", data[4:8])[0]
    if ifd_offset + 2 > len(data):
        raise ValueError("TIFF IFD 越界")
    count = struct.unpack(endian + "H", data[ifd_offset:ifd_offset + 2])[0]
    values = {}
    pos = ifd_offset + 2
    for _ in range(count):
        if pos + 12 > len(data):
            break
        tag, kind, amount = struct.unpack(endian + "HHI", data[pos:pos + 8])
        raw = data[pos + 8:pos + 12]
        if kind == 3 and amount == 1:
            values[tag] = struct.unpack(endian + "H", raw[:2])[0]
        elif kind == 4 and amount == 1:
            values[tag] = struct.unpack(endian + "I", raw)[0]
        pos += 12
    return Image(
        px_width=int(values.get(256, 0)),
        px_height=int(values.get(257, 0)),
        horz_dpi=96, vert_dpi=96,
        content_type="image/tiff", extension="tiff",
    )


def image_stream(image):
    if isinstance(image, (bytes, bytearray)):
        return io.BytesIO(bytes(image)), True, bytes(image)
    if hasattr(image, "read"):
        position = image.tell()
        data = image.read()
        image.seek(position)
        return image, False, data
    path = Path(image)
    if not path.is_file():
        raise FileNotFoundError(f"图片不存在: {path}")
    return str(path), False, path.read_bytes()

