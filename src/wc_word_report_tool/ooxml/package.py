"""Small OPC/OPC-relationships reader and writer for DOCX packages."""

from __future__ import annotations

import os
import posixpath
import tempfile
from collections import OrderedDict
from dataclasses import dataclass, field
from pathlib import Path
from zipfile import ZIP_DEFLATED, BadZipFile, ZipFile

from .xml import OxmlElement, parse_xml, serialize

CONTENT_TYPES = "/[Content_Types].xml"
PACKAGE_RELS = "/_rels/.rels"
OFFICE_DOCUMENT_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    "officeDocument"
)
IMAGE_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
)
HEADER_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
)
FOOTER_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
)
HYPERLINK_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
)
COMMENTS_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
)
CORE_PROPERTIES_REL = (
    "http://schemas.openxmlformats.org/package/2006/relationships/"
    "metadata/core-properties"
)


class PackageNotFoundError(Exception):
    pass


class Parts(OrderedDict):
    """Part mapping that iterates values, matching python-docx's package.parts."""

    def __iter__(self):
        return iter(self.values())


@dataclass
class Relationship:
    rId: str
    reltype: str
    target: str
    target_mode: str | None = None


@dataclass
class Relationships:
    partname: str
    items: OrderedDict[str, Relationship] = field(default_factory=OrderedDict)

    @classmethod
    def from_xml(cls, partname: str, data: bytes):
        root = parse_xml(data)
        rels = cls(partname)
        for item in root:
            rid = item.get("Id")
            if not rid:
                continue
            rels.items[rid] = Relationship(
                rid, item.get("Type", ""), item.get("Target", ""),
                item.get("TargetMode"),
            )
        return rels

    def to_xml(self) -> bytes:
        root = OxmlElement("pr:Relationships")
        for rel in self.items.values():
            node = OxmlElement("pr:Relationship")
            node.set("Id", rel.rId)
            node.set("Type", rel.reltype)
            node.set("Target", rel.target)
            if rel.target_mode:
                node.set("TargetMode", rel.target_mode)
            root.append(node)
        return serialize(root)

    def get(self, rid: str) -> Relationship | None:
        return self.items.get(rid)

    def next_rid(self) -> str:
        used = set(self.items)
        index = 1
        while f"rId{index}" in used:
            index += 1
        return f"rId{index}"


class Part:
    def __init__(self, package: "DocumentPackage", partname: str,
                 content_type: str, blob: bytes):
        self.package = package
        self.partname = "/" + partname.lstrip("/")
        self.content_type = content_type
        self._blob = blob
        self._element = None
        self._rels = None
        self._next_id = 1

    @property
    def blob(self) -> bytes:
        if self._element is not None:
            return serialize(self._element)
        return self._blob

    @blob.setter
    def blob(self, value):
        self._blob = value
        self._element = None

    @property
    def element(self):
        if self._element is None:
            self._element = parse_xml(self._blob)
        return self._element

    @element.setter
    def element(self, value):
        self._element = value
        self._blob = b""

    @property
    def rels(self) -> Relationships:
        if self._rels is None:
            self._rels = self.package.get_relationships(self)
        return self._rels

    @property
    def next_id(self) -> int:
        value = self._next_id
        self._next_id += 1
        return value

    def target_ref(self) -> str:
        return self.partname

    @property
    def numbering_part(self):
        return self.package.part_for_name("/word/numbering.xml")


class DocumentPackage:
    def __init__(self):
        self.parts: Parts = Parts()
        self.content_types: dict[str, str] = {}
        self.defaults: dict[str, str] = {}
        self._rels: dict[str, Relationships] = {}
        self._media_index = 1
        self._header_index = 1
        self._footer_index = 1

    @classmethod
    def from_path(cls, path: str | Path):
        return cls.from_bytes(Path(path).read_bytes())

    @classmethod
    def from_bytes(cls, data: bytes):
        package = cls()
        try:
            with ZipFile(__import__("io").BytesIO(data)) as archive:
                names = set(archive.namelist())
                if CONTENT_TYPES.lstrip("/") not in names:
                    raise PackageNotFoundError(
                        "Package not found: 不是有效的 DOCX，缺少 [Content_Types].xml"
                    )
                package._read_content_types(archive.read(CONTENT_TYPES.lstrip("/")))
                for name in archive.namelist():
                    if name.endswith("/") or name == CONTENT_TYPES.lstrip("/"):
                        continue
                    if name.startswith("_rels/") or "/_rels/" in name:
                        package._rels["/" + name] = Relationships.from_xml(
                            "/" + name, archive.read(name),
                        )
                        continue
                    partname = "/" + name
                    package.parts[partname] = Part(
                        package, partname, package._content_type(partname),
                        archive.read(name),
                    )
        except BadZipFile as exc:
            raise PackageNotFoundError(
                "Package not found: 不是有效的 DOCX，ZIP 结构损坏"
            ) from exc
        return package

    @classmethod
    def default(cls):
        return cls.from_path(Path(__file__).with_name("default.docx"))

    def _read_content_types(self, data: bytes):
        root = parse_xml(data)
        for child in root:
            tag = child.tag.rsplit("}", 1)[-1]
            if tag == "Default":
                self.defaults[child.get("Extension", "").lower()] = child.get("ContentType", "")
            elif tag == "Override":
                self.content_types["/" + child.get("PartName", "").lstrip("/")] = (
                    child.get("ContentType", "")
                )

    def _content_type(self, partname: str) -> str:
        if partname in self.content_types:
            return self.content_types[partname]
        ext = posixpath.splitext(partname)[1].lstrip(".").lower()
        return self.defaults.get(ext, "application/octet-stream")

    @property
    def main_document_part(self) -> Part:
        package_rels = self._rels.get("/_rels/.rels")
        if package_rels is None:
            raise PackageNotFoundError("DOCX 缺少包关系文件")
        for rel in package_rels.items.values():
            if rel.reltype == OFFICE_DOCUMENT_REL:
                if rel.target_mode == "External":
                    break
                target = "/" + rel.target.lstrip("/")
                if target in self.parts:
                    return self.parts[target]
        raise PackageNotFoundError("DOCX 缺少主文档关系")

    def _rels_name_for(self, part: Part) -> str:
        folder, name = posixpath.split(part.partname)
        return f"{folder}/_rels/{name}.rels"

    def get_relationships(self, part: Part) -> Relationships:
        name = self._rels_name_for(part)
        if name not in self._rels:
            self._rels[name] = Relationships(name)
        return self._rels[name]

    def resolve_target(self, source: Part, rel: Relationship) -> str | None:
        if rel.target_mode == "External":
            return None
        if rel.target.startswith("/"):
            return "/" + rel.target.lstrip("/")
        base = posixpath.dirname(source.partname)
        return "/" + posixpath.normpath(posixpath.join(base, rel.target)).lstrip("/")

    def add_relationship(self, source: Part, reltype: str, target: str,
                         *, target_mode: str | None = None) -> str:
        rels = source.rels
        for rel in rels.items.values():
            if rel.reltype == reltype and rel.target == target and rel.target_mode == target_mode:
                return rel.rId
        rid = rels.next_rid()
        rels.items[rid] = Relationship(rid, reltype, target, target_mode)
        return rid

    def add_package_relationship(self, reltype: str, target: str) -> str:
        rels = self._rels.setdefault("/_rels/.rels", Relationships("/_rels/.rels"))
        for rel in rels.items.values():
            if rel.reltype == reltype and rel.target == target:
                return rel.rId
        rid = rels.next_rid()
        rels.items[rid] = Relationship(rid, reltype, target)
        return rid

    def add_binary_part(self, folder: str, stem: str, extension: str,
                        blob: bytes, content_type: str) -> Part:
        while True:
            partname = f"/{folder}/{stem}{self._media_index}.{extension}"
            self._media_index += 1
            if partname not in self.parts:
                break
        part = Part(self, partname, content_type, blob)
        self.parts[partname] = part
        self.defaults[extension.lower()] = content_type
        return part

    def add_xml_part(self, folder: str, stem: str, root, content_type: str) -> Part:
        if stem == "header":
            index = self._header_index
            self._header_index += 1
            while f"/{folder}/{stem}{index}.xml" in self.parts:
                index = self._header_index
                self._header_index += 1
        elif stem == "footer":
            index = self._footer_index
            self._footer_index += 1
            while f"/{folder}/{stem}{index}.xml" in self.parts:
                index = self._footer_index
                self._footer_index += 1
        elif stem in ("core", "comments"):
            index = None
        else:
            index = 1
            while f"/{folder}/{stem}{index}.xml" in self.parts:
                index += 1
        partname = (
            f"/{folder}/{stem}.xml"
            if index is None
            else f"/{folder}/{stem}{index}.xml"
        )
        part = Part(self, partname, content_type, b"")
        part.element = root
        self.parts[partname] = part
        self.content_types[partname] = content_type
        return part

    def part_for_name(self, partname: str) -> Part | None:
        return self.parts.get("/" + partname.lstrip("/"))

    def _content_types_xml(self) -> bytes:
        root = OxmlElement("ct:Types")
        defaults = dict(self.defaults)
        defaults.setdefault("rels", "application/vnd.openxmlformats-package.relationships+xml")
        defaults.setdefault("xml", "application/xml")
        for ext, content_type in sorted(defaults.items()):
            node = OxmlElement("ct:Default")
            node.set("Extension", ext)
            node.set("ContentType", content_type)
            root.append(node)
        overrides = dict(self.content_types)
        overrides.setdefault(CONTENT_TYPES, "application/vnd.openxmlformats-package.content-types+xml")
        overrides.setdefault("/word/document.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")
        for partname, content_type in sorted(overrides.items()):
            node = OxmlElement("ct:Override")
            node.set("PartName", partname)
            node.set("ContentType", content_type)
            root.append(node)
        return serialize(root)

    def to_bytes(self) -> bytes:
        output = __import__("io").BytesIO()
        with ZipFile(output, "w", ZIP_DEFLATED) as archive:
            archive.writestr(CONTENT_TYPES.lstrip("/"), self._content_types_xml())
            for name, rels in self._rels.items():
                archive.writestr(name.lstrip("/"), rels.to_xml())
            for partname, part in self.parts.items():
                archive.writestr(partname.lstrip("/"), part.blob)
        return output.getvalue()

    def save(self, path: str | Path):
        target = Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        fd, temp_name = tempfile.mkstemp(
            prefix=f".{target.name}.", suffix=".tmp", dir=str(target.parent),
        )
        os.close(fd)
        try:
            Path(temp_name).write_bytes(self.to_bytes())
            os.replace(temp_name, target)
        except Exception:
            try:
                Path(temp_name).unlink()
            except FileNotFoundError:
                pass
            raise

