from __future__ import annotations

import io
import logging
import posixpath
import zipfile
from dataclasses import dataclass
from typing import Iterable
from typing import Optional

from lxml import etree

log = logging.getLogger(__name__)


@dataclass(frozen=True)
class ZipReadLimits:
    max_file_bytes: int = 200 * 1024 * 1024
    max_total_bytes: int = 600 * 1024 * 1024


class DocxPackage:
    def __init__(self, path: str, *, limits: Optional[ZipReadLimits] = None):
        self.path = path
        self._zf = zipfile.ZipFile(path, "r")
        self.limits = limits or ZipReadLimits()

    def close(self) -> None:
        self._zf.close()

    def namelist(self) -> list[str]:
        return self._zf.namelist()

    def exists(self, name: str) -> bool:
        try:
            self._zf.getinfo(name)
            return True
        except KeyError:
            return False

    def read_bytes(self, name: str, *, max_bytes: Optional[int] = None) -> bytes:
        info = self._zf.getinfo(name)
        size = info.file_size
        per_file_limit = max_bytes if max_bytes is not None else self.limits.max_file_bytes
        if size > per_file_limit:
            raise ValueError(f"zip member too large: {name} ({size} bytes > {per_file_limit})")
        with self._zf.open(info, "r") as f:
            data = f.read(per_file_limit + 1)
        if len(data) > per_file_limit:
            raise ValueError(f"zip member too large while reading: {name}")
        return data

    def read_xml(self, name: str) -> etree._Element:
        data = self.read_bytes(name, max_bytes=20 * 1024 * 1024)
        parser = etree.XMLParser(recover=True, huge_tree=False)
        return etree.fromstring(data, parser=parser)

    def iter_members_under(self, prefix: str) -> Iterable[str]:
        prefix = prefix.rstrip("/") + "/"
        for n in self._zf.namelist():
            if n.startswith(prefix) and not n.endswith("/"):
                yield n

    @staticmethod
    def normalize_target(base_part: str, target: str) -> str:
        # Relationship targets are POSIX paths.
        base_dir = posixpath.dirname(base_part)
        joined = posixpath.normpath(posixpath.join(base_dir, target))
        return joined.lstrip("/")
 

def parse_rels(rels_xml: etree._Element) -> dict[str, dict[str, str]]:
     """
     Return map: rId -> {Type, Target, TargetMode?}
     """
     ns = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
     out: dict[str, dict[str, str]] = {}
     for rel in rels_xml.findall(".//r:Relationship", namespaces=ns):
         rid = rel.get("Id")
         if not rid:
             continue
         out[rid] = {
             "Type": rel.get("Type", ""),
             "Target": rel.get("Target", ""),
             "TargetMode": rel.get("TargetMode", ""),
         }
     return out


def sniff_extension(data: bytes) -> Optional[str]:
     if data.startswith(b"PK\x03\x04"):
         # Could be docx/xlsx; the caller may decide based on inner structure if needed.
         return "zip"
     if data.startswith(bytes.fromhex("D0CF11E0A1B11AE1")):
         return "ole"
     return None
