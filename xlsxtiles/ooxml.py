"""
xlsxtiles.ooxml — accès bas niveau au zip OOXML.

On lit l'archive directement plutôt que de passer par
`openpyxl.load_workbook` : le profilage doit rester peu coûteux et borné en
mémoire sur des classeurs de plusieurs centaines de Mo, ce qu'un chargement
complet ne permet pas. openpyxl n'intervient qu'au moment du rendu, sur une
feuille à la fois.
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

__all__ = ["NS_MAIN", "NS_REL_DOC", "NS_PKG_REL", "NS_XDR",
           "local", "read_xml", "rels_for"]

# Namespaces OOXML
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"


def local(tag: str) -> str:
    """'{ns}row' -> 'row'."""
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def read_xml(z: zipfile.ZipFile, name: str) -> ET.Element | None:
    """Racine d'une partie XML, ou None si absente ou illisible."""
    try:
        with z.open(name) as fh:
            return ET.parse(fh).getroot()
    except (KeyError, ET.ParseError):
        return None


def rels_for(z: zipfile.ZipFile, part: str) -> dict[str, str]:
    """Relations d'une partie -> {rId: target absolu dans le zip}."""
    p = Path(part)
    rels_path = f"{p.parent.as_posix()}/_rels/{p.name}.rels".lstrip("/")
    root = read_xml(z, rels_path)
    if root is None:
        return {}
    base = p.parent
    out: dict[str, str] = {}
    for rel in root:
        rid = rel.get("Id")
        target = rel.get("Target", "")
        if not rid or not target or rel.get("TargetMode") == "External":
            continue
        if target.startswith("/"):
            resolved = target.lstrip("/")
        else:
            resolved = (base / target).as_posix()
            while "/../" in resolved:
                resolved = re.sub(r"[^/]+/\.\./", "", resolved, count=1)
        out[rid] = resolved
    return out
