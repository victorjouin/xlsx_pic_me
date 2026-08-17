"""
xlsxtiles.fonts — polices demandées par le classeur vs polices disponibles.

L'enjeu n'est pas esthétique : une police absente est substituée par
LibreOffice, les largeurs de texte changent, la grille pixel prédite dérive et
le surlignage de cellule devient faux. D'où le verdict `grid_reliable`.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import zipfile
from pathlib import Path
from typing import Iterable

from .ooxml import NS_MAIN, read_xml

__all__ = ["fonts_used", "installed_fonts", "check_fonts", "FONT_SUBSTITUTES"]

# Polices dont l'absence dans le conteneur fait dériver la grille pixel.
# Les substituts métriquement compatibles sont acceptés.
FONT_SUBSTITUTES = {
    "calibri": {"calibri", "carlito"},
    "cambria": {"cambria", "caladea"},
    "arial": {"arial", "liberation sans", "arimo"},
    "times new roman": {"times new roman", "liberation serif", "tinos"},
    "courier new": {"courier new", "liberation mono", "cousine"},
    "helvetica": {"helvetica", "liberation sans", "arimo"},
}


def fonts_used(z: zipfile.ZipFile) -> list[str]:
    """Familles de polices déclarées dans styles.xml."""
    root = read_xml(z, "xl/styles.xml")
    if root is None:
        return []
    names: set[str] = set()
    for font in root.iter(f"{{{NS_MAIN}}}font"):
        nm = font.find(f"{{{NS_MAIN}}}name")
        if nm is not None and nm.get("val"):
            names.add(nm.get("val"))
    return sorted(names)


def installed_fonts() -> set[str] | None:
    """Polices disponibles sur la machine. None si l'inventaire est impossible.

    fontconfig en conteneur Linux (la cible de prod), énumération du répertoire
    de polices sous Windows (les postes de dev). Sans la branche Windows, le
    verdict `grid_reliable` serait faussement False sur toute machine de dev et
    le signal deviendrait du bruit qu'on apprend à ignorer.
    """
    if os.name == "nt":
        fams: set[str] = set()
        roots = [Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"]
        local = os.environ.get("LOCALAPPDATA")
        if local:
            roots.append(Path(local) / "Microsoft" / "Windows" / "Fonts")
        for root in roots:
            try:
                for f in root.iterdir():
                    if f.suffix.lower() in (".ttf", ".otf", ".ttc"):
                        fams.add(f.stem.lower())
            except OSError:
                continue
        return fams or None

    if not shutil.which("fc-list"):
        return None
    try:
        out = subprocess.run(["fc-list", ":", "family"], capture_output=True,
                             text=True, timeout=10).stdout
    except (subprocess.SubprocessError, OSError):
        return None
    fams = set()
    for line in out.splitlines():
        for fam in line.split(","):
            fam = fam.strip().lower()
            if fam:
                fams.add(fam)
    return fams


def check_fonts(used: Iterable[str]) -> tuple[list[str], list[str]]:
    """-> (polices manquantes, avertissements)"""
    installed = installed_fonts()
    if installed is None:
        return [], ["inventaire des polices indisponible : présence non "
                    "vérifiée, grille pixel à considérer comme non fiable"]
    missing = []
    for f in used:
        key = f.lower()
        candidates = FONT_SUBSTITUTES.get(key, {key})
        if not any(any(c in fam for fam in installed) for c in candidates):
            missing.append(f)
    return missing, []
