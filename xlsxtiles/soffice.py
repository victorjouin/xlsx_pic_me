"""
xlsxtiles.soffice — pilotage de LibreOffice en headless.

Localisation du binaire, profil utilisateur jetable à locale figée, et les deux
conversions dont on a besoin : recalcul des formules et export PDF.

Tout ce qui touche au processus externe est ici ; le reste du package ne sait
pas que LibreOffice existe.
"""

from __future__ import annotations

import functools
import os
import shutil
import subprocess
import time
from pathlib import Path
from typing import Sequence

from .config import RENDER_LOCALE, SOFFICE_TIMEOUT

__all__ = ["find_soffice", "seed_profile", "to_pdf", "recalc_with_libreoffice"]

# Emplacements d'installation par défaut. LibreOffice ne s'ajoute PAS au PATH
# sous Windows, et le binaire n'est pas au même endroit selon l'OS.
_SOFFICE_CANDIDATES = {
    "nt": [
        r"C:\Program Files\LibreOffice\program\soffice.com",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.com",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ],
    "posix": [
        "/usr/bin/soffice", "/usr/lib/libreoffice/program/soffice",
        "/opt/libreoffice/program/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/opt/homebrew/bin/soffice",
    ],
}

_REGISTRY_XCU = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry"
           xmlns:xs="http://www.w3.org/2001/XMLSchema"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
 <item oor:path="/org.openoffice.Setup/L10N">
  <prop oor:name="ooLocale" oor:op="fuse"><value>{loc}</value></prop>
 </item>
 <item oor:path="/org.openoffice.Setup/L10N">
  <prop oor:name="ooSetupSystemLocale" oor:op="fuse"><value>{loc}</value></prop>
 </item>
</oor:items>
"""


@functools.lru_cache(maxsize=1)
def find_soffice() -> str:
    """Localise le binaire LibreOffice. Surchargeable par SOFFICE_PATH.

    Sous Windows on privilégie soffice.com sur soffice.exe : le .exe se détache
    du process appelant et subprocess.run rend la main avant que le PDF existe.
    """
    env = os.environ.get("SOFFICE_PATH")
    if env and Path(env).exists():
        return env
    names = (["soffice.com", "soffice.exe", "soffice"] if os.name == "nt"
             else ["soffice", "libreoffice"])
    for name in names:
        found = shutil.which(name)
        if found:
            return found
    for cand in _SOFFICE_CANDIDATES.get(os.name, []):
        if Path(cand).exists():
            return cand
    raise RuntimeError(
        "LibreOffice introuvable. Installe-le, ou renseigne la variable "
        "d'environnement SOFFICE_PATH avec le chemin complet du binaire "
        r'(ex. Windows : C:\Program Files\LibreOffice\program\soffice.com).')


def seed_profile(profile_dir: Path, locale: str = RENDER_LOCALE) -> Path:
    """Prépare un profil LibreOffice neuf avec une locale figée.

    À appeler avant la première invocation de soffice sur ce profil : une fois
    le profil bootstrappé, le fichier serait écrasé. Voir `config.RENDER_LOCALE`
    pour la raison — sans ça le rendu n'est pas reproductible d'une machine à
    l'autre.
    """
    user = profile_dir / "user"
    user.mkdir(parents=True, exist_ok=True)
    (user / "registrymodifications.xcu").write_text(
        _REGISTRY_XCU.format(loc=locale), encoding="utf-8")
    return profile_dir


def _run(args: Sequence[str], outdir: Path, profile_dir: Path) -> None:
    # as_uri() produit file:///C:/... sous Windows et encode les espaces ;
    # une concaténation "file://" + chemin casse sur les deux points du lecteur
    # et sur les chemins contenant un espace.
    cmd = [
        find_soffice(), "--headless", "--norestore", "--invisible",
        f"-env:UserInstallation={profile_dir.resolve().as_uri()}",
        *args, "--outdir", str(outdir),
    ]
    # LANG/LC_ALL : c'est ce que LibreOffice suit sous Linux, où tourne la prod.
    loc = RENDER_LOCALE.replace("-", "_")
    env = {**os.environ, "LANG": f"{loc}.UTF-8", "LC_ALL": f"{loc}.UTF-8"}
    try:
        res = subprocess.run(cmd, capture_output=True, text=True,
                             timeout=SOFFICE_TIMEOUT, env=env)
    except FileNotFoundError as exc:
        raise RuntimeError(f"binaire LibreOffice injoignable : {cmd[0]}") from exc
    if res.returncode != 0:
        raise RuntimeError(
            f"soffice a échoué ({res.returncode}) : "
            f"{(res.stderr or res.stdout)[:400]}")


def _await_output(path: Path, label: str, timeout: float = 20.0) -> Path:
    """Attend l'apparition du fichier de sortie.

    Sous Windows, soffice peut rendre la main avant que l'écriture soit
    terminée ; un simple exists() immédiat échoue alors par intermittence.
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        if path.exists() and path.stat().st_size > 0:
            return path
        time.sleep(0.25)
    raise RuntimeError(
        f"{label} : sortie introuvable ({path.name}). Vérifie qu'aucune "
        "instance de LibreOffice n'est déjà ouverte.")


def recalc_with_libreoffice(src: Path, outdir: Path, profile_dir: Path) -> Path:
    """Recalcule les formules sans valeur en cache.

    Sans cette passe, un classeur généré par une lib (openpyxl, xlsxwriter)
    rend toutes ses cellules de formule VIDES : le screenshot est propre mais
    faux. Déclenchée par profile.verdict.needs_recalc.
    """
    _run(["--convert-to", "xlsx:Calc MS Excel 2007 XML", str(src)],
         outdir, profile_dir)
    return _await_output(outdir / (src.stem + ".xlsx"), "recalcul")


def to_pdf(src: Path, outdir: Path, profile_dir: Path) -> Path:
    """Convertit un classeur mis en scène en PDF paginé."""
    _run(["--convert-to", "pdf:calc_pdf_Export", str(src)], outdir, profile_dir)
    return _await_output(outdir / (src.stem + ".pdf"), "conversion PDF")
