"""
Pilotage de LibreOffice en headless.

Localisation du binaire, profil utilisateur jetable à locale figée, conversion
en PDF. Tout ce qui touche au processus externe est ici ; le reste du package
ignore que LibreOffice existe.

Pas de passe de recalcul : mesuré, la conversion PDF évalue déjà les formules
sans valeur en cache.
"""

from __future__ import annotations

import functools
import os
import shutil
import subprocess
import time
from pathlib import Path

__all__ = ["find_soffice", "seed_profile", "to_pdf"]

TIMEOUT = 180

# Locale de rendu. NE PAS laisser LibreOffice suivre celle du système : les
# séparateurs de milliers/décimales ET l'ordre jour/mois en dépendent. Un poste
# de dev en fr-FR rend "32 370,00" et lit le format mm-dd-yy à l'envers
# (01/06/2014 = 1er juin) là où le conteneur en C/en-US rend "32,370.00" et
# 01-06-14 = 6 janvier. Même fichier, deux images, deux lectures contradictoires.
LOCALE = os.environ.get("XLSX_RENDER_LOCALE", "en-US")

# LibreOffice ne s'ajoute PAS au PATH sous Windows et n'est pas au même endroit
# selon l'OS.
_CANDIDATES = {
    "nt": [
        r"C:\Program Files\LibreOffice\program\soffice.com",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.com",
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
    """Localise le binaire. Surchargeable par la variable SOFFICE_PATH.

    Sous Windows on privilégie soffice.com sur soffice.exe : le .exe se détache
    du process appelant et subprocess.run rend la main avant que le PDF existe.
    """
    env = os.environ.get("SOFFICE_PATH")
    if env and Path(env).exists():
        return env
    names = (["soffice.com", "soffice.exe"] if os.name == "nt"
             else ["soffice", "libreoffice"])
    for name in names:
        found = shutil.which(name)
        if found:
            return found
    for cand in _CANDIDATES.get(os.name, []):
        if Path(cand).exists():
            return cand
    raise RuntimeError(
        "LibreOffice introuvable. Installe-le, ou renseigne SOFFICE_PATH avec "
        r"le chemin complet du binaire (Windows : "
        r"C:\Program Files\LibreOffice\program\soffice.com).")


def seed_profile(profile_dir: Path, locale: str = LOCALE) -> Path:
    """Prépare un profil LibreOffice neuf à locale figée.

    À appeler AVANT la première invocation sur ce profil : une fois bootstrappé,
    le fichier serait écrasé.
    """
    user = profile_dir / "user"
    user.mkdir(parents=True, exist_ok=True)
    (user / "registrymodifications.xcu").write_text(
        _REGISTRY_XCU.format(loc=locale), encoding="utf-8")
    return profile_dir


def to_pdf(src: Path, outdir: Path, profile_dir: Path) -> Path:
    """Convertit un classeur mis en scène en PDF paginé."""
    # as_uri() produit file:///C:/... et encode les espaces ; une concaténation
    # "file://" + chemin casse sur les deux points du lecteur et sur les espaces.
    cmd = [
        find_soffice(), "--headless", "--norestore", "--invisible",
        f"-env:UserInstallation={profile_dir.resolve().as_uri()}",
        "--convert-to", "pdf:calc_pdf_Export", str(src),
        "--outdir", str(outdir),
    ]
    # LANG/LC_ALL : c'est ce que LibreOffice suit sous Linux, où tourne la prod.
    loc = LOCALE.replace("-", "_")
    env = {**os.environ, "LANG": f"{loc}.UTF-8", "LC_ALL": f"{loc}.UTF-8"}
    try:
        res = subprocess.run(cmd, capture_output=True, text=True,
                             timeout=TIMEOUT, env=env)
    except FileNotFoundError as exc:
        raise RuntimeError(f"binaire LibreOffice injoignable : {cmd[0]}") from exc
    if res.returncode != 0:
        raise RuntimeError(f"soffice a échoué ({res.returncode}) : "
                           f"{(res.stderr or res.stdout)[:400]}")

    # Sous Windows, soffice peut rendre la main avant la fin de l'écriture ;
    # un exists() immédiat échoue alors par intermittence.
    out = outdir / (src.stem + ".pdf")
    deadline = time.time() + 20.0
    while time.time() < deadline:
        if out.exists() and out.stat().st_size > 0:
            return out
        time.sleep(0.25)
    raise RuntimeError(
        f"PDF introuvable ({out.name}). Vérifie qu'aucune instance de "
        "LibreOffice n'est déjà ouverte.")
