"""
xlsxtiles.encode — transport des tuiles en base64.

Permet de faire circuler une image sans dépendre d'un système de fichiers
partagé entre les étages de la pipeline : la tuile voyage dans le même JSON que
sa plage A1.

Contrepartie à connaître : le base64 pèse 4/3 de l'original, et un manifeste
qui embarque toutes ses tuiles doit être chargé en entier pour être parsé.
Mesuré sur `Financial Sample.xlsx` (83 Ko) : 56 tuiles, 12,2 Mo de PNG, donc
~16 Mo de manifeste. C'est utilisable, mais ce n'est pas gratuit — d'où
l'option, jamais le comportement par défaut.
"""

from __future__ import annotations

import base64
from pathlib import Path

__all__ = ["PNG_DATA_URI_PREFIX", "png_to_base64", "png_to_data_uri",
           "base64_to_png"]

PNG_DATA_URI_PREFIX = "data:image/png;base64,"


def png_to_base64(path: str | Path) -> str:
    """PNG sur disque -> base64 brut, sans préfixe.

    Renvoyer le base64 nu plutôt qu'une data URI laisse l'appelant décider :
    un `<img src>` veut le préfixe, un champ binaire OpenSearch le refuse.
    """
    return base64.b64encode(Path(path).read_bytes()).decode("ascii")


def png_to_data_uri(path: str | Path) -> str:
    """PNG sur disque -> data URI directement affichable dans un navigateur."""
    return PNG_DATA_URI_PREFIX + png_to_base64(path)


def base64_to_png(data: str, dest: str | Path) -> Path:
    """Écrit un PNG depuis son base64, avec ou sans préfixe data URI."""
    if data.startswith(PNG_DATA_URI_PREFIX):
        data = data[len(PNG_DATA_URI_PREFIX):]
    dest = Path(dest)
    dest.write_bytes(base64.b64decode(data))
    return dest
