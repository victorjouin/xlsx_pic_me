"""
xlsxtiles.numfmt — classification sémantique des formats de nombre OOXML.

Un code de format brut (`_("$"* #,##0.00_)`) ne dit rien d'exploitable tel
quel. Ce module le réduit à une catégorie — date, currency, percent… — qui
pilote ensuite l'estimation de largeur d'affichage, et que le manifeste
transmet à l'aval.
"""

from __future__ import annotations

import re

__all__ = ["classify_number_format", "BUILTIN_NUMFMT", "DATE_BUILTIN_IDS"]

# Formats de nombre intégrés (les seuls sans code explicite dans styles.xml).
# Seuls ceux qui portent une sémantique nous intéressent.
BUILTIN_NUMFMT = {
    0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00",
    9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
    14: "mm-dd-yy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy",
    18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss",
    22: "m/d/yy h:mm",
    37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)",
    39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)",
    45: "mm:ss", 46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@",
}
DATE_BUILTIN_IDS = set(range(14, 23)) | {45, 46, 47}

_CURRENCY_RE = re.compile(r"\[\$[^\]]*\]|[€$£¥₽¤]")
_DATE_TOKEN_RE = re.compile(r"(?<!\\)[ymdhs]")


def classify_number_format(fmt_id: int, code: str | None) -> str:
    """-> date | datetime | time | percent | currency | number | text | general

    C'est ce qui permet à l'extraction markdown de rendre 42005 en 01/01/2015
    et 1234.5 en 1 234,50 €. Sans ça, un classeur financier perd toute sa
    lisibilité dans le sidecar — perte plus grave qu'un mauvais cadrage.
    """
    if code is None:
        code = BUILTIN_NUMFMT.get(fmt_id)
    if fmt_id in DATE_BUILTIN_IDS:
        return "datetime" if fmt_id in (22,) else (
            "time" if fmt_id in (18, 19, 20, 21, 45, 46, 47) else "date")
    if not code or code == "General":
        return "general"
    if code == "@":
        return "text"

    # La devise se cherche sur le code BRUT : le symbole peut être entre
    # guillemets ("€"#,##0.00) ou dans un bloc locale ([$€-40C]).
    is_currency = bool(_CURRENCY_RE.search(code))

    # Pour les tokens de date, on nettoie d'abord ce qui n'en est pas :
    #  - littéraux entre guillemets
    #  - blocs crochetés SAUF les durées écoulées [h] [mm] [ss]
    #    (sans ça, [Red] fait passer un format monétaire pour une date)
    #  - caractères échappés par un antislash
    stripped = re.sub(r'"[^"]*"', "", code)
    stripped = re.sub(r"\[(?![hms]+\])[^\]]*\]", "", stripped)
    stripped = re.sub(r"\\.", "", stripped)

    low = stripped.lower().replace("am/pm", "").replace("a/p", "")
    if _DATE_TOKEN_RE.search(low):
        has_day = "d" in low or "y" in low
        has_clock = "h" in low or "s" in low
        if has_day and has_clock:
            return "datetime"
        return "date" if has_day else "time"
    if "%" in stripped:
        return "percent"
    if is_currency:
        return "currency"
    if any(c in stripped for c in "0#?"):
        return "number"
    return "general"
