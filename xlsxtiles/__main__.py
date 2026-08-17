"""
Ligne de commande.

    python -m xlsxtiles render <fichier.xlsx> [-o manifest.json] [--dpi 150]
    python -m xlsxtiles plan   <fichier.xlsx>
    python -m xlsxtiles extract <manifest.json> <outdir>
"""

from __future__ import annotations

import argparse
import base64
import json
import sys
from pathlib import Path

from .profile import profile_workbook
from .tiles import DPI, plan_tiles, render_workbook


def _cmd_render(args: argparse.Namespace) -> int:
    m = render_workbook(args.source, dpi=args.dpi)
    payload = json.dumps(m, ensure_ascii=False, indent=2)

    for sh in m["sheets"]:
        print(f"{sh['sheet']} ({sh['range_a1']}) : {sh['n_tiles']} tuile(s)",
              file=sys.stderr)
        for t in sh["tiles"]:
            print(f"  {t['range_a1']:>16}  {t['width_px']}x{t['height_px']}px  "
                  f"{len(t['png_base64']) / 1024:.0f} Ko base64", file=sys.stderr)

    if args.output:
        Path(args.output).write_text(payload, encoding="utf-8")
        mb = len(payload.encode()) / 1e6
        print(f"\n{args.output} ({mb:.1f} Mo)", file=sys.stderr)
    else:
        print(payload)
    return 0


def _cmd_plan(args: argparse.Namespace) -> int:
    """Plan de découpe sans rendu — utile pour vérifier un cadrage sans payer
    les conversions LibreOffice."""
    for sheet in profile_workbook(args.source):
        tiles = plan_tiles(sheet)
        print(f"{sheet.name} ({sheet.range_a1}) : {len(tiles)} tuile(s), "
              f"en-tête {sheet.header_rows} ligne(s), "
              f"{len(sheet.anchors)} zone(s) incoupable(s)")
        for t in tiles:
            print(f"  {t.range_a1}")
    return 0


def _cmd_extract(args: argparse.Namespace) -> int:
    """Réécrit les PNG sur disque depuis un manifeste — pour inspecter à l'œil."""
    m = json.loads(Path(args.manifest).read_text(encoding="utf-8"))
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    n = 0
    for sh in m["sheets"]:
        for t in sh["tiles"]:
            name = f"{sh['sheet']}_{t['index']:03d}.png".replace("/", "_")
            (outdir / name).write_bytes(base64.b64decode(t["png_base64"]))
            n += 1
    print(f"{n} image(s) écrite(s) dans {outdir}")
    return 0


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="python -m xlsxtiles",
                                description="Découpe un classeur Excel en images.")
    sub = p.add_subparsers(dest="cmd", required=True)

    pr = sub.add_parser("render", help="découpe et produit le manifeste JSON")
    pr.add_argument("source")
    pr.add_argument("-o", "--output", help="fichier de sortie (stdout par défaut)")
    pr.add_argument("--dpi", type=int, default=DPI)
    pr.set_defaults(func=_cmd_render)

    pp = sub.add_parser("plan", help="affiche le plan de découpe, sans rendu")
    pp.add_argument("source")
    pp.set_defaults(func=_cmd_plan)

    pe = sub.add_parser("extract", help="réécrit les PNG d'un manifeste sur disque")
    pe.add_argument("manifest")
    pe.add_argument("outdir")
    pe.set_defaults(func=_cmd_extract)

    args = p.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
