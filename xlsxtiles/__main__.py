"""
Ligne de commande de xlsxtiles.

    python -m xlsxtiles render <fichier.xlsx> <outdir> [--dpi N] [--patch-budget N]
    python -m xlsxtiles profile <fichier.xlsx>
    python -m xlsxtiles calibrate [--refresh]
"""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import asdict
from pathlib import Path

from .calibration import load_calibration
from .config import DEFAULT_PATCH_BUDGET, RENDER_DPI
from .manifest import render_workbook
from .profiler import profile_workbook


def _cmd_render(args: argparse.Namespace) -> int:
    m = render_workbook(args.source, args.outdir, dpi=args.dpi,
                        patch_budget=args.patch_budget,
                        include_base64=args.base64)
    for sh in m["sheets"]:
        print(f"{sh['sheet']}: {sh['n_tiles']} tuile(s) "
              f"[{sh['image_indexing']} / {sh['tiling_strategy']}]")
        for t in sh["tiles"]:
            flag = "" if t["grid_verified"] else "  NON VÉRIFIÉE"
            print(f"  {t['range_a1']:>16}  {t['width_px']}x{t['height_px']}px  "
                  f"dérive={t['grid_drift_px']}  {t['png']}{flag}")
    for w in m["warnings"]:
        print(f"  ! {w}")
    mf = Path(args.outdir) / "manifest.json"
    size = mf.stat().st_size / 1e6
    print(f"\nmanifeste : {mf} ({size:.1f} Mo"
          f"{', images incluses en base64' if args.base64 else ''})")
    return 0


def _cmd_profile(args: argparse.Namespace) -> int:
    print(profile_workbook(args.source).to_json())
    return 0


def _cmd_calibrate(args: argparse.Namespace) -> int:
    cal = load_calibration(refresh=args.refresh)
    print(json.dumps(asdict(cal), indent=2))
    if not cal.calibrated:
        print("calibration NON aboutie", file=sys.stderr)
        return 1
    return 0


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="python -m xlsxtiles",
                                description="Découpe un classeur Excel en images.")
    sub = p.add_subparsers(dest="cmd", required=True)

    pr = sub.add_parser("render", help="découpe un classeur en PNG + manifest.json")
    pr.add_argument("source")
    pr.add_argument("outdir")
    pr.add_argument("--dpi", type=int, default=RENDER_DPI)
    pr.add_argument("--patch-budget", type=int, default=DEFAULT_PATCH_BUDGET)
    pr.add_argument("--base64", action="store_true",
                    help="embarque chaque PNG dans le manifeste (champ "
                         "png_base64) ; le manifeste pèse alors ~4/3 du "
                         "poids total des images")
    pr.set_defaults(func=_cmd_render)

    pp = sub.add_parser("profile", help="affiche le profil JSON du classeur")
    pp.add_argument("source")
    pp.set_defaults(func=_cmd_profile)

    pc = sub.add_parser("calibrate", help="mesure (ou relit) la calibration pixel")
    pc.add_argument("--refresh", action="store_true",
                    help="ignore le cache et relance la feuille sonde")
    pc.set_defaults(func=_cmd_calibrate)

    args = p.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
