"""
xlsxtiles.config — tous les seuils et budgets réglables, en un seul endroit.

Ce module n'importe rien du package : c'est la feuille de l'arbre de
dépendances. On peut le lire et l'ajuster sans rien connaître du reste.

Les valeurs marquées « à calibrer » dépendent du corpus réel ou de
l'environnement de rendu et ne doivent pas être prises pour des constantes
universelles.
"""

from __future__ import annotations

import os

# --------------------------------------------------------------------------
# Profilage : routage et profondeur de scan
# --------------------------------------------------------------------------

DEEP_SCAN_MAX_MB = 20.0        # au-delà, scan superficiel de la feuille
LAMBDA_MAX_MB = 30.0           # taille fichier au-delà de laquelle on sort de Lambda
LAMBDA_MAX_CELLS = 300_000     # cellules totales au-delà desquelles on sort de Lambda
BATCH_MIN_CELLS = 1_500_000
BYTES_PER_CELL_ESTIMATE = 55   # moyenne observée sur du XML non compressé
MAX_TRACKED_ROW_HEIGHTS = 100_000  # garde-fou mémoire sur les feuilles massives

# --------------------------------------------------------------------------
# Profilage : classification des feuilles  (à calibrer sur le corpus réel)
# --------------------------------------------------------------------------

SPARSE_FILL_RATIO = 0.05       # densité en dessous de laquelle une feuille est "éparse"
LAYOUT_MERGE_COUNT = 200       # nb de fusions au-delà duquel la feuille est une mise en page
CACHED_VALUE_MIN_RATIO = 0.90  # ratio de formules avec <v> en dessous duquel on recalcule
TEXT_DOMINANCE = 0.8           # part de cellules texte au-delà de laquelle la colonne
                               # est texte, quel que soit son format de nombre

# --------------------------------------------------------------------------
# Budgets de lisibilité
# --------------------------------------------------------------------------
# Le multivectorizer redimensionne l'image vers un budget de patches. Le texte
# doit rester lisible APRÈS ce redimensionnement, sinon la tuile est du bruit.

DEFAULT_PATCH_BUDGET = 1024          # ColQwen ~1024 patches de 28x28 px
PX_PER_PATCH = 28 * 28
MIN_TEXT_PX_AFTER_RESIZE = 10.0      # hauteur de capitale minimale
CALIBRI_11_CAP_PX = 11.0             # hauteur de capitale à 96 dpi
MIN_ROWS_PER_TILE = 5

# --------------------------------------------------------------------------
# Largeurs d'affichage
# --------------------------------------------------------------------------

MAX_AUTOFIT_CHARS = 45.0
AUTOFIT_SAMPLE_ROWS = 300

# --------------------------------------------------------------------------
# Rendu
# --------------------------------------------------------------------------

RENDER_DPI = 150
SOFFICE_TIMEOUT = 180
TRIM_PAD_PT = 2.0                    # marge laissée autour de l'encre au rognage

# Marge de sécurité sur la taille de page, GÉNÉREUSE et c'est délibéré : cf.
# render.prepare_workbook(). Une page trop grande ne coûte rien (on rogne au
# contenu), une page trop petite casse le mapping page -> tuile.
PAGE_SAFETY_W = 1.15
PAGE_SAFETY_H = 1.10

# Locale de rendu. NE PAS laisser LibreOffice suivre celle du système : les
# séparateurs de milliers/décimales ET l'ordre jour/mois en dépendent. Un poste
# de dev en fr-FR rend "32 370,00" et lit le format mm-dd-yy à l'envers
# (01/06/2014 = 1er juin) là où le conteneur en C/en-US rend "32,370.00" et
# 01-06-14 = 6 janvier. Même fichier, deux images, deux lectures contradictoires.
RENDER_LOCALE = os.environ.get("XLSX_RENDER_LOCALE", "en-US")

# --------------------------------------------------------------------------
# Mesure et calibration
# --------------------------------------------------------------------------

EDGE_CLUSTER_PT = 1.5                # tolérance de regroupement des arêtes mesurées
GRID_DRIFT_TOLERANCE_PX = 2.0        # au-delà, la tuile n'est pas déclarée vérifiée
PROBE_WIDTHS = [4.0, 8.0, 12.0, 16.0, 20.0, 24.0, 28.0, 32.0]
PROBE_HEIGHTS = [10.0, 15.0, 20.0, 25.0, 30.0, 40.0]
