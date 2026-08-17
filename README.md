# xlsx_pic_me

Découpe un classeur Excel en images fidèles, chacune accompagnée de la plage A1
exacte qu'elle représente.

Destiné à une pipeline d'ingestion documentaire : les images partent à la
vectorisation, et la plage permet de rendre à l'utilisateur le morceau d'Excel
correspondant à l'information citée. **Le périmètre s'arrête au découpage** —
vectorisation, indexation et restitution sont des étages aval.

## Installation

```bash
pip install -r requirements.txt
```

LibreOffice doit être installé séparément (voir `requirements.txt`) : c'est lui
qui convertit le classeur en PDF.

## Utilisation

```bash
python -m xlsxtiles render classeur.xlsx -o manifest.json
```

```python
from xlsxtiles import render_workbook

manifest = render_workbook("classeur.xlsx")
tile = manifest["sheets"][0]["tiles"][0]
tile["range_a1"]    # "A2:G25"
tile["png_base64"]  # les octets du PNG, sans préfixe
```

Le préfixe de data URI (`data:image/png;base64,`) est fourni à part dans
`manifest["data_uri_prefix"]` : un `<img src>` le veut, un champ binaire
OpenSearch le refuse.

Deux commandes d'appoint : `plan` affiche le découpage sans payer les
conversions LibreOffice, `extract` réécrit les PNG d'un manifeste sur disque
pour les inspecter à l'œil.

## Ce qui garantit la fidélité

**La locale est figée à `en-US`.** Sans ça, le rendu n'est pas reproductible :
un poste en fr-FR affiche `32 370,00` et lit le format `mm-dd-yy` à l'envers
(`01/06/2014` = 1er juin), là où un conteneur en C/en-US affiche `32,370.00` et
`6/1/2014`. Même fichier, deux images, deux lectures contradictoires.
Surchargeable par `XLSX_RENDER_LOCALE`.

**Les colonnes sont élargies avant rendu.** Une colonne numérique trop étroite
n'affiche pas un texte tronqué : elle affiche `###`, perte totale. La largeur
voulue par l'auteur est un plancher, jamais un plafond.

**La page PDF est surdimensionnée puis rognée au contenu réel.** C'est le
principe qui tient tout l'édifice : une page trop grande ne coûte rien, le
rognage l'efface ; une page trop petite fait déborder la tuile sur une seconde
page. C'est pourquoi la géométrie prédite n'a pas besoin d'être exacte — et
pourquoi il n'y a pas de calibration dans ce code.

**Une bande dont la pagination ne correspond pas au plan est reprise tuile par
tuile.** On ne devine jamais quelle page porte quelle plage : livrer une image
dont on ignore la plage serait pire que payer N conversions.

**Aucune coupe ne tombe au milieu d'un graphique ou d'une image.** Deux
demi-graphiques ne portent aucune information.

## Limites assumées

- Le profilage passe par openpyxl, donc **charge tout le classeur en mémoire** :
  inadapté à un classeur de plusieurs centaines de Mo.
- Les largeurs d'affichage sont **estimées**, pas mesurées ; l'estimation
  majore, quitte à produire des colonnes plus larges que nécessaire.
- Le manifeste embarquant les images pèse ~4/3 du poids total des PNG (17 Mo
  pour le classeur de test de 83 Ko) et doit être chargé en entier pour être
  parsé.
- Pas de surlignage de la cellule citée : il exigerait une grille pixel exacte,
  donc une passe de calibration retirée volontairement.

## Structure

```
xlsxtiles/
  profile.py   géométrie des feuilles, via openpyxl
  soffice.py   pilotage de LibreOffice
  tiles.py     plan de découpe, mise en scène, rasterisation, manifeste
  __main__.py  CLI
```
