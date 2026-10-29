#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_cable_list.py
======================
Génère automatiquement la CABLE LIST (feuille "EQUIPMENT") à partir d'une IO LIST.

Principe (identique à l'exemple CABLE_LIST.xlsx) :
  * Chaque LOOP démarre au FACP et retourne au FACP.
        FACP -> appareil 1 -> appareil 2 -> ... -> dernier appareil -> FACP
  * L'ordre des appareils = l'ordre des lignes de l'IO LIST.
  * Un appareil "hors boucle" (pas de n° de loop ou pas d'adresse : FD, PSH, SOV,
    LS, CT, CAM ...) n'est PAS sur la boucle : il est câblé en dérivation depuis
    le dernier appareil de boucle qui le précède (typiquement un module MIO).
    Le câble de la boucle continue ensuite depuis ce module, pas depuis la dérivation.
  * Les colonnes "DA" (FROM) / "A" (TO) / "ETICHETTA" sont des FORMULES Excel,
    comme dans le fichier d'exemple.

Usage :
    python generate_cable_list.py IO_list.xlsx
    python generate_cable_list.py IO_list.xlsx CABLE_LIST_out.xlsx
    python generate_cable_list.py IO_list.xlsx out.xlsx --sheet Foglio1

Dépendance : openpyxl   (pip install openpyxl)
"""
import argparse
import re
import sys
from collections import OrderedDict

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# =============================================================================
# PARAMÈTRES MODIFIABLES
# =============================================================================
CABLE_TYPE = "2x1,5mmq"      # colonne TIPO (modifiable ensuite à la main)
SHIELD = "SH"                # colonne SHIELD
ARMORED = "-"                # colonne BLINDATO
FACP_NAME = "FACP"           # nom de la centrale
FACP_DESCRIPTION = ""        # description affichée sur la ligne de retour FACP

# Texte à supprimer des descriptions (comme dans l'exemple : "-PORTA TAGLIAFUOCO")
DESCRIPTION_REMOVE = ["-PORTA TAGLIAFUOCO"]

# Préfixes qui servent de "concentrateur" pour les dérivations suivantes
# (ex. dans l'exemple : les CAM sont câblées depuis le CCP, lui-même sur un MIO)
HUB_PREFIXES = {"CCP"}

# Noms de colonnes reconnus dans l'IO LIST (en majuscules)
HEADER_ALIASES = {
    "loop":    ["LOOP"],
    "prefix":  ["PREFISSO", "PREFIX"],
    "number":  ["NUMERO DELL'APPARECCHIATURA", "NUMERO", "NUMBER"],
    "address": ["INDIRIZZO", "ADDRESS"],
    "desc":    ["DESCRIZIONE", "DESCRIPTION"],
    "area":    ["AREA", "ZONE"],
}

# =============================================================================
# LECTURE DE L'IO LIST
# =============================================================================
def norm(v):
    return re.sub(r"\s+", " ", str(v)).strip().upper() if v is not None else ""


def clean(v):
    """Texte propre ou '' (None, espaces, '-' => vide)."""
    s = "" if v is None else str(v).strip()
    return "" if s in ("-", "--") else s


def find_header(ws):
    """Retourne (n° ligne d'en-tête, {champ: n° colonne})."""
    for r in range(1, min(ws.max_row, 15) + 1):
        cols = {}
        heads = {c: norm(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)}
        for field, aliases in HEADER_ALIASES.items():
            for a in aliases:                       # 1) correspondance exacte
                hit = [c for c, h in heads.items() if h == a]
                if hit:
                    cols[field] = hit[0]
                    break
            else:
                for a in aliases:                   # 2) début de mot
                    hit = [c for c, h in heads.items() if h.startswith(a)]
                    if hit:
                        cols[field] = hit[0]
                        break
        if {"loop", "prefix", "number"} <= set(cols):
            return r, cols
    return None, None


def read_io_list(path, sheet=None):
    wb = load_workbook(path, data_only=False)
    sheets = [wb[sheet]] if sheet else wb.worksheets
    for ws in sheets:
        hr, cols = find_header(ws)
        if hr:
            break
    else:
        sys.exit("ERREUR : en-têtes LOOP / PREFISSO / NUMERO introuvables dans l'IO LIST.")

    def cell(r, f):
        return ws.cell(r, cols[f]).value if f in cols else None

    devices = []
    for r in range(hr + 1, ws.max_row + 1):
        prefix = clean(cell(r, "prefix"))
        if not prefix:
            continue
        num = cell(r, "number")
        num = f"{num:02d}" if isinstance(num, (int, float)) else clean(num)
        desc = clean(cell(r, "desc"))
        for rm in DESCRIPTION_REMOVE:
            desc = re.sub(re.escape(rm), "", desc, flags=re.I)
        devices.append({
            "src_row": r,
            "loop": clean(cell(r, "loop")),
            "prefix": prefix,
            "num": num,
            "addr": clean(cell(r, "address")),
            "desc": re.sub(r"\s+", " ", desc).strip(" -"),
            "area": clean(cell(r, "area")),
        })
    return ws.title, devices


# =============================================================================
# CONSTRUCTION DES BOUCLES  (FACP -> ... -> FACP)
# =============================================================================
def build_rows(devices):
    """Retourne la liste ordonnée des lignes de la cable list + les avertissements."""
    warnings = []
    loops = OrderedDict()
    current = None
    for d in devices:
        if d["loop"]:
            current = d["loop"]
        if current is None:
            warnings.append(f"IO LIST ligne {d['src_row']} ({d['prefix']}-{d['num']}) : "
                            f"aucun loop avant cet appareil -> ignoré")
            continue
        d["on_loop"] = bool(d["loop"] and d["addr"])      # appareil adressé sur la boucle
        d["loop_of"] = current
        loops.setdefault(current, []).append(d)

    rows = []                         # chaque élément = dict décrivant une ligne
    seq_row = lambda: len(rows) + 2   # n° de ligne Excel de la prochaine ligne

    for loop, devs in loops.items():
        prev_row = None               # dernière ligne "appareil de boucle"
        hub_row = None                # concentrateur actif (CCP)
        last_area = devs[0]["area"]
        tags = {}
        for d in devs:
            tag = f"{d['prefix']}-{d['num']}"
            tags[tag] = tags.get(tag, 0) + 1
            r = seq_row()
            if d["on_loop"]:
                src = FACP_NAME if prev_row is None else f"=J{prev_row}"
                prev_row, hub_row = r, None
            else:
                parent = hub_row or prev_row
                if parent is None:
                    warnings.append(f"{loop} {tag} : dérivation sans appareil précédent -> câblée au {FACP_NAME}")
                    src = FACP_NAME
                else:
                    src = f"=J{parent}"
                if d["prefix"] in HUB_PREFIXES:
                    hub_row = r
            if d["desc"] == "":
                warnings.append(f"{loop} {tag} : description vide dans l'IO LIST")
            last_area = d["area"] or last_area
            rows.append({
                "tag": tag, "addr": d["addr"],
                # D (LOOP) : renseigné si l'IO LIST donne un loop à l'appareil
                "loop": d["loop"] if d["loop"] else "",
                "src": src, "desc": d["desc"], "area": d["area"], "kind": "device",
            })
        for t, n in tags.items():
            if n > 1:
                warnings.append(f"{loop} : tag '{t}' présent {n} fois -> étiquettes de câble en double")
        # ligne de retour au FACP
        r = seq_row()
        rows.append({
            "tag": FACP_NAME, "addr": "", "loop": loop,
            "src": FACP_NAME if prev_row is None else f"=J{prev_row}",
            "desc": FACP_DESCRIPTION, "area": last_area, "kind": "return",
        })
    return rows, loops, warnings


# =============================================================================
# ÉCRITURE DU FICHIER EXCEL
# =============================================================================
def write_workbook(rows, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "EQUIPMENT"

    font = Font(name="Calibri", size=14)
    head_font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="4472C4")
    yellow = PatternFill("solid", fgColor="FFFF00")
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # (colonne de départ, colonne de fin pour fusion, en-tête)
    layout = [
        (1, 1, "ARTICOLO"), (2, 2, "ATTREZZATURA"), (3, 3, "ETICHETTA PER CAVI"),
        (4, 7, "DA"), (8, 8, "A"), (9, 10, "TIPO"),
        (11, 11, "SHIELD"), (12, 12, "BLINDATO"), (13, 13, "LUNGHEZZA (m)"),
        (14, 17, "DESCRIZIONE"), (18, 21, "AREA"),
    ]

    def put(row, c1, c2, value, fill=None, fnt=font, fmt=None):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row, c)
            cell.border, cell.alignment, cell.font = border, center, fnt
            if fill:
                cell.fill = fill
        ws.cell(row, c1).value = value
        if fmt:
            ws.cell(row, c1).number_format = fmt
        if c2 > c1:
            ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)

    for c1, c2, title in layout:
        put(1, c1, c2, title, fill=head_fill, fnt=head_font)
    ws.row_dimensions[1].height = 40

    for i, d in enumerate(rows):
        r = i + 2
        label = f'=CONCATENATE("{d["loop"]}","-",B{r})' if d["loop"] and d["kind"] == "device" \
            else f"=CONCATENATE(B{r})"
        put(r, 1, 1, i + 1, fmt="0")
        put(r, 2, 2, d["tag"], fmt="@")
        put(r, 3, 3, f"=H{r}")
        put(r, 4, 7, d["src"].replace("=J", "=H"))
        put(r, 8, 8, label)
        put(r, 9, 10, CABLE_TYPE, fill=yellow)           # jaune = à vérifier / modifier
        put(r, 11, 11, SHIELD)
        put(r, 12, 12, ARMORED)
        put(r, 13, 13, None, fill=yellow, fmt="0")       # jaune = longueur à saisir
        put(r, 14, 17, d["desc"] or None)
        put(r, 18, 21, d["area"] or None)
        ws.row_dimensions[r].height = 20.1

    # barre de fermeture du tableau (comme l'exemple)
    end = len(rows) + 2
    for c in range(1, 22):
        ws.cell(end, c).fill = PatternFill("solid", fgColor="000000")
    ws.merge_cells(start_row=end, start_column=1, end_row=end, end_column=21)
    ws.row_dimensions[end].height = 6

    widths = {"A": 10, "B": 16, "C": 26, "D": 7, "E": 7, "F": 7, "G": 7, "H": 26,
              "I": 8, "J": 8, "K": 10, "L": 13, "M": 17,
              "N": 30, "O": 10, "P": 10, "Q": 10, "R": 20, "S": 12, "T": 12, "U": 12}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:U{end - 1}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(out_path)


# =============================================================================
# MAIN
# =============================================================================
def main():
    ap = argparse.ArgumentParser(description="Génère la CABLE LIST à partir de l'IO LIST.")
    ap.add_argument("io_list", help="fichier IO LIST (.xlsx)")
    ap.add_argument("output", nargs="?", default="CABLE_LIST_generated.xlsx",
                    help="fichier de sortie (défaut : CABLE_LIST_generated.xlsx)")
    ap.add_argument("--sheet", help="nom de la feuille de l'IO LIST (auto-détection par défaut)")
    args = ap.parse_args()

    sheet, devices = read_io_list(args.io_list, args.sheet)
    rows, loops, warnings = build_rows(devices)
    write_workbook(rows, args.output)

    print(f"IO LIST lue (feuille '{sheet}') : {len(devices)} appareils, {len(loops)} loops")
    for loop, devs in loops.items():
        n_loop = sum(d["on_loop"] for d in devs)
        print(f"  {loop} : {n_loop} appareils sur boucle + {len(devs) - n_loop} dérivations "
              f"-> {len(devs) + 1} câbles (retour {FACP_NAME} inclus)")
    print(f"Cable list écrite : {args.output}  ({len(rows)} lignes)")
    if warnings:
        print(f"\n{len(warnings)} point(s) à vérifier :")
        for w in warnings:
            print("  -", w)


if __name__ == "__main__":
    main()
