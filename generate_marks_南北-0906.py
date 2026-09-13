#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Shipping mark generator (Excel output)
======================================

Reads the "Shipping details ... xlsx" workbook and produces a shipping-mark
workbook: one sheet per pallet, each sheet laid out like the reference
"Wichita Batch2-\u5473\u5934 1-2 \u5927\u67dc.pdf" (supplier / project /
LOT info, part list with quantities, weight, dimensions, storage info,
block numbers and a photo placeholder).

Business rules (North-South only)
1. One row on the main sheet (e.g. "Block 5-8 Packing List") = one
   combiner box (CMB.xxx):
   - column A = box number, column B = "BOX  TYPE" (box type),
     columns D..AF = products (N/S PPH collectors + SH cables)
   - column "Drum Size" (first) = N/S drum size, e.g. 1000*350*500
     (diameter * bore * height; only diameter and height matter)
   - column "QTY 1" (first) = number of N/S drums of this box
   - EW* product columns and "QTY 2" (East-West) are ignored on purpose.
2. Column B must be headed "BOX  TYPE" and every box row must carry a box
   type; otherwise the run stops with an error. Boxes are classified by
   that value: a maximal run of consecutive same-type boxes is packed on
   its own, so a pallet never mixes two box types while same-type boxes do
   share pallets. Plain sheet row order no longer decides the packing, so
   boxes of another type sitting between two same-type blocks cannot end
   up on their pallets.
3. The background color of the "QTY 1" cell groups drums by container;
   drums of different colors must never share a pallet.
4. Packing per (container-color group, drum size, box-type run), in sheet
   row order inside a run:
   - diameter 1000 -> pallet 1050*1050*1150, 2 drums per pallet stacked;
     one leftover -> pallet 1050*1050*650
   - diameter 750  -> pallet 1500*750*1150, 4 drums per pallet (2x2);
     leftover <=2 -> pallet 1500*750*650
   - products per drum = cell value / drum count of that box. If any
     product count is not divisible by the drum count (e.g. 5 items on
     2 drums), the whole box stays contiguous on one pallet (it may share
     a pallet with later boxes only as an indivisible block).
   - the leftover partial pallets of a group's runs are then combined
     first-fit (up to the pallet capacity), so an odd drum is not shipped
     on a pallet of its own.
5. Pallet weight = net weight of all products + pallet 25 kg
   + drum weight (28 kg for diameter 1000, 25 kg for 750).
   A PPH product = one RED + one BLK unit pair; the pair is counted once.
6. Output: Marks.xlsx, one sheet per pallet (sheet name "LOT n-total").

Reusability
- Column headers are FIXED (row1 group names PPH/SH, row2 product names,
  row4 RED part numbers, row5 BLK part numbers); rows and values may
  change freely. Columns are located by header text, not letters.
- Shipment header info (supplier, project, SO/PO...) is in CONFIG below.
- Part-number -> description mapping: built-in templates (PPH/SH);
  override any of them in CONFIG["descriptions"].

Usage (run from the folder that holds the Excel):
    python3 generate_marks.py "Shipping details for Wichita BLOCK  1-4(3).xlsx"
    python3 generate_marks.py xxx.xlsx --sheet "Block 5-8 Packing List" \
        --gw-sheet Batch_01_GW --out Marks.xlsx
Options:
    --sheet    packing-list sheet name (default "Block 5-8 Packing List")
    --gw-sheet net weight sheet name  (default "Batch_01_GW")
    --out      output workbook file name (no folders, default "Marks.xlsx")
Deps:  pip install openpyxl
"""

import argparse
import os
import re
import sys
from collections import OrderedDict

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, Side
except ImportError:
    sys.exit("missing openpyxl, run: pip install openpyxl")


# ----------------------------------------------------------------------------
# CONFIG -- edit for every new shipment/batch (keys are fixed, values free)
# ----------------------------------------------------------------------------
CONFIG = {
    "supplier_id": "4047270",
    "country_of_origin": "THAILAND",
    "project_number": "3889589",
    "project_name": "Wichita TX - 805MW - eBOS WHA",
    "so_number": "40225",
    "po_number": "UAE90711",
    "type_of_storage": "INDOOR",
    "stackability": "NO",
    "pallet_weight_kg": 25.0,                        # fixed pallet mass (any size)
    "axle_weight_kg": {"1000": 28.0, "750": 25.0},   # drum weight by diameter
    # drum size -> (main pallet size, drums per pallet,
    #               leftover pallet size, max drums on leftover pallet)
    "pallet_sizes": {
        "1000*350*500": ["1050*1050*1150", 2, "1050*1050*650", 1],
        "750*350*500":  ["1500*750*1150", 4, "1500*750*650", 2],
    },
    # part-number -> description overrides; default templates are built in
    "descriptions": {},
}


def resolve_input(path):
    """Return the absolute path of an existing input file; abort otherwise."""
    if not isinstance(path, str) or not path.strip():
        sys.exit("input file path is empty")
    raw = os.path.expanduser(path.strip())
    for seg in raw.replace("\\", "/").split("/"):
        if seg == "..":
            sys.exit("invalid path (dot-dot segment): %s" % raw)
    p = os.path.normpath(os.path.abspath(raw))
    if not os.path.isfile(p):
        sys.exit("input file not found: %s" % p)
    return p


# ----------------------------------------------------------------------------
# Part description templates (match the reference PDF)
# ----------------------------------------------------------------------------
def describe_part(pn, product_key):
    """Build the PART DESCRIPTION for a part number + product key."""
    if pn.startswith("PPH-"):
        # PPH-9.P.RED.5.00007 + "9-1 T5" -> PPH,MRTSN,WCHTA,PPH,9-1,TP5,POS,RED
        # product keys may look like "10 -1 T1" (the 10-1 block is written
        # with a space), so the group is the key minus the trailing type
        # token, with the space dropped: "10 -1 T4" -> "10-1", "9-1 T1" -> "9-1"
        parts = product_key.strip().split()
        g = "".join(parts[:-1])
        t = parts[-1]
        tp = "TP" + t[1:] if t.startswith("T") else t
        pos_neg = "POS,RED" if ".P.RED." in pn else "NEG,BLK"
        return "PPH,MRTSN,WCHTA,PPH,%s,%s,%s" % (g, tp, pos_neg)
    if pn.startswith("SH-"):
        # SH-2.8.RED.BLK.00003 + "T2" -> MRTSN,WCHTA,SH,N/S,PAIR 2,#8CU-#6AL
        n = int(re.match(r"^T(\d+)$", product_key.strip()).group(1))
        desc = "MRTSN,WCHTA,SH,N/S,PAIR %d,#8CU-#6AL" % n
        if n >= 10:                                  # reference: PAIR 10 has XTRACKER
            desc += ",XTRACKER"
        return desc
    return pn


def fmt_num(x):
    """Trim trailing zeros: 465.86 / 465.1 / 445 / 354"""
    s = ("%.2f" % round(float(x), 2)).rstrip("0").rstrip(".")
    return s if s else "0"


# ----------------------------------------------------------------------------
# Excel reading
# ----------------------------------------------------------------------------
def norm_text(value):
    """Collapse spaces in header text for tolerant comparisons."""
    if not isinstance(value, str):
        return None
    return re.sub(r"\s+", " ", value).strip()


def validate_packing_list_fields(ws):
    """Sanity-check rows 1-2 (columns A..AU) and the first data row.

    A1/B1/C1 carry the header labels, AO/AR/AT/AU carry 'Drum Size' /
    'QTY 1' / 'Drum Size' / 'QTY 2', row2 of the product columns carries
    the product keys; the data area starts at row 6 with 'CMB...'.
    B1 is the box type column the pallet packing is grouped by, so it is
    checked first and with its own message.
    """
    def expect(row, col, want):
        got = norm_text(ws.cell(row=row, column=col).value)
        if got != want:
            sys.exit("packing list header error: %s%d must be %r, got %r"
                     % (openpyxl.utils.get_column_letter(col), row, want, got))

    b1_raw = ws.cell(row=1, column=2).value
    if norm_text(b1_raw) != "BOX TYPE":
        sys.exit("packing list header error: B1 must be 'BOX  TYPE' "
                 "(the box type column the pallets are grouped by), got %r"
                 % (b1_raw,))
    expect(1, 1, "COMBINER BOX NO.")
    expect(1, 3, "Harness Qty")
    expect(1, 41, "Drum Size")          # AO  (N/S drum size)
    expect(1, 44, "QTY 1")              # AR
    expect(1, 46, "Drum Size")          # AT  (E/W drum size)
    expect(1, 47, "QTY 2")              # AU
    key_pat = re.compile(r"^(\d+ ?-1 T\d+|T\d+|EW\d+)$")
    n_products = 0
    first_group = norm_text(ws.cell(row=1, column=4).value)
    if first_group not in ("PPH", "SH"):
        sys.exit("packing list header error: D1 must be PPH or SH, got %r"
                 % first_group)
    for c in range(4, ws.max_column + 1):            # D..AU
        r1 = norm_text(ws.cell(row=1, column=c).value)
        r2 = ws.cell(row=2, column=c).value
        if not isinstance(r1, str) or not isinstance(r2, str):
            continue
        k = r2.strip()
        if r1 in ("PPH", "SH") and key_pat.match(k):
            n_products += 1
        else:
            sys.exit("packing list header error: column %s row1/row2 = %r/%r"
                     % (openpyxl.utils.get_column_letter(c), r1, k))
    if n_products < 1:
        sys.exit("packing list header error: no product columns found (D..AM)")
    a6 = ws.cell(row=6, column=1).value
    if not (isinstance(a6, str) and a6.startswith("CMB")):
        sys.exit("packing list data error: A6 must start with 'CMB', got %r"
                 % a6)


def validate_gw_sheet(ws, gw):
    """The weight sheet is the transposed product table of the packing list.

    column A = row1 group name (PPH/SH), column B = row2 product key,
    column C = net weight (a positive number); C1 must be 'NW'; columns
    D+ are ignored. The A/B rows and the D..AU rows 1/2 must match.
    """
    if norm_text(gw.cell(row=1, column=3).value) != "NW":
        sys.exit("weight sheet header error: C1 must be 'NW', got %r"
                 % gw.cell(row=1, column=3).value)
    sheet_pairs = set()
    for c in range(4, ws.max_column + 1):            # D..AU
        r1 = norm_text(ws.cell(row=1, column=c).value)
        r2 = ws.cell(row=2, column=c).value
        if isinstance(r1, str) and isinstance(r2, str):
            sheet_pairs.add((r1, r2.strip()))
    gw_pairs = set()
    for r in range(4, gw.max_row + 1):
        a = gw.cell(row=r, column=1).value
        b = gw.cell(row=r, column=2).value
        c = gw.cell(row=r, column=3).value
        if a is None and b is None:
            break
        if not (isinstance(a, str) and a.strip()) or \
           not (isinstance(b, str) and b.strip()):
            sys.exit("weight sheet error: row %d needs text in A and B" % r)
        grp = norm_text(a)
        key = b.strip()
        if grp not in ("PPH", "SH"):
            sys.exit("weight sheet error: row %d group A must be PPH or SH, "
                     "got %r" % (r, a))
        if not isinstance(c, (int, float)) or float(c) <= 0:
            sys.exit("weight sheet error: row %d C must be a positive number, "
                     "got %r" % (r, c))
        gw_pairs.add((grp, key))
    if gw_pairs != sheet_pairs:
        missing = sorted(sheet_pairs - gw_pairs)
        extra = sorted(gw_pairs - sheet_pairs)
        sys.exit("weight sheet mismatch with packing list: missing=%s extra=%s"
                 % (missing, extra))


def cell_fill_key(cell):
    """Normalize a cell background color to a comparable string."""
    f = cell.fill
    if f is None or not f.patternType:
        return "none"
    fg = f.fgColor
    try:
        if fg.type == "rgb" and isinstance(fg.rgb, str):
            return fg.rgb
    except Exception:
        pass
    try:
        if fg.type == "theme":
            return "theme%d_%s" % (fg.theme, fg.tint)
    except Exception:
        pass
    return "other"


def find_header_columns(ws):
    """Locate columns by header text.

    Product columns = row2 product names (PPH: '9-1 T5'; SH N/S: 'T1'..'T11').
    N/S drum size column = row1 'Drum Size' (first), drum count = 'QTY 1'
    (first). EW* products and 'QTY 2' are skipped automatically.
    """
    key_cols = {}                                    # product name -> column
    for c in range(1, ws.max_column + 1):
        h2 = ws.cell(row=2, column=c).value
        if not isinstance(h2, str):
            continue
        k = h2.strip()
        if re.match(r"^\d+ ?-1 T\d+$", k) or re.match(r"^T\d+$", k):
            key_cols[k] = c
    drum_col = qty_col = None
    for c in range(1, ws.max_column + 1):
        h1 = ws.cell(row=1, column=c).value
        if h1 == "Drum Size" and drum_col is None:
            drum_col = c
        if h1 == "QTY 1" and qty_col is None:
            qty_col = c
    return key_cols, drum_col, qty_col


def read_workbook(xlsx_path, sheet_name, gw_sheet):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        sys.exit("sheet '%s' not found, available: %s" % (sheet_name, wb.sheetnames))
    if gw_sheet not in wb.sheetnames:
        sys.exit("weight sheet '%s' not found, available: %s" % (gw_sheet, wb.sheetnames))
    ws = wb[sheet_name]
    validate_packing_list_fields(ws)

    key_cols, drum_col, qty_col = find_header_columns(ws)
    if not key_cols or drum_col is None or qty_col is None:
        sys.exit("header detection failed: product/drum size/QTY 1 not found")

    # products & part numbers: row4 = RED part, row5 = BLK part (PPH only)
    products = []            # (short name, pn_red, pn_blk, kind)
    for key, c in key_cols.items():
        pr = ws.cell(row=4, column=c).value
        pb = ws.cell(row=5, column=c).value
        if pr is None:
            sys.exit("row4 part number missing for product '%s'" % key)
        products.append((key, str(pr),
                         str(pb) if pb else None,
                         "PPH" if not key.startswith("T") else "SH"))

    # weight sheet: col A = category (PPH/SH), col B = product name, C = net kg
    gw = wb[gw_sheet]
    validate_gw_sheet(ws, gw)
    unit_weight = {}
    for r in range(1, gw.max_row + 1):
        key = gw.cell(row=r, column=2).value
        nw = gw.cell(row=r, column=3).value
        if isinstance(key, str) and isinstance(nw, (int, float)):
            unit_weight[key.strip()] = float(nw)

    # combiner box rows
    boxes = []
    for r in range(1, ws.max_row + 1):
        name = ws.cell(row=r, column=1).value
        if not (isinstance(name, str) and re.match(r"^CMB\.", name)):
            continue
        ar = ws.cell(row=r, column=qty_col).value
        drum = ws.cell(row=r, column=drum_col).value
        if not ar or not drum:
            sys.exit("row %d (%s): drum size / QTY 1 missing" % (r, name))
        ar = int(ar)
        # column B: the box type the pallets are grouped by (rule 2)
        box_type = norm_text(ws.cell(row=r, column=2).value)
        if not box_type:
            sys.exit("row %d (%s): box type (column B) is empty" % (r, name))
        raw_drum = str(drum)
        # the size cell may list several sizes one per line (mixed-size box)
        sizes = [s.strip() for s in raw_drum.split("\n") if s.strip()] \
            if "\n" in raw_drum else [raw_drum]
        if len(sizes) == 1:
            drum_sizes = [sizes[0]] * ar
        else:
            base, rem = divmod(ar, len(sizes))
            counts = [base + 1] * rem + [base] * (len(sizes) - rem)
            drum_sizes = []
            for sz, cnt in zip(sizes, counts):
                drum_sizes += [sz] * cnt
        for sz in set(drum_sizes):
            if sz not in CONFIG["pallet_sizes"]:
                sys.exit("row %d (%s): unknown drum size %s"
                         "(add it to CONFIG[pallet_sizes])" % (r, name, sz))
        # items per drum: even split when divisible; if not, the first
        # remainder drums carry one extra item (totals stay exact)
        items_per_drum = [OrderedDict() for _ in range(ar)]
        splittable = True
        for key, pn_r, pn_b, typ in products:
            v = ws.cell(row=r, column=key_cols[key]).value
            if v is None:
                continue
            v = int(round(float(v)))
            if v == 0:
                continue
            base, rem = divmod(v, ar)
            splittable = splittable and rem == 0
            for i in range(ar):
                if base or (i < rem):
                    items_per_drum[i][key] = base + (1 if i < rem else 0)
        if not any(items_per_drum):
            continue                                 # no N/S products, skip
        boxes.append({
            "row": r, "name": name,
            "box_type": box_type,
            "axles": ar,
            "drum_sizes": drum_sizes,
            "items_per_drum": items_per_drum,
            "splittable": splittable,
            "container": cell_fill_key(ws.cell(row=r, column=qty_col)),
        })
    return boxes, products, unit_weight


# ----------------------------------------------------------------------------
# Pallet packing
# ----------------------------------------------------------------------------
def pack(entries, pallet_main, cap, pallet_rem, cap_rem):
    """Pack the drums of one box-type run in sheet row order.

    `entries` is a list of (box, first_drum_index, drum_count) covering one
    maximal run of same-type boxes of a (container, drum size) group, in
    sheet row order. A single greedy queue: drums of splittable boxes fill
    the open pallet, leftovers chain to the next box. A non-splittable box
    is an indivisible block: its drums join the open pallet only if they all
    fit at once, otherwise they get their own pallet(s) in order, while the
    queue keeps chaining (same behavior as the reference PDF). Only the
    final remainder of a run may use the low 650-sized pallet; the leftovers
    of the group's runs are combined later by merge_leftovers().
    Returns [ {"dims": size, "chunks": [(box, first_drum, count)] } ].
    """
    pallets, cur, cur_n = [], [], 0

    def flush_pallet(as_main=True):
        nonlocal cur, cur_n
        if cur_n:
            pallets.append({"dims": pallet_main if as_main else pallet_rem,
                            "chunks": list(cur)})
            cur, cur_n = [], 0

    def take(box, start, n):
        nonlocal cur, cur_n
        if cur and cur[-1][0] is box and cur[-1][1] + cur[-1][2] == start:
            cur[-1][2] += n                       # contiguous, merge
        else:
            cur.append([box, start, n])
        cur_n += n

    for box, start, n in entries:
        if box["splittable"]:
            off, left = start, n
            while left > 0:
                m = min(left, cap - cur_n)
                take(box, off, m)
                off += m
                left -= m
                if cur_n == cap:
                    flush_pallet()
        else:
            if n <= cap - cur_n:                   # fits as one block
                take(box, start, n)
                if cur_n == cap:
                    flush_pallet()
            else:                                  # own pallet(s), queue untouched
                off, left = start, n
                while left > 0:
                    m = min(left, cap)
                    pallets.append({"dims": pallet_main,
                                    "chunks": [(box, off, m)]})
                    off += m
                    left -= m
    if cur_n:
        flush_pallet(as_main=(cur_n > cap_rem))
    return pallets


def split_into_runs(entries):
    """Split one group's entries into maximal runs of the same box type.

    `entries` is a list of (box, first_drum_index, drum_count) in sheet row
    order. A box has exactly one type, so a run ends where the type changes.
    """
    runs = []
    for entry in entries:
        if runs and runs[-1][0][0]["box_type"] == entry[0]["box_type"]:
            runs[-1].append(entry)
        else:
            runs.append([entry])
    return runs


def drum_count(pallet):
    """Number of drums loaded on a pallet."""
    return sum(count for _, _, count in pallet["chunks"])


def merge_leftovers(partials, pallet_main, cap, pallet_rem, cap_rem):
    """Combine the leftover pallets of one (container, drum size) group.

    Every box-type run leaves at most one partial pallet behind. Two
    partials share a pallet as long as their combined drum count still fits
    the capacity (first fit, in run order); the merged pallet is anchored on
    the earlier box so it keeps its place in sheet order. A merged pallet is
    re-sized for its final drum count (taller than the low leftover pallet
    -> main size), while a leftover that stays alone keeps the size pack()
    gave it.
    Returns [ {"dims": size, "chunks": [...], "anchor": row} ].
    """
    merged = []
    for part in partials:
        for target in merged:
            if not target["open"]:
                continue
            if drum_count(target) + drum_count(part) > cap:
                continue
            target["chunks"].extend(part["chunks"])
            target["anchor"] = min(target["anchor"], part["anchor"])
            target["open"] = drum_count(target) < cap
            target["dims"] = (pallet_main if drum_count(target) > cap_rem
                              else pallet_rem)
            break
        else:
            merged.append({"dims": part["dims"],
                           "chunks": list(part["chunks"]),
                           "anchor": part["anchor"],
                           "open": drum_count(part) < cap})
    return merged


def build_pallets(boxes, pallet_sizes):
    """Group drums by (container color, drum size); group order = first
    appearance. Inside a group the drums are packed per box type: a maximal
    run of consecutive same-type boxes is packed on its own, then the
    leftover partials of the group's runs are combined. A box with several
    drum sizes contributes to each group. Within a group, pallets are
    ordered by the sheet row of their first box.
    """
    groups = OrderedDict()            # (container, size) -> [(box, start, n)]
    for b in boxes:
        i = 0
        while i < len(b["drum_sizes"]):
            sz = b["drum_sizes"][i]
            j = i
            while j < len(b["drum_sizes"]) and b["drum_sizes"][j] == sz:
                j += 1
            groups.setdefault((b["container"], sz), []).append((b, i, j - i))
            i = j
    pallets = []
    for (container, drum), entries in groups.items():
        pallet_main, cap, pallet_rem, cap_rem = pallet_sizes[drum]
        full, partials = [], []
        for run in split_into_runs(entries):
            for p in pack(run, pallet_main, cap, pallet_rem, cap_rem):
                p["anchor"] = min(b["row"] for b, _, _ in p["chunks"])
                (partials if drum_count(p) < cap else full).append(p)
        grouped = full + merge_leftovers(partials, pallet_main, cap,
                                         pallet_rem, cap_rem)
        grouped.sort(key=lambda p: p["anchor"])
        pallets.extend({"container": container, "drum": drum, **p}
                       for p in grouped)
    return pallets


def short_key_of(pn, products):
    """Part number -> product short name (unit weight & description lookup)."""
    for key, pn_r, pn_b, typ in products:
        if pn == pn_r or pn == pn_b:
            return key
    return pn


def pallet_items(chunks, products):
    """Pallet item list: sum the per-drum item counts for the drums on it.

    Lines follow the product column order of the packing list (D..AM), not
    the order the boxes were added to the pallet, so a pallet combining two
    boxes always lists the parts the same way. A PPH drum carries one RED
    and one BLK item of each kind (same qty), listed as RED then BLK.
    """
    qty_by_key = OrderedDict()                       # product key -> quantity
    for box, off, n in chunks:
        for per in box["items_per_drum"][off:off + n]:
            for key, qty in per.items():
                qty_by_key[key] = qty_by_key.get(key, 0) + qty
    merged = OrderedDict()                           # part number -> quantity
    for key, pn_r, pn_b, typ in products:            # column order
        qty = qty_by_key.get(key, 0)
        if not qty:
            continue
        merged[pn_r] = merged.get(pn_r, 0) + qty
        if typ == "PPH" and pn_b:
            merged[pn_b] = merged.get(pn_b, 0) + qty
    items = []
    for pn, qty in merged.items():
        desc = CONFIG["descriptions"].get(pn) or describe_part(
            pn, short_key_of(pn, products))
        items.append({"pn": pn, "desc": desc, "qty": qty})
    return items


# ----------------------------------------------------------------------------
# Excel mark output: one sheet per pallet, layout mirrors the reference PDF
# (labels in column A, values in column B, item table A/B/C, photo box)
# ----------------------------------------------------------------------------
def _est_desc_height(desc, chars_per_line=50):
    """Rows needed to show a wrapped description in the B column."""
    lines = max(1, (len(desc) + chars_per_line - 1) // chars_per_line)
    return lines


def write_mark_sheet(ws, cfg, lot, total, items, weight, dims, blocks):
    """Fill one worksheet with a single pallet's mark (like one PDF page)."""
    thin = Side(style="thin", color="00000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    font = Font(name="Arial", size=10)
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 48
    ws.column_dimensions["C"].width = 13

    r = 1
    head = [
        ("SUPPLIER ID:", cfg["supplier_id"]),
        ("COUNTRY OF ORIGIN:", cfg["country_of_origin"]),
        ("PROJECT NUMBER", cfg["project_number"]),
        ("PROJECT NAME", cfg["project_name"]),
        ("SO NUMBER", cfg["so_number"]),
        ("PO NUMBER", cfg["po_number"]),
        ("LOT NUMBER", "%d-%d" % (lot, total)),
    ]
    for label, value in head:
        a = ws.cell(row=r, column=1, value=label)
        b = ws.cell(row=r, column=2, value=value)
        a.font = font
        b.font = font
        ws.row_dimensions[r].height = 20
        r += 1

    # item table header
    for col, text in ((1, "PART NUMBER"), (2, "PART DESCRIPTION"),
                      (3, "QUANTITY/PN")):
        cell = ws.cell(row=r, column=col, value=text)
        cell.font = font
    ws.row_dimensions[r].height = 20
    r += 1

    for item in items:
        a = ws.cell(row=r, column=1, value=item["pn"])
        b = ws.cell(row=r, column=2, value=item["desc"])
        c = ws.cell(row=r, column=3, value=item["qty"])
        for cell in (a, b, c):
            cell.font = font
        b.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 13 * _est_desc_height(item["desc"]) + 6
        r += 1

    r += 1
    tail = [
        ("WEIGHT (IN KG)", fmt_num(weight)),
        ("DIMENSIONS (MM) LxWxH", "%d x %d x %d mm" % tuple(dims)),
        ("TYPE OF STORAGE", cfg["type_of_storage"]),
        ("STACKABILITY", cfg["stackability"]),
        ("BLOCK NUMBER", blocks),
        ("PART PHOTO", ""),
    ]
    for label, value in tail:
        a = ws.cell(row=r, column=1, value=label)
        b = ws.cell(row=r, column=2, value=value)
        a.font = font
        b.font = font
        ws.row_dimensions[r].height = 20
        r += 1

    # photo placeholder box (bordered merged cells)
    for i in range(6):
        for col in (2, 3):
            cell = ws.cell(row=r + i, column=col)
            cell.border = border
        ws.row_dimensions[r + i].height = 22
    pc = ws.cell(row=r, column=2, value="PHOTO")
    pc.font = Font(name="Arial", size=9)
    pc.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=r, start_column=2, end_row=r + 5, end_column=3)

    ws.print_area = "A1:C%d" % (r + 5)               # printable mark area
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = 9                      # A4
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A1"


def save_workbook(wb, path):
    """Save the workbook without ever leaving a half-written file behind.

    openpyxl truncates the target as soon as it opens it, so an error in the
    middle of saving used to leave an unreadable file under the requested
    name. Write a sibling temp file first and move it into place only once
    the workbook is complete; on failure drop the temp file and leave
    whatever was at `path` untouched.
    """
    tmp_path = path + ".part.xlsx"
    try:
        wb.save(tmp_path)
        os.replace(tmp_path, path)
    except BaseException:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise


def make_workbook(path, pallets, cfg):
    if not pallets:
        sys.exit("nothing to write: no pallet could be built from this sheet")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)                             # drop default empty sheet
    total = len(pallets)
    for i, p in enumerate(pallets, 1):
        blocks = "\uff0c".join(
            OrderedDict((b["name"], None) for b, _, _ in p["chunks"]))
        ws = wb.create_sheet(title="LOT %d-%d" % (i, total))
        write_mark_sheet(ws, cfg, i, total, p["items"], p["weight"],
                         [int(x) for x in p["dims"].split("*")], blocks)
    save_workbook(wb, path)


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------
def sanitize_out_name(raw):
    """Strictly validate the output file name (a plain file name, no folders)."""
    if not isinstance(raw, str) or not raw.strip():
        sys.exit("output file name is empty")
    name = raw.strip()
    if os.sep in name or "/" in name or "\\" in name:
        sys.exit("output file name must not contain folders: %s" % raw)
    if name in (".", "..") or ".." in name:
        sys.exit("invalid output file name: %s" % raw)
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    return os.path.basename(name)


def main():
    ap = argparse.ArgumentParser(description="shipping marks workbook (N/S only)")
    ap.add_argument("excel", help="input Excel (Shipping details ...)")
    ap.add_argument("--sheet", default="Block 5-8 Packing List",
                    help="packing list sheet name")
    ap.add_argument("--gw-sheet", default="Batch_01_GW",
                    help="net weight sheet name")
    ap.add_argument("--out", default="Marks.xlsx",
                    help="output workbook file name (no folders)")
    args = ap.parse_args()

    xlsx_path = resolve_input(args.excel)
    cfg = dict(CONFIG)

    boxes, products, unit_weight = read_workbook(
        xlsx_path, args.sheet, args.gw_sheet)
    pallets = build_pallets(boxes, cfg["pallet_sizes"])

    for p in pallets:
        p["items"] = pallet_items(p["chunks"], products)
        axles = sum(n for _, _, n in p["chunks"])
        w = cfg["pallet_weight_kg"] + \
            cfg["axle_weight_kg"][p["drum"].split("*")[0]] * axles
        for it in p["items"]:
            key = short_key_of(it["pn"], products)
            # a PPH product = one RED + one BLK unit pair; weight is counted
            # once (the pair), the BLK line adds no extra mass
            for k2, pn_r2, pn_b2, typ2 in products:
                if it["pn"] == pn_b2 and typ2 == "PPH":
                    break
            else:
                w += unit_weight[key] * it["qty"]
        p["weight"] = round(w, 2)

    out_name = sanitize_out_name(args.out)
    make_workbook(out_name, pallets, cfg)

    for i, p in enumerate(pallets, 1):
        blocks = "\uff0c".join(
            OrderedDict((b["name"], None) for b, _, _ in p["chunks"]))
        dims = p["dims"].replace("*", "x")
        print("LOT %d | %s | %s | %s kg" % (i, blocks, dims,
                                            fmt_num(p["weight"])))
    print("pallets  : %d" % len(pallets))
    print("mark XLSX: %s" % out_name)


if __name__ == "__main__":
    main()
