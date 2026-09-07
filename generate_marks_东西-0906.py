#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Shipping mark generator -- East-West (EW) direction, Excel output
=================================================================

Reads the "Shipping details ... xlsx" workbook and produces a shipping-mark
workbook: one sheet per pallet (same layout as generate_marks_南北-0906.py:
supplier / project / LOT info, part list with quantities, weight,
dimensions, storage info, block numbers and a photo placeholder).

Business rules (East-West only, per the 需求描述-东西.txt memo)
1. The EW goods of each combiner box are the EW1..EW7 columns (row2,
   AG..AM); part numbers are row4/row5 AG..AM (EW1 = SH-2.8.RED.BLK.00013,
   EW2 = SH-2.8.RED.BLK.00014, ...).
2. Column "Drum Size" (SECOND, i.e. AT) is the EW drum size; "QTY 2"
   (AU) is the number of EW drums. AT/AU cells are MERGED across several
   rows: the merged block is ONE axle group carrying all boxes of those
   rows. E.g. AT6:AT7 = CMB.005.4.01 + CMB.005.4.02 on one EW axle;
   a non-merged cell (e.g. row 44) is a single-box group.
3. Packing per drum size, in sheet row order (no container-color rule,
   the memo mixes colors freely):
   - 750*350*500 -> pallet 1500*750*1150 holds 4 axles (2 side by side,
     2 layers); left over <=2 axles -> 1500*750*650
   - 1000*350*500 -> pallet 1050*1050*1150 holds 2 axles stacked;
     1 left over -> 1050*1050*650
4. Pallet weight = pallet 25 kg + drum weight (1000 = 28 kg, 750 = 25 kg)
   * axle count + net weight of all EW harnesses on the pallet
   (unit weights from the GW sheet rows EW1..EW7, column C).
   Example: CMB.008.2.17 (EW1=4, EW2=8) = 4*1.23+8*3.27 = 31.08 kg;
   pallet of CMB.008.2.17+CMB.008.2.18 = 25 + 1*25 + 31.08 + 31.08
   = 112.16 kg on a 1500*750*650 pallet.
5. Output: Marks-EW.xlsx, one sheet per pallet ("LOT n-total").

Reusability
- Column headers are FIXED (row1 group names, row2 EW1..EW7, row4/row5
  part numbers); rows and values may change freely. Columns are located
  by header text, not letters.
- Shipment header info (supplier, project, SO/PO...) is in CONFIG below.
- Part-number -> description mapping: built-in template (SH);
  override any of them in CONFIG["descriptions"].

Usage (run from the folder that holds the Excel):
    python3 generate_marks_东西-0906.py "Shipping details ... xlsx"
    python3 generate_marks_东西-0906.py xxx.xlsx --sheet "Block 5-8 Packing List" \
        --gw-sheet Batch_01_GW --out Marks-EW.xlsx
Options:
    --sheet    packing-list sheet name (default "Block 5-8 Packing List")
    --gw-sheet net weight sheet name  (default "Batch_01_GW")
    --out      output workbook file name (no folders, default "Marks-EW.xlsx")
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
    # drum size -> (main pallet size, axles per pallet,
    #               leftover pallet size, max axles on leftover pallet)
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
# Part description templates
# ----------------------------------------------------------------------------
def describe_part(pn, product_key):
    """Build the PART DESCRIPTION for EW SH part numbers."""
    if pn.startswith("SH-"):
        # SH-2.8.RED.BLK.00013 + "EW1" -> MRTSN,WCHTA,SH,E/W,PAIR 1,#8CU-#6AL
        m = re.match(r"^EW(\d+)$", product_key.strip())
        if m:
            n = int(m.group(1))
            desc = "MRTSN,WCHTA,SH,E/W,PAIR %d,#8CU-#6AL" % n
            if n >= 10:                  # same rule as the N/S template
                desc += ",XTRACKER"
            return desc
    return pn


def fmt_num(x):
    """Trim trailing zeros: 112.16 / 112.1 / 112"""
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
    """
    def expect(row, col, want):
        got = norm_text(ws.cell(row=row, column=col).value)
        if got != want:
            sys.exit("packing list header error: %s%d must be %r, got %r"
                     % (openpyxl.utils.get_column_letter(col), row, want, got))

    expect(1, 1, "COMBINER BOX NO.")
    expect(1, 2, "BOX TYPE")
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


def find_east_west_columns(ws):
    """Locate the EW product columns and the EW drum size/count columns.

    Product columns = row2 names matching 'EW1'..'EW7' (AG..AM).
    The EW drum size column = the SECOND row1 'Drum Size' (AT), the EW
    drum count = row1 'QTY 2' (AU). The first Drum Size / QTY 1 belong
    to the North-South direction and are ignored here.
    """
    key_cols = {}                                    # e.g. 'EW1' -> column
    for c in range(1, ws.max_column + 1):
        h2 = ws.cell(row=2, column=c).value
        if isinstance(h2, str) and re.match(r"^EW\d+$", h2.strip()):
            key_cols[h2.strip()] = c
    drum_cols = []
    qty1_col = qty2_col = None
    for c in range(1, ws.max_column + 1):
        h1 = ws.cell(row=1, column=c).value
        if h1 == "Drum Size":
            drum_cols.append(c)
        if h1 == "QTY 1" and qty1_col is None:
            qty1_col = c
        if h1 == "QTY 2" and qty2_col is None:
            qty2_col = c
    if len(drum_cols) < 2 or qty2_col is None:
        sys.exit("header detection failed: EW Drum Size / QTY 2 not found")
    return key_cols, drum_cols[1], qty2_col


def read_workbook(xlsx_path, sheet_name, gw_sheet):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        sys.exit("sheet '%s' not found, available: %s" % (sheet_name, wb.sheetnames))
    if gw_sheet not in wb.sheetnames:
        sys.exit("weight sheet '%s' not found, available: %s" % (gw_sheet, wb.sheetnames))
    ws = wb[sheet_name]
    validate_packing_list_fields(ws)

    key_cols, at_col, au_col = find_east_west_columns(ws)
    if not key_cols:
        sys.exit("header detection failed: no EW1..EW7 product columns")

    # products & part numbers: row4 (RED/BLK combined code for SH lines)
    products = []            # (short name, part number, kind)
    for key, c in key_cols.items():
        pr = ws.cell(row=4, column=c).value
        if pr is None:
            sys.exit("row4 part number missing for product '%s'" % key)
        products.append((key, str(pr)))

    # weight sheet: col A = category (PPH/SH), col B = product name, C = net kg
    gw = wb[gw_sheet]
    validate_gw_sheet(ws, gw)
    unit_weight = {}
    for r in range(1, gw.max_row + 1):
        key = gw.cell(row=r, column=2).value
        nw = gw.cell(row=r, column=3).value
        if isinstance(key, str) and isinstance(nw, (int, float)):
            unit_weight[key.strip()] = float(nw)

    # ---- merged AT/AU ranges -> effective (drum size, axle count) per row ----
    drum_merge = {}          # row -> (size, axle count)
    for rng in ws.merged_cells.ranges:
        if rng.min_col == at_col:
            size = ws.cell(row=rng.min_row, column=rng.min_col).value
            count = ws.cell(row=rng.min_row, column=rng.min_col + 1).value
            if isinstance(size, str) and isinstance(count, (int, float)):
                for rr in range(rng.min_row, rng.max_row + 1):
                    drum_merge[rr] = (size.strip(), int(count))

    def eff_at_au(row):
        if row in drum_merge:
            return drum_merge[row]
        size = ws.cell(row=row, column=at_col).value
        count = ws.cell(row=row, column=au_col).value
        if isinstance(size, str) and count is not None:
            return (size.strip(), int(count))
        return None

    # ---- groups: consecutive rows sharing the same EW axle (merged or solo) --
    groups = []
    for r in range(1, ws.max_row + 1):
        name = ws.cell(row=r, column=1).value
        if not (isinstance(name, str) and re.match(r"^CMB\.", name)):
            continue
        info = eff_at_au(r)
        box_ew = OrderedDict()
        for key, pn_r in products:
            v = ws.cell(row=r, column=key_cols[key]).value
            if isinstance(v, (int, float)) and v:
                box_ew[key] = int(round(v))
        if info is None:
            if not box_ew:
                continue                              # no EW goods, skip
            sys.exit("row %d (%s): EW goods without a QTY 2 / drum-size cell"
                     "(missing merge?)" % (r, name))
        size, count = info
        if size not in CONFIG["pallet_sizes"]:
            sys.exit("row %d (%s): unknown EW drum size %s"
                     "(add it to CONFIG[pallet_sizes])" % (r, name, size))
        if not box_ew:
            continue                                  # no EW goods, skip
        # a group starts at a merged-range top-left cell or a solo cell
        is_group_start = ws.cell(row=r, column=at_col).value is not None
        if groups and not is_group_start and \
                groups[-1]["size"] == size and \
                groups[-1]["count"] == count and \
                groups[-1]["end"] == r - 1:
            groups[-1]["boxes"].append({"row": r, "name": name, "ew": box_ew})
            groups[-1]["end"] = r
        else:
            groups.append({"size": size, "count": count,
                           "start": r, "end": r,
                           "boxes": [{"row": r, "name": name, "ew": box_ew}]})
    return groups, products, unit_weight


# ----------------------------------------------------------------------------
# Axle groups -> per-axle item lists -> pallet packing
# ----------------------------------------------------------------------------
def group_axle_items(group):
    """Items carried by each EW axle of one axle group.

    Each box's harness count is divided evenly over the group's axle
    count (first `rem` axles get one extra item); all boxes of the group
    sit on the same axles, so the item lists are summed per axle.
    """
    n = group["count"]
    lists = [OrderedDict() for _ in range(n)]
    for box in group["boxes"]:
        for key, v in box["ew"].items():
            base, rem = divmod(v, n)
            for i in range(n):
                q = base + (1 if i < rem else 0)
                if q:
                    lists[i][key] = lists[i].get(key, 0) + q
    return lists


def pack(entries, pallet_main, cap, pallet_rem, cap_rem):
    """Pack axle groups of one drum size in sheet row order, 4-per-slot
    style: groups fill the current pallet; a group that does not fit gets
    its own pallet(s) and the queue keeps chaining; only the final
    remainder of the size group may use the low 650-sized pallet.
    `entries` = [(group, first_axle, axle_count), ...].
    Returns [{"dims": size, "chunks": [(group, off, n)]}].
    """
    pallets, cur, cur_n = [], [], 0

    def flush_pallet(as_main=True):
        nonlocal cur, cur_n
        if cur_n:
            pallets.append({"dims": pallet_main if as_main else pallet_rem,
                            "chunks": list(cur)})
            cur, cur_n = [], 0

    def take(group, off, n):
        nonlocal cur, cur_n
        if cur and cur[-1][0] is group and cur[-1][1] + cur[-1][2] == off:
            cur[-1][2] += n
        else:
            cur.append([group, off, n])
        cur_n += n

    for group, off, n in entries:
        if n <= cap - cur_n:
            take(group, off, n)
            if cur_n == cap:
                flush_pallet()
        else:
            left, so = n, off
            while left > 0:                       # own pallet(s), queue untouched
                m = min(left, cap)
                pallets.append({"dims": pallet_main, "chunks": [(group, so, m)]})
                so += m
                left -= m
    if cur_n:
        flush_pallet(as_main=(cur_n > cap_rem))
    return pallets


def build_pallets(groups, pallet_sizes):
    """Group axle groups by drum size (first appearance order) and pack."""
    by_size = OrderedDict()
    for g in groups:
        by_size.setdefault(g["size"], []).append(g)
    pallets = []
    for size, glist in by_size.items():
        pallet_main, cap, pallet_rem, cap_rem = pallet_sizes[size]
        entries = [(g, 0, g["count"]) for g in glist]
        grouped = []
        for p in pack(entries, pallet_main, cap, pallet_rem, cap_rem):
            p["anchor"] = min(g["start"] for g, _, _ in p["chunks"])
            grouped.append({"size": size, **p})
        grouped.sort(key=lambda p: p["anchor"])
        pallets.extend(grouped)
    return pallets


def short_key_of(pn, products):
    """Part number -> product short name (unit weight & description lookup)."""
    for key, pn_r in products:
        if pn == pn_r:
            return key
    return pn


def pallet_items(chunks, products):
    """Pallet item list: sum the EW harness counts of its axles."""
    merged = OrderedDict()                           # part number -> quantity
    for group, off, n in chunks:
        for axle in group_axle_items(group)[off:off + n]:
            for key, qty in axle.items():
                for k2, pn_r in products:
                    if k2 == key:
                        merged[pn_r] = merged.get(pn_r, 0) + qty
                        break
    items = []
    for pn, qty in merged.items():
        desc = CONFIG["descriptions"].get(pn) or describe_part(
            pn, short_key_of(pn, products))
        items.append({"pn": pn, "desc": desc, "qty": qty})
    return items


# ----------------------------------------------------------------------------
# Excel mark output (same layout style as the N/S generator)
# ----------------------------------------------------------------------------
def _est_desc_height(desc, chars_per_line=50):
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

    for i in range(6):
        for col in (2, 3):
            cell = ws.cell(row=r + i, column=col)
            cell.border = border
        ws.row_dimensions[r + i].height = 22
    pc = ws.cell(row=r, column=2, value="PHOTO")
    pc.font = Font(name="Arial", size=9)
    pc.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=r, start_column=2, end_row=r + 5, end_column=3)

    ws.print_area = "A1:C%d" % (r + 5)
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = 9                      # A4
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A1"


def make_workbook(path, pallets, cfg):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    total = len(pallets)
    for i, p in enumerate(pallets, 1):
        blocks = "\uff0c".join(
            OrderedDict((b["name"], None) for g, _, _ in p["chunks"]
                        for b in g["boxes"]))
        ws = wb.create_sheet(title="LOT %d-%d" % (i, total))
        write_mark_sheet(ws, cfg, i, total, p["items"], p["weight"],
                         [int(x) for x in p["dims"].split("*")], blocks)
    wb.save(path)


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
    ap = argparse.ArgumentParser(description="shipping marks workbook (E/W only)")
    ap.add_argument("excel", help="input Excel (Shipping details ...)")
    ap.add_argument("--sheet", default="Block 5-8 Packing List",
                    help="packing list sheet name")
    ap.add_argument("--gw-sheet", default="Batch_01_GW",
                    help="net weight sheet name")
    ap.add_argument("--out", default="Marks-EW.xlsx",
                    help="output workbook file name (no folders)")
    args = ap.parse_args()

    xlsx_path = resolve_input(args.excel)
    cfg = dict(CONFIG)

    groups, products, unit_weight = read_workbook(
        xlsx_path, args.sheet, args.gw_sheet)
    pallets = build_pallets(groups, cfg["pallet_sizes"])

    for p in pallets:
        p["items"] = pallet_items(p["chunks"], products)
        axles = sum(n for _, _, n in p["chunks"])
        w = cfg["pallet_weight_kg"] + \
            cfg["axle_weight_kg"][p["size"].split("*")[0]] * axles
        for it in p["items"]:
            w += unit_weight[short_key_of(it["pn"], products)] * it["qty"]
        p["weight"] = round(w, 2)

    out_name = sanitize_out_name(args.out)
    make_workbook(out_name, pallets, cfg)

    for i, p in enumerate(pallets, 1):
        blocks = "\uff0c".join(
            OrderedDict((b["name"], None) for g, _, _ in p["chunks"]
                        for b in g["boxes"]))
        dims = p["dims"].replace("*", "x")
        print("LOT %d | %s | %s | %s kg" % (i, blocks, dims,
                                            fmt_num(p["weight"])))
    print("pallets  : %d" % len(pallets))
    print("mark XLSX: %s" % out_name)


if __name__ == "__main__":
    main()
