"""Check the recalculated workbook: no formula errors, and Scoring matches the Python reference."""
import json
import sys

import openpyxl

wbp, refp, colp = sys.argv[1:4]
ref = json.load(open(refp))
cm = json.load(open(colp))
L = cm["scoring"]

wb_v = openpyxl.load_workbook(wbp, data_only=True)
wb_f = openpyxl.load_workbook(wbp)

# 1) formula errors / uncomputed formulas
errs, empty_formulas, n_formulas = [], 0, 0
for ws in wb_f.worksheets:
    wv = wb_v[ws.title]
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                n_formulas += 1
                v = wv[c.coordinate].value
                if isinstance(v, str) and v.startswith("#"):
                    errs.append((ws.title, c.coordinate, v))
print("Formulas:", n_formulas, "| error cells:", len(errs))
for e in errs[:20]:
    print("  ", e)

# 2) Scoring vs reference
sc = wb_v["Scoring"]
mism = []
for r in range(cm["first"], cm["last"] + 1):
    mid = sc[f"{L['id']}{r}"].value
    o = ref[mid]
    got = dict(flag=sc[f"{L['flag']}{r}"].value, main=sc[f"{L['main']}{r}"].value or "",
               rank=sc[f"{L['rank']}{r}"].value, track=sc[f"{L['track']}{r}"].value,
               sources=sc[f"{L['sources']}{r}"].value, tag=sc[f"{L['tag']}{r}"].value,
               value=sc[f"{L['value']}{r}"].value)
    exp = dict(flag=o["flagged"], main=o["main"], rank=o["rank"], track=o["track"], sources=o["sources"],
               tag=o["tag"], value=o["value"])
    for k in exp:
        a, b = got[k], exp[k]
        same = abs(a - b) < 1e-6 if isinstance(b, float) else a == b
        if not same:
            mism.append((mid, k, a, b))
print("Scoring vs reference mismatches:", len(mism))
for m in mism[:20]:
    print("  ", m)

# 3) Lists and summary
def dump(ws_name, first_row, cols, n):
    ws = wb_v[ws_name]
    out = []
    for r in range(first_row, first_row + n):
        vals = [ws.cell(r, c).value for c in cols]
        if vals[0] in (None, ""):
            continue
        out.append(vals)
    return out

print("\nCSM top 15 tab:")
for row in dump("CSM top 15", 6, [2, 4, 9, 11, 13, 15, 16], 40):
    print("  ", row)
print("count cell:", wb_v["CSM top 15"]["C3"].value)
print("\nLight-touch tab:")
for row in dump("Light-touch", 6, [2, 4, 9, 11, 13, 17], 40):
    print("  ", row)
print("count cell:", wb_v["Light-touch"]["C3"].value)
print("\nRanked list first 3 / last 2:")
rl = dump("Ranked list", 2, [1, 3, 6, 7, 9, 11, 16], 100)
for row in rl[:3] + rl[-2:]:
    print("  ", row)
print("\nSummary:")
sm = wb_v["Summary"]
for r in range(4, sm.max_row + 1):
    a, b = sm.cell(r, 1).value, sm.cell(r, 2).value
    if a is not None:
        print(f"   {a}: {b}" + ("" if sm.cell(r, 3).value is None else f" | {sm.cell(r,3).value}") +
              ("" if sm.cell(r, 4).value is None else f" | {sm.cell(r,4).value} | {sm.cell(r,5).value}"))
print("\nTest summary:")
ts = wb_v["Test vs labels"]
for r in range(3, 10):
    print("  ", ts.cell(r, 1).value, "->", ts.cell(r, 5).value)
print("\nTest rows (merchant, you, system, your drv, sys drv, your pb, sys pb, track):")
for r in range(11, 31):
    print("  ", [ts.cell(r, c).value for c in (1, 2, 8, 4, 9, 6, 10, 11)])
print("\nSample reasons:")
for r in range(2, 102):
    mid = sc[f"{L['id']}{r}"].value
    if mid in ("C2004", "C2088", "C2008", "C2061", "C2012", "C2002", "C2003", "C2091"):
        print(f"   {mid}: {sc[L['reason'] + str(r)].value}")
print("\nNote tags counts sum:", sum(wb_v["Note tags"].cell(r, 3).value or 0 for r in range(4, 26)))
