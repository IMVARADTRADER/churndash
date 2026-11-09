"""Re-run the fact check on the CSM drafts as they are now in the workbook.

Usage (from D:\\Master CV\\Razorpay):
  python -I scripts/recheck_drafts.py P2_Churn_Scoring.xlsx           check only: prints results, changes nothing
  python -I scripts/recheck_drafts.py P2_Churn_Scoring.xlsx --write   also updates the Messages tab and scripts/*.json

What --write does:
  - Messages tab: refreshes "Numbers used", "Fact check", "Words" and "Drafted by" for every draft,
    and the note in A2. No other cell is touched, so Excel's calculated values stay as they are.
  - scripts/messages.json gets your edited subjects, bodies and CSM notes, so a rebuild keeps them.
    scripts/messages_claude_original.json holds Claude's own wording (first drafts plus Claude's 9 Nov revisions
    of C2018 and C2062); a draft that differs from it is labelled "edited by you".
    Claude's untouched first drafts are kept in scripts/messages_claude_first_drafts.json.
  - scripts/reference.json and scripts/factcheck.json are rewritten (the rebuild reads them).
  - A backup of the workbook goes to backups/ first.
Close the workbook in Excel before using --write.
"""
import datetime as dt
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from xml.sax.saxutils import escape

import openpyxl

sys.dont_write_bytecode = True  # keep scripts/ free of __pycache__
HERE = os.path.dirname(os.path.abspath(__file__))
PROJ = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import reference as R  # noqa: E402

wb_path = os.path.abspath(sys.argv[1])
WRITE = "--write" in sys.argv[2:]
SRC = os.path.join(PROJ, "Problem Statement 2.xlsx")
MSGS = os.path.join(HERE, "messages.json")
ORIGINAL = os.path.join(HERE, "messages_claude_original.json")
FIRST_ROW = 5

# 1) Drafts as they are in the workbook now
ws = openpyxl.load_workbook(wb_path, data_only=True)["Messages"]
rows = []
r = FIRST_ROW
while isinstance(ws.cell(r, 1).value, str) and re.fullmatch(r"C\d{4}", ws.cell(r, 1).value.strip()):
    mid = ws.cell(r, 1).value.strip()
    rows.append((r, {"id": mid, "subject": ws.cell(r, 4).value or "", "body": ws.cell(r, 5).value or "",
                     "check": ws.cell(r, 6).value or ""}))
    r += 1

stored = json.load(open(MSGS, encoding="utf-8"))
original = json.load(open(ORIGINAL if os.path.exists(ORIGINAL) else MSGS, encoding="utf-8"))
orig_by_id = {m["id"]: m for m in original["csm"]}
for _, m in rows:
    o = orig_by_id.get(m["id"])
    m["edited"] = o is None or any((m[k] or "").strip() != (o.get(k) or "").strip() for k in ("subject", "body", "check"))
new_msgs = {"csm": [m for _, m in rows], "templates": stored["templates"]}

# 2) Reference results and the fact check (same script that checked the original drafts), in a temp folder first
out_dir = tempfile.mkdtemp()
msg_file = os.path.join(out_dir, "messages.json")
ref_file, fc_file = os.path.join(out_dir, "reference.json"), os.path.join(out_dir, "factcheck.json")
ref, _ = R.run(SRC)
json.dump(ref, open(ref_file, "w"), indent=1, default=float)
json.dump(new_msgs, open(msg_file, "w", encoding="utf-8"), indent=1, ensure_ascii=False)
done = subprocess.run([sys.executable, "-I", os.path.join(HERE, "factcheck.py"), SRC, msg_file, ref_file, fc_file],
                      capture_output=True, text=True)
print(done.stdout, done.stderr, sep="")
if done.returncode != 0:
    sys.exit("The fact check did not run; nothing written.")
facts = json.load(open(fc_file))
edited = [m["id"] for _, m in rows if m["edited"]]
print("Drafts you have edited:", edited or "none")
if not WRITE:
    print("\nCheck only: nothing was changed. Add --write to update the workbook and scripts/*.json.")
    sys.exit(0)

# 3) Write the results into the Messages tab, touching only those cells
lock = os.path.join(os.path.dirname(wb_path), "~$" + os.path.basename(wb_path))
if os.path.exists(lock):
    sys.exit("The workbook looks open in Excel (lock file found). Close it and run again; nothing written to it.")
os.makedirs(os.path.join(PROJ, "backups"), exist_ok=True)
stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
backup = os.path.join(PROJ, "backups", f"P2_Churn_Scoring.before-recheck-{stamp}.xlsx")
shutil.copy2(wb_path, backup)

with zipfile.ZipFile(backup) as z:
    items = [(i, z.read(i.filename)) for i in z.infolist()]
files = {i.filename: d for i, d in items}
wbx = files["xl/workbook.xml"].decode("utf-8")
sheet_el = next(e for e in re.findall(r"<sheet\b[^>]*>", wbx) if 'name="Messages"' in e)
rid = re.search(r'\br:id="([^"]+)"', sheet_el).group(1)
rels = files["xl/_rels/workbook.xml.rels"].decode("utf-8")
rel_el = next(e for e in re.findall(r"<Relationship\b[^>]*>", rels) if f'Id="{rid}"' in e)
target = re.search(r'\bTarget="([^"]+)"', rel_el).group(1)
sheet_name = "xl/" + target.lstrip("/").removeprefix("xl/")
x = files[sheet_name].decode("utf-8")


def put(xml, ref_, value):
    m = re.search(rf'<c r="{ref_}"(?P<a>[^>]*?)(?:/>|>.*?</c>)', xml, re.S)
    if not m:
        sys.exit(f"Cell {ref_} not found on Messages; workbook not changed (backup at {backup}).")
    s = re.search(r'\ss="(\d+)"', m.group("a"))
    style = f' s="{s.group(1)}"' if s else ""
    if isinstance(value, (int, float)):
        new = f'<c r="{ref_}"{style}><v>{value}</v></c>'
    else:
        new = f'<c r="{ref_}"{style} t="inlineStr"><is><t xml:space="preserve">{escape(value)}</t></is></c>'
    return xml[:m.start()] + new + xml[m.end():]


for r, m in rows:
    fc = facts[m["id"]]
    x = put(x, f"G{r}", fc["numbers"])
    x = put(x, f"H{r}", fc["status"])
    x = put(x, f"I{r}", fc["words"])
    x = put(x, f"J{r}", "Claude draft, edited by you" if m["edited"] else "Claude (AI draft)")
n_ed = len(edited)
note = ("Drafted by Claude from each merchant's row" + (f"; {n_ed} edited by you" if n_ed else "") +
        f". Every number is checked automatically against the data (Fact check column, last run "
        f"{dt.date.today():%d %b %Y} with scripts/recheck_drafts.py). Nothing is sent until a CSM reviews and approves it.")
x = put(x, "A2", note)
files[sheet_name] = x.encode("utf-8")

with zipfile.ZipFile(wb_path, "w") as dst:
    for info, _ in items:
        dst.writestr(info, files[info.filename], compress_type=info.compress_type)

# 4) Only after the workbook is written: keep Claude's originals once, then update scripts/*.json
if not os.path.exists(ORIGINAL):
    shutil.copy2(MSGS, ORIGINAL)
for f in (msg_file, ref_file, fc_file):
    shutil.copy2(f, os.path.join(HERE, os.path.basename(f)))
fails = sum(not v["status"].startswith("Passed") for v in facts.values())
print(f"\nWorkbook updated: {len(rows)} drafts, {fails} failing, {n_ed} edited. Backup: {backup}")
