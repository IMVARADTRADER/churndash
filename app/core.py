"""Logic for the churn early-warning app. No Streamlit in here, so the tests can import it directly.

What lives where:
- The rules that flag and rank merchants come only from scripts/reference.py, imported unchanged.
- This file reads and checks the uploaded file, finds notes the reviewed tag table hasn't seen, and builds the
  words the workbook shows (reason, risk label, playbook, evidence, light-touch template) from the same settings.
- Nothing is written to disk. The uploaded file stays in memory for the session.
"""
from __future__ import annotations

import hashlib
import io
import json
import os
import subprocess
import sys
import tempfile
import threading
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
WORKBOOK = ROOT / "P2_Churn_Scoring.xlsx"
ANSWERS_PDF = ROOT / "Problem 2 Written Answers Razorpay Build-a-thon.pdf"
ORIGINAL_DATA = ROOT / "Problem Statement 2.xlsx"

sys.dont_write_bytecode = True  # importing the scripts must not leave __pycache__ files in scripts/
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))
import reference  # noqa: E402  (the user's rules, unchanged)
from test_tagging import DEFINITIONS  # noqa: E402  (tag definitions used in the tagging test)

S = reference.S
REVIEWED_TAGS = dict(reference.TAGS)  # the 22 notes reviewed on 9 November
TAG_CHOICES = ["Performance", "Service", "Pricing", "Disengagement", "Seasonal", "Positive", "Neutral"]
RISK_TAGS = {"Performance", "Service", "Pricing", "Disengagement"}
TAG_EFFECT = {  # what each tag does in scripts/reference.py (shown on the Tag notes page)
    "Pricing": "The only way the Pricing flag fires: a Pricing note plus renewal within 60 days. Also counts as "
               "evidence (the account manager source).",
    "Performance": "Fires the Performance flag on its own, even without a success-rate drop. Counts as evidence.",
    "Service": "Fires the Service flag on its own, even with fewer than 2 open tickets. Counts as evidence.",
    "Disengagement": "Fires the Disengagement flag only together with the TPV alarm. Counts as evidence.",
    "Seasonal": "No flag. Labels an unflagged volume dip as a seasonal dip, so it isn't read as churn.",
    "Positive": "No effect on flags or ranking.",
    "Neutral": "No effect on flags or ranking.",
}
SHEET = "P2_Merchants"
MAX_BYTES = 5 * 1024 * 1024
MAX_ROWS = 999  # reference.py breaks ties with (1000 - Excel row), so it assumes fewer than 1,000 rows

TEXT_COLS = ["merchant_id", "business_name", "segment", "industry"]
NUM_COLS = ["tenure_months"] + [f"tpv_m{k}_inr_lakh" for k in range(1, 7)] + [
    "success_rate_m1_m3_pct", "success_rate_m4_m6_pct", "support_tickets_last_90d", "unresolved_tickets",
    "products_used", "logins_last_30d", "days_to_contract_renewal"]
OPTIONAL_NUM = ["last_csat_1_to_5"]  # blank CSAT is allowed and ignored (rule 4)
NOTE_COL = "account_manager_note"  # blank note is allowed
REQUIRED = TEXT_COLS + NUM_COLS + OPTIONAL_NUM + [NOTE_COL]
RAW_COLS = ["merchant_id", "business_name", "segment", "industry", "tenure_months"] + [
    f"tpv_m{k}_inr_lakh" for k in range(1, 7)] + [
    "success_rate_m1_m3_pct", "success_rate_m4_m6_pct", "support_tickets_last_90d", "unresolved_tickets",
    "last_csat_1_to_5", "products_used", "logins_last_30d", "days_to_contract_renewal", NOTE_COL]

PLAYBOOK_FOR = {"Performance": "PB-1", "Service": "PB-2", "Pricing": "PB-3", "Disengagement": "PB-4"}
SIGN = "[Your name], Customer Success, Razorpay"


# ---------------------------------------------------------------- 1. Upload and check
@dataclass
class Checked:
    ok: bool
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    df: pd.DataFrame | None = None
    digest: str = ""


def check_file(data: bytes) -> Checked:
    """Read the P2_Merchants sheet and check names, columns and types. Returns plain-language errors."""
    if not data:
        return Checked(False, ["The file is empty."])
    if len(data) > MAX_BYTES:
        return Checked(False, [f"The file is {len(data) / 1e6:.1f} MB. The limit is 5 MB."])
    try:
        book = pd.ExcelFile(io.BytesIO(data), engine="openpyxl")
    except Exception:
        return Checked(False, ["This isn't an Excel .xlsx file the app can read. Save it as .xlsx and try again."])
    if SHEET not in book.sheet_names:
        return Checked(False, [f"There's no sheet called '{SHEET}'. Sheets found: {', '.join(book.sheet_names)}."])
    try:
        df = pd.read_excel(book, sheet_name=SHEET)
    except Exception as e:  # pragma: no cover
        return Checked(False, [f"The '{SHEET}' sheet couldn't be read ({type(e).__name__})."])

    errors: list[str] = []
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        return Checked(False, [f"Missing column{'s' if len(missing) > 1 else ''}: {', '.join(missing)}. "
                               "Column names must match the Data Pack exactly."])
    if df.empty:
        return Checked(False, ["The sheet has headers but no merchants."])
    if len(df) > MAX_ROWS:
        return Checked(False, [f"The sheet has {len(df)} merchants. The limit is {MAX_ROWS}."])

    def where(mask):
        rows = [str(i + 2) for i in df.index[mask][:5]]
        more = int(mask.sum()) - len(rows)
        return "Excel row " + ", ".join(rows) + (f" and {more} more" if more > 0 else "")

    for c in TEXT_COLS:
        blank = df[c].isna() | (df[c].astype(str).str.strip() == "")
        if blank.any():
            errors.append(f"'{c}' is blank in {where(blank)}.")
    ids = df["merchant_id"].astype(str).str.strip()
    dup = ids.duplicated(keep=False) & df["merchant_id"].notna()
    if dup.any():
        errors.append(f"Merchant IDs repeat: {', '.join(sorted(set(ids[dup])))}.")
    notes = df[NOTE_COL]
    bad_note = notes.notna() & ~notes.map(lambda v: isinstance(v, str))
    if bad_note.any():
        errors.append(f"'{NOTE_COL}' must be text or blank; check {where(bad_note)}.")

    for c in NUM_COLS + OPTIONAL_NUM:
        vals = pd.to_numeric(df[c], errors="coerce")
        not_num = df[c].notna() & vals.isna()
        if not_num.any():
            errors.append(f"'{c}' must be a number; check {where(not_num)}.")
        if c in NUM_COLS and df[c].isna().any():
            errors.append(f"'{c}' is blank in {where(df[c].isna())}. Only CSAT and the note may be blank.")
        if not_num.any() or (c in NUM_COLS and df[c].isna().any()):
            continue
        if c in OPTIONAL_NUM:
            out = vals.notna() & ((vals < 1) | (vals > 5))
            if out.any():
                errors.append(f"'{c}' must be 1 to 5 or blank; check {where(out)}.")
        elif c.startswith("success_rate"):
            out = (vals < 0) | (vals > 100)
            if out.any():
                errors.append(f"'{c}' must be between 0 and 100; check {where(out)}.")
        elif (vals < 0).any():
            errors.append(f"'{c}' can't be negative; check {where(vals < 0)}.")
    if errors:
        return Checked(False, errors)

    tpv = df[[f"tpv_m{k}_inr_lakh" for k in range(1, 7)]].astype(float)
    zero = tpv.sum(axis=1) == 0
    if zero.any():
        return Checked(False, [f"Six months of zero TPV in {where(zero)}: the TPV alarm can't compare to an average of 0."])

    warnings = []
    odd = df["unresolved_tickets"] > df["support_tickets_last_90d"]
    if odd.any():
        warnings.append(f"{int(odd.sum())} merchant{'s have' if odd.sum() > 1 else ' has'} more unresolved tickets "
                        "than tickets raised in 90 days (probably older tickets). Kept as they are.")
    if "city" not in df.columns:
        warnings.append("No 'city' column. The rules don't use it, so that's fine.")
    return Checked(True, [], warnings, df, hashlib.sha256(data).hexdigest()[:16])


# ---------------------------------------------------------------- 2. Notes the reviewed table hasn't seen
def note_table(df: pd.DataFrame) -> pd.DataFrame:
    """Each distinct note, how many merchants have it, and its reviewed tag (blank if unseen)."""
    notes = df[NOTE_COL].dropna()
    counts = notes.value_counts()
    return pd.DataFrame({"note": counts.index, "merchants": counts.values,
                         "reviewed_tag": [REVIEWED_TAGS.get(n, "") for n in counts.index]})


def unseen_notes(df: pd.DataFrame) -> list[str]:
    return [n for n in note_table(df)["note"] if n not in REVIEWED_TAGS]


def pending_notes(df: pd.DataFrame, confirmed: dict[str, str]) -> list[str]:
    """Unseen notes no person has tagged yet. Scoring is held while this list is not empty."""
    return [n for n in unseen_notes(df) if confirmed.get(n) not in TAG_CHOICES]


# ---------------------------------------------------------------- 3. Score with reference.py
_LOCK = threading.Lock()


def score(data: bytes, confirmed: dict[str, str] | None = None) -> tuple[dict, list[str]]:
    """Run reference.run() on the uploaded file. Person-confirmed tags for unseen notes are added to the
    lookup table for this run only; the 22 reviewed tags always win."""
    extra = {k: v for k, v in (confirmed or {}).items() if k not in REVIEWED_TAGS}
    with _LOCK:  # reference.TAGS is a module-level table, so one run at a time
        saved = reference.TAGS
        reference.TAGS = {**extra, **REVIEWED_TAGS}
        try:
            return reference.run(io.BytesIO(data))
        finally:
            reference.TAGS = saved


# ---------------------------------------------------------------- 4/5. Words the workbook shows
def xl_round(x: float, places: int) -> Decimal:
    """Round like Excel's TEXT(): halves away from zero."""
    return Decimal(repr(float(x))).quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)


def pct0(x: float) -> str:
    return f"{xl_round(x * 100, 0)}%"


def num(v) -> str:
    """How Excel writes a number joined into text: 2.0 -> '2', 4.5 -> '4.5'."""
    f = float(v)
    return str(int(f)) if f.is_integer() else repr(f)


def has(v) -> bool:
    return v is not None and not pd.isna(v)


def explain(df: pd.DataFrame, out: dict, order: list[str], playbooks: pd.DataFrame) -> pd.DataFrame:
    """One row per merchant, in rank order, with the workbook's reason, risk label, playbook and evidence.

    The four driver flags are recomputed here only to show them; every merchant is checked against
    reference.py's own result, so the page can never disagree with the ranking."""
    pb_name = dict(zip(playbooks["playbook_id"], playbooks["name"]))
    rows = []
    by_id = df.set_index(df["merchant_id"].astype(str))
    for mid in order:
        r, o = by_id.loc[mid], out[mid]
        t = [float(r[f"tpv_m{k}_inr_lakh"]) for k in range(1, 7)]
        avg = sum(t) / 6
        latest, lowest = t[5] / avg - 1, min(t) / avg - 1
        srchg = float(r.success_rate_m4_m6_pct) - float(r.success_rate_m1_m3_pct)
        note = r[NOTE_COL] if isinstance(r[NOTE_COL], str) else ""
        tag, csat = o["tag"], r.last_csat_1_to_5
        unres, logins, renew = int(r.unresolved_tickets), int(r.logins_last_30d), int(r.days_to_contract_renewal)

        alarm = latest <= S["AlarmLine"]
        src = {"Volume": alarm, "Payments": srchg <= -S["SRDrop"],
               "Support": unres >= S["UnresolvedMin"] or (has(csat) and csat <= S["LowCSAT"]),
               "Engagement": logins <= S["LowLogins"], "Account manager note": tag in RISK_TAGS}
        flags = {"Performance": srchg <= -S["SRDrop"] or tag == "Performance",
                 "Service": unres >= S["UnresolvedMin"] or tag == "Service",
                 "Pricing": tag == "Pricing" and renew <= S["PricingRenewal"],
                 "Disengagement": alarm and (logins <= S["LowLogins"] or tag == "Disengagement")}
        main = next((d for d in ("Performance", "Service", "Pricing", "Disengagement") if flags[d]), "")
        if (any(flags.values()), main, sum(src.values())) != (o["flagged"], o["main"], o["sources"]):
            raise RuntimeError(f"{mid}: the evidence shown would disagree with reference.py. Nothing is displayed.")
        also = " ".join(d for d in ("Service", "Pricing", "Disengagement") if flags[d] and d != main)

        htpv = (f"latest TPV {pct0(-latest)} below 6-month average" if latest < 0
                else f"latest TPV {pct0(latest)} above 6-month average")
        am = f'AM note "{note}"'
        if o["flagged"]:
            helper = {
                "Performance": (f"success rate fell from {xl_round(r.success_rate_m1_m3_pct, 1)}% to "
                                f"{xl_round(r.success_rate_m4_m6_pct, 1)}%") if src["Payments"] else am,
                "Service": (f"{unres} unresolved tickets" + (f", last CSAT {num(csat)}/5" if has(csat) else ""))
                if unres >= S["UnresolvedMin"] else am,
                "Pricing": f"{am}, renewal in {renew} days",
                "Disengagement": (f"only {logins} {'login' if logins == 1 else 'logins'} in 30 days"
                                  if logins <= S["LowLogins"] else "")
                + ((", " if logins <= S["LowLogins"] else "") + am if tag == "Disengagement" else ""),
            }[main]
            reason = f"{main}: {helper}; {htpv}" + (f" (also {also})" if also else "")
            risk, kind, pb = ("High" if o["sources"] >= 3 else "Medium"), "", PLAYBOOK_FOR[main]
        else:
            first_min = t.index(min(t)) + 1
            if o["pb5"]:
                kind = "Healthy, expansion candidate"
                reason = f"Stable TPV, good or no CSAT, uses {int(r.products_used)} product(s): expansion candidate"
            elif tag == "Seasonal" or (r.industry in S["SeasonalInd"] and lowest <= -S["SeasonalDip"] and first_min >= 4):
                kind = "Seasonal dip"
                reason = (f"Seasonal dip: lowest month {pct0(-lowest)} below average, {htpv}"
                          + (" (AM note confirms)" if tag == "Seasonal" else ""))
            elif alarm:
                kind, reason = "TPV dip, no driver", f"No driver signal yet; {htpv}. Watch next month"
            elif t[5] < S["LowVolume"]:
                kind, reason = "Low-volume stable", f"Low volume (Rs {xl_round(t[5], 1)} lakh/month), stable"
            elif r.support_tickets_last_90d >= S["BusyTickets"]:
                kind = "Busy but healthy"
                reason = f"Busy but healthy: {int(r.support_tickets_last_90d)} tickets in 90 days, {unres} unresolved"
            else:
                kind, reason = "Healthy", "No risk signal"
            risk, pb = ("Low (watch)" if alarm else "Low"), ("PB-5" if o["pb5"] else "PB-6")

        d = renew
        urg = S["U30"] if d <= 30 else S["U60"] if d <= 60 else S["U90"] if d <= 90 else 1.0
        lik = S["L"][min(max(o["sources"], 1), 4)] if S["UseLikelihood"] == 1 else 1.0
        rows.append(dict(
            rank=o["rank"], merchant_id=mid, business_name=r.business_name, segment=r.segment, industry=r.industry,
            track=o["track"], risk=risk, sources=o["sources"], main=main, also=also, playbook=pb,
            playbook_name=pb_name.get(pb, ""), value=o["value"], score=o["score"], renewal=renew, urgent=o["urgent"],
            reason=reason, kind=kind, tag=tag, note=note, latest=latest, lowest=lowest, avg=avg, tpv=t,
            sr1=float(r.success_rate_m1_m3_pct), sr2=float(r.success_rate_m4_m6_pct), srchg=srchg,
            tickets=int(r.support_tickets_last_90d), unresolved=unres, csat=float(csat) if has(csat) else None,
            logins=logins, products=int(r.products_used), tenure=int(r.tenure_months),
            evidence=src, flags=flags, weight=S["W"].get(main, 0.0), urgency=urg, likelihood=lik))
    return pd.DataFrame(rows)


# ---------------------------------------------------------------- 6. Drafts
def template(m: dict) -> str:
    """The light-touch template, filled from the data exactly as the workbook's Light-touch tab fills it."""
    hi = "Hi [contact name], "
    if m["main"] == "Performance":
        body = (f"your payment success rate on Razorpay was {xl_round(m['sr2'], 1)}% over the last three months, down "
                f"from {xl_round(m['sr1'], 1)}% in the three months before. " if m["srchg"] <= -S["SRDrop"]
                else "you recently told us about payment failures. ") + (
            "We would like to help fix this. Reply to this email and our payments team will review failures by payment "
            "method and bank with you, then share a fix plan within 3 days of that review. ")
    elif m["main"] == "Service":
        body = (f"you have {m['unresolved']} support tickets still open with us. " if m["unresolved"] >= S["UnresolvedMin"]
                else "you told us our support has been slow. ") + (
            "I am sorry this has taken longer than it should. I am now your single point of contact. Reply with a good "
            "time to talk and I will go through each open issue with you and confirm where it stands. ")
    elif m["main"] == "Pricing":
        body = (f"your Razorpay contract is due for renewal in {m['renewal']} days. Before then, your account manager can "
                "walk you through your current rate card, volume tiers and renewal terms. Reply to this email to set up a time. ")
    else:
        body = (f"we want to make sure we are in touch with the right person at {m['business_name']} for your Razorpay "
                "account. Could you reply with the best contact for payments and reporting? We would also be glad to run "
                "a short session on the reports and products on your account. ")
    return hi + body + SIGN


def load_json(name: str) -> dict:
    with open(SCRIPTS / name, encoding="utf-8") as f:
        return json.load(f)


STORED = {m["id"]: m for m in load_json("messages.json")["csm"]}
FACTS = load_json("factcheck.json")


def draft_for(m: dict, same_as_submitted: bool, prefer_template: bool = False) -> dict | None:
    """The stored draft (original file, CSM list) or the rules template. Returns None when there's no outreach."""
    if (m["track"] == "CSM outreach" and same_as_submitted and m["merchant_id"] in STORED
            and not prefer_template):
        s, fc = STORED[m["merchant_id"]], FACTS.get(m["merchant_id"], {})
        return dict(kind="stored", subject=s["subject"], body=s["body"], check=s["check"],
                    by="Claude draft, edited by you" if s.get("edited") else "Claude (AI draft)",
                    fact_status=fc.get("status", "Not checked"), fact_numbers=fc.get("numbers", ""))
    if m["track"] in ("CSM outreach", "Light-touch"):
        subject = {"Performance": "Your payment success rate", "Service": "Your open support tickets",
                   "Pricing": "Before your renewal", "Disengagement": "Checking we have the right contact"}[m["main"]]
        return dict(kind="template", subject=subject, body=template(m),
                    check="Template filled by the rules from this merchant's row. Read it once before approving.",
                    by="Rules (template, no AI)", fact_status="", fact_numbers="")
    return None


# ---------------------------------------------------------------- 7. Live fact check (scripts/factcheck.py, unchanged)
# Two rules are relaxed, and only where they can't apply:
#  - rules-filled templates choose their wording by driver already, so the driver key-phrase rule is skipped for them
#    (the workbook's Performance template says "payments team", not "Payments Ops");
#  - light-touch merchants are meant to be off the CSM list, so that rule is skipped for them.
RELAX_TEMPLATE = "does not match its driver"
RELAX_LIGHT = "merchant is not on the CSM list"
# One rule the app adds on top of factcheck.py, which allows 3 and 48 for every merchant: a playbook's time
# commitment may appear only in that playbook's messages. (A live DeepSeek draft for a PB-4 merchant promised
# PB-1's "fix plan within 3 days" and still passed factcheck.py.)
PLAYBOOK_PROMISES = [("Performance", r"\b3\s+(?:working\s+|business\s+)?days?\b", "PB-1's 3-day fix plan"),
                     ("Service", r"\b48\s*(?:-\s*)?hours?\b", "PB-2's 48-hour callback")]


def promise_problems(text: str, main: str) -> list[str]:
    import re
    return [f"promises {what}, but this merchant's driver is {main}"
            for owner, pattern, what in PLAYBOOK_PROMISES if owner != main and re.search(pattern, text, re.I)]


def fact_check(data: bytes, out: dict, items: list[dict]) -> list[dict]:
    """Run scripts/factcheck.py on each item {id, subject, body, kind, track} and return its verdicts.

    factcheck.py reads files, so the merchant file and the drafts go into a temporary folder that is
    deleted as soon as the check finishes (a second or two)."""
    def one(args):
        i, it, d = args
        msg = os.path.join(d, f"m{i}.json")
        res = os.path.join(d, f"r{i}.json")
        with open(msg, "w", encoding="utf-8") as f:
            json.dump({"csm": [{"id": it["id"], "subject": it.get("subject", ""), "body": it["body"]}]}, f)
        run = subprocess.run([sys.executable, "-I", str(SCRIPTS / "factcheck.py"), src, msg, ref, res],
                             capture_output=True, text=True, timeout=120)
        if run.returncode != 0 or not os.path.exists(res):
            return dict(ok=False, problems=["the fact check could not run"], numbers="", words=0)
        with open(res) as f:
            r = json.load(f)[it["id"]]
        status = r["status"]
        problems = [] if status == "Passed" else status.removeprefix("FAILED: ").split("; ")
        if it.get("kind") == "template":
            problems = [p for p in problems if not p.startswith(RELAX_TEMPLATE)]
        if it.get("track") == "Light-touch":
            problems = [p for p in problems if not p.startswith(RELAX_LIGHT)]
        problems += promise_problems(it.get("subject", "") + " " + it["body"], r["driver"])
        return dict(ok=not problems, problems=problems, numbers=r["numbers"], words=r["words"])

    with tempfile.TemporaryDirectory() as d:
        src = os.path.join(d, "data.xlsx")
        ref = os.path.join(d, "reference.json")
        with open(src, "wb") as f:
            f.write(data)
        with open(ref, "w") as f:
            json.dump(out, f, default=float)
        with ThreadPoolExecutor(4) as ex:
            return list(ex.map(one, [(i, it, d) for i, it in enumerate(items)]))


# ---------------------------------------------------------------- 3b. Does it match the submitted workbook?
def workbook_snapshot() -> tuple[pd.DataFrame, dict, dict]:
    """The submitted workbook's raw data and the ranks and tracks Excel calculated (saved values)."""
    import openpyxl
    cm = load_json("colmap.json")
    L = cm["scoring"]
    wb = openpyxl.load_workbook(WORKBOOK, data_only=True, read_only=True)
    ws = wb["Scoring"]
    raw, ranks, tracks = [], {}, {}
    from openpyxl.utils import column_index_from_string as ci
    for row in ws.iter_rows(min_row=cm["first"], max_row=cm["last"], values_only=True):
        raw.append(row[:len(RAW_COLS)])
        ranks[row[ci(L["id"]) - 1]] = row[ci(L["rank"]) - 1]
        tracks[row[ci(L["id"]) - 1]] = row[ci(L["track"]) - 1]
    wb.close()
    return pd.DataFrame(raw, columns=RAW_COLS), ranks, tracks


def same_data(df: pd.DataFrame, wb_raw: pd.DataFrame) -> bool:
    """True when the upload holds exactly the workbook's merchants, in the same row order, with the same values."""
    if len(df) != len(wb_raw):
        return False
    for c in RAW_COLS:
        a, b = df[c].tolist(), wb_raw[c].tolist()
        for x, y in zip(a, b):
            xn, yn = not has(x), not has(y)
            if xn or yn:
                if xn != yn:
                    return False
            elif isinstance(x, str) or isinstance(y, str):
                if str(x) != str(y):
                    return False
            elif abs(float(x) - float(y)) > 1e-9:
                return False
    return True


def compare_ranks(out: dict, wb_ranks: dict, wb_tracks: dict) -> dict:
    rank_same = sum(out[m]["rank"] == wb_ranks.get(m) for m in out)
    track_same = sum(out[m]["track"] == wb_tracks.get(m) for m in out)
    return dict(n=len(out), ranks=rank_same, tracks=track_same,
                differ=[m for m in out if out[m]["rank"] != wb_ranks.get(m)])


def playbooks_table() -> pd.DataFrame:
    """P2_Playbooks, as copied unchanged into the workbook's Playbooks tab."""
    pb = pd.read_excel(WORKBOOK, sheet_name="Playbooks", nrows=6)
    return pb[["playbook_id", "name", "what_happens", "when_it_applies"]]
