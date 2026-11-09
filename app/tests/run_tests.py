"""Section 11 tests 1, 2 and 4, plus checks that the app's words match the workbook.

Run from the project folder:  python -B app/tests/run_tests.py
"""
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
sys.dont_write_bytecode = True

import openpyxl  # noqa: E402

import core  # noqa: E402
import fixtures  # noqa: E402

results = []


def check(name, ok, detail=""):
    results.append(ok)
    print(f"  [{'PASS' if ok else 'FAIL'}] {name}" + (f": {detail}" if detail else ""))


def run(data, confirmed=None):
    chk = core.check_file(data)
    out, order = core.score(data, confirmed)
    table = core.explain(chk.df, out, order, core.playbooks_table())
    return chk, out, order, table


wb_raw, wb_ranks, wb_tracks = core.workbook_snapshot()

# ---------------------------------------------------------------- Test 1
print("\nTest 1: the original file gives the same 100 ranks as the submitted workbook")
data = fixtures.original_bytes()
chk, out, order, table = run(data)
check("file passes the upload check", chk.ok, "; ".join(chk.errors) or f"{len(chk.df)} merchants")
check("no unseen notes", core.unseen_notes(chk.df) == [])
same = core.same_data(chk.df, wb_raw)
check("data is identical to the workbook's Scoring tab", same)
m = core.compare_ranks(out, wb_ranks, wb_tracks)
check(f"badge: matches the submitted workbook: {m['ranks']} of {m['n']} ranks", m["ranks"] == m["n"] == 100)
check(f"tracks match: {m['tracks']} of {m['n']}", m["tracks"] == 100)
ref = json.load(open(core.SCRIPTS / "reference.json"))
check("ranks also match scripts/reference.json", all(out[k]["rank"] == ref[k]["rank"] for k in ref))
csm = [k for k in order if out[k]["track"] == "CSM outreach"]
lt = [k for k in order if out[k]["track"] == "Light-touch"]
check("CSM 15 in the handoff order", csm == ["C2018", "C2022", "C2036", "C2092", "C2040", "C2004", "C2088", "C2090",
                                            "C2062", "C2079", "C2027", "C2097", "C2061", "C2100", "C2093"], ", ".join(csm))
check("light-touch 9 in the handoff order", lt == ["C2084", "C2091", "C2001", "C2034", "C2032", "C2015", "C2045",
                                                  "C2008", "C2020"], ", ".join(lt))

# words shown in the app vs the workbook's saved values
wbv = openpyxl.load_workbook(core.WORKBOOK, data_only=True, read_only=True)
cm = json.load(open(core.SCRIPTS / "colmap.json"))
from openpyxl.utils import column_index_from_string as ci  # noqa: E402

L = cm["scoring"]
xl = {}
for row in wbv["Scoring"].iter_rows(min_row=cm["first"], max_row=cm["last"], values_only=True):
    g = lambda k: row[ci(L[k]) - 1]  # noqa: E731
    xl[g("id")] = dict(reason=g("reason"), risk=g("risk"), playbook=g("pb"), also=g("also") or "",
                       sources=g("sources"))
by = table.set_index("merchant_id")
for field in ("reason", "risk", "playbook", "also"):
    bad = [k for k in xl if by.loc[k, field] != xl[k][field]]
    check(f"{field} text matches the workbook for all 100", not bad, ", ".join(bad[:5]))
lt_ws = wbv["Light-touch"]
lt_xl = {r[3]: r[16] for r in lt_ws.iter_rows(min_row=6, max_row=45, values_only=True) if r[3]}
bad = [k for k, txt in lt_xl.items() if core.template(by.loc[k].to_dict() | {"merchant_id": k}) != txt]
check(f"light-touch templates match the workbook's Light-touch tab ({len(lt_xl)} merchants)", not bad and len(lt_xl) == 9,
      ", ".join(bad))
stored = [k for k in csm if core.draft_for(by.loc[k].to_dict() | {"merchant_id": k}, True)["kind"] == "stored"]
check("all 15 CSM merchants get their stored, fact-checked draft", len(stored) == 15 and all(
    core.FACTS[k]["status"] == "Passed" for k in stored))

# ---------------------------------------------------------------- Test 2
print("\nTest 2: reworded notes trigger an 'unseen notes' step instead of silently dropping Pricing flags")
rdata, answer = fixtures.reworded_bytes()
rchk = core.check_file(rdata)
unseen = core.unseen_notes(rchk.df)
check("reworded file passes the upload check", rchk.ok)
check(f"{len(unseen)} unseen notes found (every note was reworded)", len(unseen) == len(answer) == 22)
pricing_unseen = [n for n in unseen if answer[n] == "Pricing"]
check(f"{len(pricing_unseen)} of them are Pricing notes", len(pricing_unseen) == 4, " | ".join(pricing_unseen))
part = dict(list(answer.items())[:21])
check("the gate holds scoring until every unseen note is tagged (21 of 22 tagged: still held)",
      core.pending_notes(rchk.df, {}) == unseen and len(core.pending_notes(rchk.df, part)) == 1
      and core.pending_notes(rchk.df, answer) == [])
# what would happen without the gate: the danger the step exists to stop
out0, order0 = core.score(rdata, {})
pri0 = sum(o["main"] == "Pricing" or (o["flagged"] and o["tag"] == "Pricing") for o in out0.values())
csm0 = [k for k in order0 if out0[k]["track"] == "CSM outreach"]
check(f"without the step: Pricing-driven merchants drop from 6 to {pri0}, and {len(set(csm) - set(csm0))} of the "
      "CSM 15 change", pri0 == 0 and len(set(csm) - set(csm0)) == 6)
# with a person's tags (the reviewed tags of the original notes), the result is the original again
out1, order1 = core.score(rdata, answer)
check("after a person tags them correctly, all 100 ranks equal the original's",
      all(out1[k]["rank"] == out[k]["rank"] and out1[k]["track"] == out[k]["track"] for k in out))
check("reviewed tags can't be overridden by a confirmed tag",
      core.score(data, {"Comparing quotes from two providers": "Neutral"})[0]["C2018"]["track"] == "CSM outreach")

# ---------------------------------------------------------------- Test 3 (offline part)
print("\nTest 3: a planted 'offer them a discount' gets the draft blocked by the fact check")
pdata = fixtures.planted_bytes()
pchk = core.check_file(pdata)
check("planted note shows up as an unseen note", core.unseen_notes(pchk.df) == [fixtures.PLANTED_NOTE])
pout, porder = core.score(pdata, {fixtures.PLANTED_NOTE: "Pricing"})
check(f"tagged Pricing by a person, {fixtures.PLANTED_ID} stays on the CSM list",
      pout[fixtures.PLANTED_ID]["track"] == "CSM outreach")
c2004 = core.STORED[fixtures.PLANTED_ID]["body"]
planted = [
    ("a draft that echoes the note", c2004.replace("Could you", "We can offer you a discount before you renew. Could you"),
     "banned wording: 'discount'"),
    ("a 'lower rate' promise", c2004 + " We'll get you a lower rate.", "banned wording: 'lower rate'"),
    ("a number not in the row", c2004.replace("39 days", "45 days"), "number 45 not in this merchant's data"),
    ("another merchant's name", c2004 + " Banyan Mart already switched.", "mentions another merchant: Banyan Mart"),
    ("a draft without the driver's phrase", "Hi [contact name], your renewal is in 39 days. Can we talk?",
     "does not match its driver (Pricing)"),
    ("another playbook's promise (app rule)", c2004 + " We'll share a fix plan within 3 days.",
     "promises PB-1's 3-day fix plan, but this merchant's driver is Pricing"),
]
verdicts = core.fact_check(pdata, pout, [dict(id=fixtures.PLANTED_ID, subject="Before your renewal", body=b,
                                             kind="ai", track="CSM outreach") for _, b, _ in planted])
for (name, _, phrase), v in zip(planted, verdicts):
    check(f"blocked: {name}", not v["ok"] and any(phrase in p for p in v["problems"]), "; ".join(v["problems"]))
ok_v = core.fact_check(data, out, [dict(id=k, subject=core.STORED[k]["subject"], body=core.STORED[k]["body"],
                                        kind="stored", track="CSM outreach") for k in csm])
check("all 15 stored drafts pass the live check", all(v["ok"] for v in ok_v))
lt_items = []
for k in lt:
    d = core.draft_for(by.loc[k].to_dict() | {"merchant_id": k}, True)
    lt_items.append(dict(id=k, subject=d["subject"], body=d["body"], kind="template", track="Light-touch"))
check("all 9 light-touch templates pass the live check", all(v["ok"] for v in core.fact_check(data, out, lt_items)))

# ---------------------------------------------------------------- Test 4
print("\nTest 4: a broken file gets a clear error")
for name, (bad_bytes, phrase) in fixtures.broken_files().items():
    c = core.check_file(bad_bytes)
    msg = " ".join(c.errors)
    check(f"{name} rejected", not c.ok and phrase in msg, msg)

# ---------------------------------------------------------------- Live AI checks (only with --live; they cost a little)
if "--live" in sys.argv:
    import ai  # noqa: E402

    tags_live = "--drafts-only" not in sys.argv  # --drafts-only skips the 27 tag calls
    if tags_live:
        print("\nLive: DeepSeek tag suggestions on the 22 reworded notes")
        sug = ai.suggest_tags(unseen, "DeepSeek")
        right = [n for n in unseen if sug[n]["tag"] == answer[n]]
        check(f"DeepSeek suggests the reviewed tag for {len(right)} of {len(unseen)}", len(right) >= 20,
              "; ".join(f"{n!r}: {sug[n]['tag'] or sug[n]['error']} (key {answer[n]})" for n in unseen if n not in right))
        check("DeepSeek catches all 4 Pricing notes", all(sug[n]["tag"] == "Pricing" for n in pricing_unseen))
        print("\nLive: the Jev switch, on the 4 Pricing notes")
        jsug = ai.suggest_tags(pricing_unseen, "Jev")
        check("Jev catches all 4 Pricing notes", all(jsug[n]["tag"] == "Pricing" for n in pricing_unseen),
              "; ".join(f"{jsug[n]['tag'] or jsug[n]['error']} ({jsug[n]['confidence']})" for n in pricing_unseen))

    print("\nLive: test 3 end to end (planted note -> DeepSeek tag -> DeepSeek draft -> fact check)")
    if tags_live:
        ptag = ai.suggest_tags([fixtures.PLANTED_NOTE], "DeepSeek")[fixtures.PLANTED_NOTE]
        check(f"DeepSeek tags the planted note: {ptag['tag']}", ptag["tag"] == "Pricing")
    pt = core.explain(pchk.df, pout, porder, core.playbooks_table()).set_index("merchant_id")
    pm = pt.loc[fixtures.PLANTED_ID].to_dict() | {"merchant_id": fixtures.PLANTED_ID}
    pbk = core.playbooks_table().set_index("playbook_id").loc[pm["playbook"]].to_dict()
    for attempt in range(1, 4):
        dr = ai.draft(pm, pbk)
        ve, vw = core.fact_check(pdata, pout, [
            dict(id=fixtures.PLANTED_ID, subject=dr["subject"], body=dr["email"], kind="ai", track="CSM outreach"),
            dict(id=fixtures.PLANTED_ID, subject="", body=dr["whatsapp"], kind="ai", track="CSM outreach")])
        said = "discount" in (dr["email"] + dr["whatsapp"]).lower()
        print(f"     attempt {attempt}: mentions 'discount': {said} | email {'Passed' if ve['ok'] else ve['problems']}"
              f" | WhatsApp {'Passed' if vw['ok'] else vw['problems']}")
        if said:
            check("a draft that mentions the discount is blocked", not ve["ok"] or not vw["ok"])
    print("     email:", dr["subject"], "|", dr["email"].replace("\n", " ")[:400])
    print("     whatsapp:", dr["whatsapp"].replace("\n", " "))

    print("\nLive: DeepSeek drafts for one CSM merchant per driver, fact-checked")
    for k in ("C2018", "C2088", "C2079", "C2090", "C2061"):
        mk = by.loc[k].to_dict() | {"merchant_id": k}
        pbm = core.playbooks_table().set_index("playbook_id").loc[mk["playbook"]].to_dict()
        dr = ai.draft(mk, pbm)
        ve, vw = core.fact_check(data, out, [
            dict(id=k, subject=dr["subject"], body=dr["email"], kind="ai", track="CSM outreach"),
            dict(id=k, subject="", body=dr["whatsapp"], kind="ai", track="CSM outreach")])
        miss = ai.missing_facts(mk, dr["email"]) + ai.missing_facts(mk, dr["whatsapp"])
        check(f"{k} drafts state the key facts: " + "; ".join(ai.must_include(mk)), not miss, "; ".join(miss))
        print(f"     {k} ({mk['main']}): email {len(dr['email'].split())} words, WhatsApp "
              f"{len(dr['whatsapp'].split())} words | email {'Passed' if ve['ok'] else ve['problems']} | "
              f"WhatsApp {'Passed' if vw['ok'] else vw['problems']}")
        print("        email:", dr["subject"], "|", dr["email"].replace("\n", " "))
        print("        whatsapp:", dr["whatsapp"].replace("\n", " "))
        results.append(ve["ok"] and vw["ok"])

print(f"\n{sum(results)} of {len(results)} checks passed")
sys.exit(0 if all(results) else 1)
