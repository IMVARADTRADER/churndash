"""Fact-check the CSM drafts against each merchant's row.

Checks: every number in subject+body exists in that merchant's data (or is a playbook constant),
no other merchant's business name appears, no banned promise words, and the message fits the driver.
Usage: python -I factcheck.py <Problem Statement 2.xlsx> <messages.json> <reference.json> <out.json>
"""
import json
import re
import sys

import pandas as pd

src, msgp, refp, outp = sys.argv[1:5]
df = pd.read_excel(src, sheet_name="P2_Merchants").set_index("merchant_id")
msgs = json.load(open(msgp, encoding="utf-8"))["csm"]
ref = json.load(open(refp))

CONSTANTS = {3.0: "PB-1 fix plan within 3 days", 48.0: "PB-2 senior callback within 48 hours", 90.0: "ticket window (90 days)"}
BANNED = ["discount", "waive", "waiver", "refund", "guarantee", "free of", "for free", "lower rate", "match their",
          "match the", "cheaper", "compensat", "cashback", "will be settled by", "by tomorrow"]
DRIVER_WORDS = {"Pricing": ["rate card"], "Performance": ["Payments Ops"], "Service": ["single point of contact"],
                "Disengagement": ["right contact", "right person"]}
all_names = set(df["business_name"])
FIELDS = {"success_rate_m1_m3_pct": "success rate m1-m3", "success_rate_m4_m6_pct": "success rate m4-m6",
          "support_tickets_last_90d": "tickets in 90 days", "unresolved_tickets": "unresolved tickets",
          "last_csat_1_to_5": "last CSAT", "days_to_contract_renewal": "days to renewal",
          "products_used": "products used", "logins_last_30d": "logins (30 days)", "tenure_months": "tenure (months)"}

results, failures = {}, 0
for m in msgs:
    mid = m["id"]
    row = df.loc[mid]
    allowed = {}
    for col, label in FIELDS.items():
        v = row[col]
        if not pd.isna(v):
            allowed.setdefault(round(float(v), 1), []).append(label)
    text = m["subject"] + " " + m["body"]
    nums = [float(x) for x in re.findall(r"(?<![A-Za-z\-])(\d+(?:\.\d+)?)", text)]
    used, problems = [], []
    for n in nums:
        n1 = round(n, 1)
        if n1 in allowed:
            used.append(f"{n:g} = " + " / ".join(allowed[n1]))
        elif n1 in CONSTANTS:
            used.append(f"{n:g} = {CONSTANTS[n1]}")
        else:
            problems.append(f"number {n:g} not in this merchant's data")
    for name in all_names:
        if name in text and name != row["business_name"]:
            problems.append(f"mentions another merchant: {name}")
    low = text.lower()
    for b in BANNED:
        if b in low:
            problems.append(f"banned wording: '{b}'")
    driver = ref[mid]["main"]
    if not any(w.lower() in low for w in DRIVER_WORDS[driver]):
        problems.append(f"does not match its driver ({driver})")
    if ref[mid]["track"] != "CSM outreach":
        problems.append(f"merchant is not on the CSM list (track = {ref[mid]['track']})")
    words = len(m["body"].split())
    status = "Passed" if not problems else "FAILED: " + "; ".join(problems)
    failures += bool(problems)
    results[mid] = {"status": status, "numbers": "; ".join(dict.fromkeys(used)) or "none", "words": words, "driver": driver}
    print(f"{mid} [{driver}, {words} words]: {status} | {results[mid]['numbers']}")

csm_ids = {k for k, v in ref.items() if v["track"] == "CSM outreach"}
missing = csm_ids - {m["id"] for m in msgs}
print("\nCSM merchants without a draft:", sorted(missing) or "none")
print("Failures:", failures)
json.dump(results, open(outp, "w"), indent=1)
