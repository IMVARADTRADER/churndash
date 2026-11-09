"""Reference implementation of the user's decided rules, used to check the Excel formulas.

Usage: python -I reference.py <Problem Statement 2.xlsx> <out.json>
"""
import json
import sys
import pandas as pd

S = dict(AlarmLine=-0.105, SRDrop=5, UnresolvedMin=2, LowLogins=5, PricingRenewal=60, LowCSAT=2,
         Cap=15, UrgentDays=30, W={"Performance": 1.0, "Service": 0.9, "Pricing": 0.8, "Disengagement": 0.8},
         U30=1.5, U60=1.25, U90=1.1, ValuePower=1.0, UseLikelihood=1, L={1: 0.5, 2: 0.7, 3: 0.85, 4: 1.0},
         NetRev=0.01, ValueMonths=12, StableBand=0.10, PB5Products=2, GoodCSAT=4, LowVolume=3,
         BusyTickets=5, SeasonalDip=0.15,
         SeasonalInd={"Gifting & Festive", "Education / Coaching", "Travel & Hospitality"})

TAGS = {
    "Owner handed operations to a new manager, unclear contact": "Disengagement",
    "No response to last two check-ins": "Disengagement",
    "Dashboard rarely used since new manager joined": "Disengagement",
    "Seasonal dip as per last year pattern": "Seasonal",
    "Seasonal business; volumes dip after the festive season every year": "Seasonal",
    "Merchant said slow quarter after peak season, expects pickup": "Seasonal",
    "Asked for a review of success rates; no clear answer yet": "Performance",
    "Merchant complained about UPI drop-offs on last call": "Performance",
    "Says support replies take too long": "Service",
    "Escalated to AM twice this quarter": "Service",
    "Raised several tickets on settlements; still not closed": "Service",
    "Asked for the rate card ahead of renewal": "Pricing",
    "Comparing quotes from two providers": "Pricing",
    "Mentioned a competitor offering lower MDR": "Pricing",
    "Said 'we are exploring options'": "Pricing",
    "Small merchant, low touch": "Neutral",
    "Never raised tickets": "Neutral",
    "Complained once about invoice format; fixed": "Neutral",
    "Asked about settlement reports; resolved": "Neutral",
    "Interested in adding subscriptions product": "Positive",
    "Quarterly check-in done, no concerns": "Positive",
    "Growing fast; many integration questions, all resolved": "Positive",
}


def run(src):
    df = pd.read_excel(src, sheet_name="P2_Merchants")
    out = {}
    keys = {}
    for i, r in df.iterrows():
        excel_row = i + 2
        t = [r[f"tpv_m{k}_inr_lakh"] for k in range(1, 7)]
        avg = sum(t) / 6
        latest = r.tpv_m6_inr_lakh / avg - 1
        lowest = min(t) / avg - 1
        src_ = r.success_rate_m4_m6_pct - r.success_rate_m1_m3_pct
        note = r.account_manager_note if isinstance(r.account_manager_note, str) else ""
        tag = TAGS.get(note, "None" if not note else "Untagged")
        csat = r.last_csat_1_to_5
        has_csat = not pd.isna(csat)
        alarm = latest <= S["AlarmLine"]
        perf = src_ <= -S["SRDrop"] or tag == "Performance"
        serv = r.unresolved_tickets >= S["UnresolvedMin"] or tag == "Service"
        pri = tag == "Pricing" and r.days_to_contract_renewal <= S["PricingRenewal"]
        dis = alarm and (r.logins_last_30d <= S["LowLogins"] or tag == "Disengagement")
        flagged = perf or serv or pri or dis
        main = "Performance" if perf else "Service" if serv else "Pricing" if pri else "Disengagement" if dis else ""
        sources = (int(alarm) + int(src_ <= -S["SRDrop"])
                   + int(r.unresolved_tickets >= S["UnresolvedMin"] or (has_csat and csat <= S["LowCSAT"]))
                   + int(r.logins_last_30d <= S["LowLogins"])
                   + int(tag in ("Performance", "Service", "Pricing", "Disengagement")))
        value = r.tpv_m6_inr_lakh * 100000 * S["NetRev"] * S["ValueMonths"]
        d = r.days_to_contract_renewal
        urg = S["U30"] if d <= 30 else S["U60"] if d <= 60 else S["U90"] if d <= 90 else 1.0
        lik = (S["L"][min(max(sources, 1), 4)] if S["UseLikelihood"] == 1 else 1.0)
        score = (value ** S["ValuePower"]) * S["W"].get(main, 0) * urg * lik if flagged else None
        urgent = flagged and d <= S["UrgentDays"]
        key = ((1e8 if urgent else 0) + score) if flagged else (-1000 + sources * 10 - latest)
        key += (1000 - excel_row) / 1e7
        keys[r.merchant_id] = key
        pb5 = (not flagged and abs(latest) <= S["StableBand"] and lowest >= -S["StableBand"]
               and (not has_csat or csat >= S["GoodCSAT"]) and r.products_used <= S["PB5Products"])
        out[r.merchant_id] = dict(flagged=flagged, main=main, sources=sources, value=value, score=score,
                                  urgent=urgent, pb5=pb5, tag=tag, latest=latest)
    order = sorted(keys, key=lambda m: -keys[m])
    for rank, m in enumerate(order, start=1):
        o = out[m]
        o["rank"] = rank
        if o["flagged"]:
            o["track"] = "CSM outreach" if rank <= S["Cap"] else "Light-touch"
        else:
            o["track"] = "Expansion (PB-5)" if o["pb5"] else "Monitor"
    return out, order


if __name__ == "__main__":
    out, order = run(sys.argv[1])
    with open(sys.argv[2], "w") as f:
        json.dump(out, f, indent=1, default=float)
    csm = [m for m in order if out[m]["track"] == "CSM outreach"]
    lt = [m for m in order if out[m]["track"] == "Light-touch"]
    print("Flagged:", sum(o["flagged"] for o in out.values()))
    print("CSM top 15:", [(m, out[m]["main"], "URGENT" if out[m]["urgent"] else "") for m in csm])
    print("Light-touch:", [(m, out[m]["main"]) for m in lt])
    print("PB-5:", [m for m in order if out[m]["track"] == "Expansion (PB-5)"])
    print("Monitor count:", sum(o["track"] == "Monitor" for o in out.values()))
