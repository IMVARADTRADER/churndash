"""Build P2_Churn_Scoring.xlsx: the user's churn rules as live Excel formulas.

Usage: python -I build_scoring.py <Problem Statement 2.xlsx> <P2_Labelling.xlsx> <out.xlsx> <colmap.json>
"""
import json
import sys

import openpyxl
import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

SRC, LAB, OUT, COLMAP = sys.argv[1:5]
MSGS = json.load(open(sys.argv[5], encoding="utf-8")) if len(sys.argv) > 5 else None
FACTS = json.load(open(sys.argv[6], encoding="utf-8")) if len(sys.argv) > 6 else {}
merchants = pd.read_excel(SRC, sheet_name="P2_Merchants")
playbooks = pd.read_excel(SRC, sheet_name="P2_Playbooks")
assumptions = pd.read_excel(SRC, sheet_name="P2_Assumptions")
N = len(merchants)
FIRST, LAST = 2, N + 1

# ---------------------------------------------------------------- styles
FONT = "Arial"
NAVY, BLUE_HDR, GREY_HDR, GREEN_HDR = "0B2A4A", "1F5FBF", "7A8494", "2E7D4F"
f_base = Font(name=FONT, size=10)
f_bold = Font(name=FONT, size=10, bold=True)
f_hdr = Font(name=FONT, size=10, bold=True, color="FFFFFF")
f_title = Font(name=FONT, size=14, bold=True, color=NAVY)
f_sub = Font(name=FONT, size=11, bold=True, color=NAVY)
f_input = Font(name=FONT, size=10, color="0000FF")
f_note = Font(name=FONT, size=9, italic=True, color="555555")
fill_input = PatternFill("solid", fgColor="FFF4CC")
thin = Side(style="thin", color="C9CED6")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap_top = Alignment(wrap_text=True, vertical="top")
center = Alignment(horizontal="center", vertical="center", wrap_text=True)

TRACK_COLORS = {"CSM outreach": "F8CBAD", "Light-touch": "FFE699", "Expansion (PB-5)": "C6E0B4", "Monitor": "E7E9EE"}


def hdr(ws, row, col, text, color=NAVY, width=None):
    c = ws.cell(row=row, column=col, value=text)
    c.font, c.fill, c.alignment, c.border = f_hdr, PatternFill("solid", fgColor=color), center, border
    if width:
        ws.column_dimensions[get_column_letter(col)].width = width
    return c


def track_cf(ws, rng):
    for text, color in TRACK_COLORS.items():
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{text}"'],
                                                      fill=PatternFill("solid", fgColor=color)))


def set_font_all(ws):
    for row in ws.iter_rows():
        for c in row:
            if c.font is None or c.font.name != FONT:
                c.font = Font(name=FONT, size=c.font.size if c.font else 10, bold=c.font.bold if c.font else False,
                              italic=c.font.italic if c.font else False, color=c.font.color if c.font else None)


wb = Workbook()

# ================================================================ Read me (filled at the end)
readme = wb.active
readme.title = "Read me"

# ================================================================ Settings
st = wb.create_sheet("Settings")
st["A1"] = "Settings: every threshold and weight used by the scoring"
st["A1"].font = f_title
st["A2"] = "Edit only the yellow cells with blue numbers. Every other tab recalculates from these."
st["A2"].font = f_note
for col, (text, w) in enumerate([("Setting", 46), ("Value", 12), ("What it means / where it comes from", 95)], start=1):
    hdr(st, 3, col, text, width=w)

settings = [
    ("section", "Flag rules"),
    ("AlarmLine", "TPV alarm line (latest month vs own 6-month average)", -0.105, "0.0%",
     "Your decision. Tuned on this data: it sits between C2094 (seasonal, -10.1%) and C2036 (-11.9%)."),
    ("SRDrop", "Success-rate drop that counts as a performance signal (points)", 5, "0.0",
     "Months 4-6 vs months 1-3. In this data every drop is either under 2 points or over 5."),
    ("UnresolvedMin", "Unresolved tickets that count as a service signal", 2, "0",
     "All 10 merchants at 2+ unresolved also have CSAT 2 or below, so CSAT is merged into this signal."),
    ("LowLogins", "Low logins (30 days) for disengagement", 5, "0",
     "Only counts when the TPV alarm is also on (your decision), so tiny stable merchants are not flagged."),
    ("PricingRenewal", "Pricing flag: renewal within (days)", 60, "0",
     "Pricing flag = pricing note + renewal within this window. No TPV alarm needed (your decision)."),
    ("LowCSAT", "Low CSAT (counts as a support signal)", 2, "0", "Blank CSAT is ignored (your decision)."),
    ("section", "Ranking"),
    ("Cap", "CSM outreach cap this month", 15, "0", "Source: P2_Assumptions (capacity)."),
    ("UrgentDays", "Urgent renewal: flagged merchants renewing within (days) skip the queue", 30, "0",
     "Your decision: only merchants renewing before next month's re-check skip the queue. They are ranked first, "
     "then everyone else by priority score. At 60 days, 8 urgent merchants took 8 slots and pushed C2061 out."),
    ("EscalateDays", "Light-touch escalation: renewal within (days) at next month's re-check", 60, "0",
     "Your decision: a light-touch merchant moves into next month's CSM list if it replies, its signals worsen, "
     "or its renewal is within this window."),
    ("W_Perf", "Savability weight: Performance", 1.0, "0.00",
     "Your decision (9 Nov 2026): Claude proposed the starting numbers; you tested them and kept them. "
     "The order is yours (Performance > Service > Pricing = Disengagement)."),
    ("W_Serv", "Savability weight: Service", 0.9, "0.00", "Your decision after testing (9 Nov 2026)."),
    ("W_Price", "Savability weight: Pricing", 0.8, "0.00", "Your decision after testing (9 Nov 2026)."),
    ("W_Dis", "Savability weight: Disengagement", 0.8, "0.00", "Your decision after testing (9 Nov 2026)."),
    ("Urg_Days1", "Urgency band 1: renewal within (days)", 30, "0", "Your decision after testing (9 Nov 2026)."),
    ("Urg_Boost1", "Urgency boost for band 1", 1.5, "0.00", "Your decision after testing (9 Nov 2026)."),
    ("Urg_Days2", "Urgency band 2: renewal within (days)", 60, "0", "Your decision after testing (9 Nov 2026)."),
    ("Urg_Boost2", "Urgency boost for band 2", 1.25, "0.00", "Your decision after testing (9 Nov 2026)."),
    ("Urg_Days3", "Urgency band 3: renewal within (days)", 90, "0", "Your decision after testing (9 Nov 2026)."),
    ("Urg_Boost3", "Urgency boost for band 3", 1.1, "0.00", "Your decision after testing (9 Nov 2026). Beyond band 3 the boost is 1."),
    ("ValuePower", "Value weight (1 = full value, 0.5 = square root, 0 = ignore size)", 1.0, "0.0",
     "Tested 9 November with the likelihood term on: 0.5 swaps one merchant (C2084 in, C2097 out); 0 swaps five."),
    ("UseLikelihood", "Use likelihood term? (1 = on, 0 = off)", 1, "0",
     "Likelihood = how many of 5 independent sources see the problem (see Scoring tab). Turned on 9 November 2026."),
    ("Lik_1", "Likelihood multiplier: 1 source", 0.5, "0.00", "Your decision after testing (9 Nov 2026)."),
    ("Lik_2", "Likelihood multiplier: 2 sources", 0.7, "0.00", "Your decision after testing (9 Nov 2026)."),
    ("Lik_3", "Likelihood multiplier: 3 sources", 0.85, "0.00", "Your decision after testing (9 Nov 2026)."),
    ("Lik_4", "Likelihood multiplier: 4 or more sources", 1.0, "0.00", "Your decision after testing (9 Nov 2026)."),
    ("section", "Value and cost (from P2_Assumptions)"),
    ("NetRev", "Net revenue as share of TPV", 0.01, "0.0%", "Source: P2_Assumptions (net revenue)."),
    ("ValueMonths", "Months of revenue counted as value of a saved merchant", 12, "0",
     "Source: P2_Assumptions (value of a saved merchant = 12 months at latest TPV)."),
    ("FalseAlarmCost", "Cost of one outreach / false alarm (Rs)", 1500, "#,##0",
     "Source: P2_Assumptions (1.5 hours of CSM time)."),
    ("SaveRate", "Scenario: share of contacted at-risk merchants a CSM call saves", 0.25, "0%",
     "Assumption for the Summary scenario only, NOT from the data. Change it to test the business case."),
    ("section", "Not-at-risk labels and PB-5"),
    ("StableBand", "Stable TPV band (latest and lowest month within +/- this of average)", 0.10, "0%",
     "Used for PB-5. Excludes seasonal dippers, whose lowest month is 23-26% below average."),
    ("PB5Products", "PB-5: products used at or below", 2, "0", "Your PB-5 rule: stable TPV, good or no CSAT, few products."),
    ("GoodCSAT", "PB-5: good CSAT at or above", 4, "0", "Blank CSAT also passes (your decision to ignore missing CSAT)."),
    ("LowVolume", "Low-volume label: latest TPV below (Rs lakh)", 3, "0.0", "Descriptive label only; does not change flags."),
    ("BusyTickets", "Busy-but-healthy label: tickets in 90 days at or above", 5, "0", "Descriptive label only."),
    ("SeasonalDip", "Seasonal-dip label: lowest month at least this far below average", 0.15, "0%",
     "Descriptive label only, for seasonal industries; the lowest month must fall in months 4-6."),
]
names = {}
r = 4
for item in settings:
    if item[0] == "section":
        c = st.cell(row=r, column=1, value=item[1])
        c.font = f_sub
        r += 1
        continue
    name, label, value, fmt, note = item
    st.cell(row=r, column=1, value=label).font = f_base
    v = st.cell(row=r, column=2, value=value)
    v.font, v.fill, v.number_format, v.border = f_input, fill_input, fmt, border
    v.alignment = Alignment(horizontal="right")
    n = st.cell(row=r, column=3, value=note)
    n.font, n.alignment = f_base, Alignment(wrap_text=True, vertical="top")
    names[name] = f"Settings!$B${r}"
    r += 1

r += 1
st.cell(row=r, column=1, value="Seasonal industries (source: P2_Assumptions)").font = f_sub
r += 1
seas_first = r
for ind in ["Gifting & Festive", "Education / Coaching", "Travel & Hospitality"]:
    c = st.cell(row=r, column=2, value=ind)
    c.font, c.fill, c.border = f_input, fill_input, border
    r += 1
names["SeasonalInd"] = f"Settings!$B${seas_first}:$B${r - 1}"
st.column_dimensions["B"].width = 24
dv01 = DataValidation(type="list", formula1='"0,1"', allow_blank=False)
st.add_data_validation(dv01)
dv01.add(names["UseLikelihood"].split("!")[1].replace("$", ""))

# ================================================================ Note tags
nt = wb.create_sheet("Note tags")
nt["A1"] = "Account manager note tags"
nt["A1"].font = f_title
nt["A2"] = ("Each distinct note is tagged once; the Scoring tab looks the tag up. Tags were suggested by Claude using "
            "keyword rules; you reviewed all 22 and kept them on 9 November 2026. Change a yellow cell to retag a note.")
nt["A2"].font = f_note
for col, (text, w) in enumerate([("Account manager note (exact text)", 62), ("Tag", 16),
                                 ("Merchants with this note", 14), ("Suggested by", 26), ("Reviewed by you? (Y/N)", 14)],
                                start=1):
    hdr(nt, 3, col, text, width=w)
TAGS = {
    "Owner handed operations to a new manager, unclear contact": "Disengagement",
    "No response to last two check-ins": "Disengagement",
    "Dashboard rarely used since new manager joined": "Disengagement",
    "Asked for a review of success rates; no clear answer yet": "Performance",
    "Merchant complained about UPI drop-offs on last call": "Performance",
    "Says support replies take too long": "Service",
    "Escalated to AM twice this quarter": "Service",
    "Raised several tickets on settlements; still not closed": "Service",
    "Asked for the rate card ahead of renewal": "Pricing",
    "Comparing quotes from two providers": "Pricing",
    "Mentioned a competitor offering lower MDR": "Pricing",
    "Said 'we are exploring options'": "Pricing",
    "Seasonal dip as per last year pattern": "Seasonal",
    "Seasonal business; volumes dip after the festive season every year": "Seasonal",
    "Merchant said slow quarter after peak season, expects pickup": "Seasonal",
    "Interested in adding subscriptions product": "Positive",
    "Quarterly check-in done, no concerns": "Positive",
    "Growing fast; many integration questions, all resolved": "Positive",
    "Small merchant, low touch": "Neutral",
    "Never raised tickets": "Neutral",
    "Complained once about invoice format; fixed": "Neutral",
    "Asked about settlement reports; resolved": "Neutral",
}
data_notes = set(merchants["account_manager_note"].dropna().unique())
assert data_notes == set(TAGS), f"note mismatch: {data_notes ^ set(TAGS)}"
nt_first = 4
for i, (note_text, tag) in enumerate(TAGS.items()):
    rr = nt_first + i
    nt.cell(row=rr, column=1, value=note_text).font = f_base
    t = nt.cell(row=rr, column=2, value=tag)
    t.font, t.fill, t.border = f_input, fill_input, border
    nt.cell(row=rr, column=3, value=f"=COUNTIF(Scoring!$T${FIRST}:$T${LAST},A{rr})").font = f_base
    nt.cell(row=rr, column=4, value="Claude (keyword rules)").font = f_base
    y = nt.cell(row=rr, column=5, value="Y")  # reviewed by the user, 9 Nov 2026
    y.fill, y.border, y.font = fill_input, border, f_input
nt_last = nt_first + len(TAGS) - 1
names["NoteTags"] = f"'Note tags'!$A${nt_first}:$B${nt_last}"
dv_tag = DataValidation(type="list", formula1='"Performance,Service,Pricing,Disengagement,Seasonal,Positive,Neutral"')
nt.add_data_validation(dv_tag)
dv_tag.add(f"B{nt_first}:B{nt_last}")
dv_yn = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
nt.add_data_validation(dv_yn)
dv_yn.add(f"E{nt_first}:E{nt_last}")
nt.freeze_panes = "A4"

# ================================================================ Playbooks
pbs = wb.create_sheet("Playbooks")
for col, (text, w) in enumerate([("playbook_id", 12), ("name", 28), ("what_happens", 70), ("when_it_applies", 55)], start=1):
    hdr(pbs, 1, col, text, width=w)
for i, row in playbooks.iterrows():
    for j, col in enumerate(playbooks.columns, start=1):
        c = pbs.cell(row=i + 2, column=j, value=row[col])
        c.font, c.alignment = f_base, wrap_top
names["PBTable"] = f"Playbooks!$A$2:$B${len(playbooks) + 1}"
pbs["A10"] = "Source: Data Pack, P2_Playbooks (copied unchanged)."
pbs["A10"].font = f_note

# ================================================================ Scoring
sc = wb.create_sheet("Scoring", 1)
RAW = [
    ("id", "merchant_id", 11), ("name", "business_name", 18), ("seg", "segment", 11), ("ind", "industry", 18),
    ("ten", "tenure_months", 8),
    ("t1", "tpv_m1_inr_lakh", 8), ("t2", "tpv_m2_inr_lakh", 8), ("t3", "tpv_m3_inr_lakh", 8),
    ("t4", "tpv_m4_inr_lakh", 8), ("t5", "tpv_m5_inr_lakh", 8), ("t6", "tpv_m6_inr_lakh", 8),
    ("sr1", "success_rate_m1_m3_pct", 9), ("sr2", "success_rate_m4_m6_pct", 9),
    ("tix", "support_tickets_last_90d", 8), ("unres", "unresolved_tickets", 9), ("csat", "last_csat_1_to_5", 8),
    ("prod", "products_used", 8), ("logins", "logins_last_30d", 8), ("renew", "days_to_contract_renewal", 9),
    ("note", "account_manager_note", 34),
]
CALC = [
    # key, header, formula, number format, width, group
    ("avg", "Avg TPV, 6 months (Rs lakh)", "=AVERAGE({t1}{r}:{t6}{r})", "0.0", 10, "calc"),
    ("latest", "Latest month vs avg", "=IF({avg}{r}=0,0,{t6}{r}/{avg}{r}-1)", "0.0%", 10, "calc"),
    ("lowest", "Lowest month vs avg", "=IF({avg}{r}=0,0,MIN({t1}{r}:{t6}{r})/{avg}{r}-1)", "0.0%", 10, "calc"),
    ("srchg", "Success rate change (points)", "={sr2}{r}-{sr1}{r}", "0.0", 10, "calc"),
    ("tag", "Note tag", '=IF({note}{r}="","None",IFERROR(VLOOKUP({note}{r},NoteTags,2,FALSE),"Untagged"))', None, 13, "calc"),
    ("alarm", "TPV alarm", "={latest}{r}<=AlarmLine", None, 8, "flag"),
    ("fperf", "Performance flag", '=OR({srchg}{r}<=-SRDrop,{tag}{r}="Performance")', None, 11, "flag"),
    ("fserv", "Service flag", '=OR({unres}{r}>=UnresolvedMin,{tag}{r}="Service")', None, 9, "flag"),
    ("fpri", "Pricing flag", '=AND({tag}{r}="Pricing",{renew}{r}<=PricingRenewal)', None, 9, "flag"),
    ("fdis", "Disengagement flag (needs TPV alarm)",
     '=AND({alarm}{r},OR({logins}{r}<=LowLogins,{tag}{r}="Disengagement"))', None, 13, "flag"),
    ("flag", "Flagged (any driver)", "=OR({fperf}{r},{fserv}{r},{fpri}{r},{fdis}{r})", None, 9, "flag"),
    ("main", "Main driver",
     '=IF({fperf}{r},"Performance",IF({fserv}{r},"Service",IF({fpri}{r},"Pricing",IF({fdis}{r},"Disengagement",""))))',
     None, 13, "out"),
    ("also", "Also flagged for",
     '=TRIM(IF(AND({fserv}{r},{main}{r}<>"Service"),"Service ","")&IF(AND({fpri}{r},{main}{r}<>"Pricing"),"Pricing ","")'
     '&IF(AND({fdis}{r},{main}{r}<>"Disengagement"),"Disengagement",""))', None, 12, "out"),
    ("sources", "Evidence sources (of 5)",
     '=({alarm}{r})*1+({srchg}{r}<=-SRDrop)*1+OR({unres}{r}>=UnresolvedMin,AND(ISNUMBER({csat}{r}),{csat}{r}<=LowCSAT))*1'
     '+({logins}{r}<=LowLogins)*1+OR({tag}{r}="Performance",{tag}{r}="Service",{tag}{r}="Pricing",{tag}{r}="Disengagement")*1',
     "0", 9, "calc"),
    ("risk", "Risk label", '=IF({flag}{r},IF({sources}{r}>=3,"High","Medium"),IF({alarm}{r},"Low (watch)","Low"))',
     None, 10, "out"),
    ("value", "Value at stake (Rs)", "={t6}{r}*100000*NetRev*ValueMonths", "#,##0", 12, "calc"),
    ("weight", "Savability weight",
     '=IF({main}{r}="Performance",W_Perf,IF({main}{r}="Service",W_Serv,IF({main}{r}="Pricing",W_Price,'
     'IF({main}{r}="Disengagement",W_Dis,""))))', "0.00", 9, "calc"),
    ("urg", "Urgency boost", "=IF({renew}{r}<=Urg_Days1,Urg_Boost1,IF({renew}{r}<=Urg_Days2,Urg_Boost2,IF({renew}{r}<=Urg_Days3,Urg_Boost3,1)))", "0.00", 9, "calc"),
    ("lik", "Likelihood multiplier",
     "=IF(UseLikelihood=1,IF({sources}{r}>=4,Lik_4,IF({sources}{r}=3,Lik_3,IF({sources}{r}=2,Lik_2,Lik_1))),1)", "0.00", 10, "calc"),
    ("score", "Priority score", '=IF({flag}{r},({value}{r}^ValuePower)*{weight}{r}*{urg}{r}*{lik}{r},"")', "#,##0", 11, "calc"),
    ("urgent", "Urgent renewal (skips queue)", "=AND({flag}{r},{renew}{r}<=UrgentDays)", None, 10, "flag"),
    ("key", "Sort key (urgent first, then score; unflagged last)",
     "=IF({flag}{r},IF({urgent}{r},100000000,0)+{score}{r},-1000+{sources}{r}*10-{latest}{r})+(1000-ROW())/10000000",
     "#,##0.00", 13, "help"),
    ("rank", "Rank (1 = contact first)", "=RANK({key}{r},${key}$" + str(FIRST) + ":${key}$" + str(LAST) + ",0)", "0", 9, "out"),
    ("pb5", "PB-5 eligible",
     "=AND(NOT({flag}{r}),ABS({latest}{r})<=StableBand,{lowest}{r}>=-StableBand,"
     "OR(NOT(ISNUMBER({csat}{r})),{csat}{r}>=GoodCSAT),{prod}{r}<=PB5Products)", None, 9, "flag"),
    ("track", "Track",
     '=IF({flag}{r},IF({rank}{r}<=Cap,"CSM outreach","Light-touch"),IF({pb5}{r},"Expansion (PB-5)","Monitor"))', None, 15, "out"),
    ("pb", "Playbook",
     '=IF({flag}{r},IF({main}{r}="Performance","PB-1",IF({main}{r}="Service","PB-2",IF({main}{r}="Pricing","PB-3","PB-4"))),'
     'IF({pb5}{r},"PB-5","PB-6"))', None, 9, "out"),
    ("pbname", "Playbook name", '=IFERROR(VLOOKUP({pb}{r},PBTable,2,FALSE),"")', None, 22, "out"),
    ("type", "Not-at-risk type",
     '=IF({flag}{r},"",IF({pb5}{r},"Healthy, expansion candidate",IF(OR({tag}{r}="Seasonal",AND(COUNTIF(SeasonalInd,{ind}{r})>0,'
     '{lowest}{r}<=-SeasonalDip,MATCH(MIN({t1}{r}:{t6}{r}),{t1}{r}:{t6}{r},0)>=4)),"Seasonal dip",IF({alarm}{r},"TPV dip, no driver",IF({t6}{r}<LowVolume,"Low-volume stable",'
     'IF({tix}{r}>=BusyTickets,"Busy but healthy","Healthy"))))))', None, 16, "out"),
    ("reason", "Reason (plain language)",
     '=IF({flag}{r},{main}{r}&": "&IF({main}{r}="Performance",{hperf}{r},IF({main}{r}="Service",{hserv}{r},'
     'IF({main}{r}="Pricing",{hpri}{r},{hdis}{r})))&"; "&{htpv}{r}&IF({also}{r}<>""," (also "&{also}{r}&")",""),'
     'IF({type}{r}="Healthy, expansion candidate","Stable TPV, good or no CSAT, uses "&{prod}{r}&" product(s): expansion candidate",'
     'IF({type}{r}="Seasonal dip","Seasonal dip: lowest month "&TEXT(-{lowest}{r},"0%")&" below average, "&{htpv}{r}'
     '&IF({tag}{r}="Seasonal"," (AM note confirms)",""),'
     'IF({type}{r}="TPV dip, no driver","No driver signal yet; "&{htpv}{r}&". Watch next month",'
     'IF({type}{r}="Low-volume stable","Low volume (Rs "&TEXT({t6}{r},"0.0")&" lakh/month), stable",'
     'IF({type}{r}="Busy but healthy","Busy but healthy: "&{tix}{r}&" tickets in 90 days, "&{unres}{r}&" unresolved",'
     '"No risk signal"))))))', None, 70, "out"),
    ("hperf", "Helper: performance text",
     '=IF({srchg}{r}<=-SRDrop,"success rate fell from "&TEXT({sr1}{r},"0.0")&"% to "&TEXT({sr2}{r},"0.0")&"%",'
     '"AM note """&{note}{r}&"""")', None, 30, "help"),
    ("hserv", "Helper: service text",
     '=IF({unres}{r}>=UnresolvedMin,{unres}{r}&" unresolved tickets"&IF(ISNUMBER({csat}{r}),", last CSAT "&{csat}{r}&"/5",""),'
     '"AM note """&{note}{r}&"""")', None, 30, "help"),
    ("hpri", "Helper: pricing text", '="AM note """&{note}{r}&""", renewal in "&{renew}{r}&" days"', None, 30, "help"),
    ("hdis", "Helper: disengagement text",
     '=IF({logins}{r}<=LowLogins,"only "&{logins}{r}&IF({logins}{r}=1," login"," logins")&" in 30 days","")&IF({tag}{r}="Disengagement",'
     'IF({logins}{r}<=LowLogins,", ","")&"AM note """&{note}{r}&"""","")', None, 30, "help"),
    ("htpv", "Helper: TPV text",
     '=IF({latest}{r}<0,"latest TPV "&TEXT(-{latest}{r},"0%")&" below 6-month average",'
     '"latest TPV "&TEXT({latest}{r},"0%")&" above 6-month average")', None, 30, "help"),
]
L = {}
for i, (k, _, _) in enumerate(RAW, start=1):
    L[k] = get_column_letter(i)
for i, item in enumerate(CALC, start=len(RAW) + 1):
    L[item[0]] = get_column_letter(i)

GROUP_COLOR = {"raw": GREY_HDR, "calc": BLUE_HDR, "flag": "8E5A9B", "out": GREEN_HDR, "help": "A0A7B4"}
for i, (k, h, w) in enumerate(RAW, start=1):
    hdr(sc, 1, i, h, GROUP_COLOR["raw"], w)
for i, (k, h, f, fmt, w, g) in enumerate(CALC, start=len(RAW) + 1):
    hdr(sc, 1, i, h, GROUP_COLOR[g], w)
sc.row_dimensions[1].height = 48

for idx, row in merchants.iterrows():
    rr = FIRST + idx
    for i, (k, col, _) in enumerate(RAW, start=1):
        v = row[col]
        if pd.isna(v):
            v = None
        elif isinstance(v, float) and v.is_integer() and k not in ("t1", "t2", "t3", "t4", "t5", "t6", "sr1", "sr2"):
            v = int(v)
        c = sc.cell(row=rr, column=i, value=v)
        c.font = f_base
    for i, (k, h, f, fmt, w, g) in enumerate(CALC, start=len(RAW) + 1):
        formula = f.format(r=rr, **L)
        c = sc.cell(row=rr, column=i, value=formula)
        c.font = f_base
        if fmt:
            c.number_format = fmt
    sc.cell(row=rr, column=len(RAW) + [x[0] for x in CALC].index("reason") + 1).alignment = Alignment(vertical="top")

rng = lambda k: f"{L[k]}{FIRST}:{L[k]}{LAST}"
track_cf(sc, rng("track"))
for k in ("alarm", "fperf", "fserv", "fpri", "fdis", "flag", "urgent", "pb5"):
    sc.conditional_formatting.add(rng(k), CellIsRule(operator="equal", formula=["TRUE"],
                                                     fill=PatternFill("solid", fgColor="F4CCCC" if k != "pb5" else "C6E0B4")))
sc.freeze_panes = "C2"
sc.auto_filter.ref = f"A1:{get_column_letter(len(RAW) + len(CALC))}{LAST}"

# ================================================================ helper for list tabs
def scol(k):
    return f"Scoring!${L[k]}${FIRST}:${L[k]}${LAST}"

LIST_FIELDS = [
    ("Merchant", "id", 10, None), ("Name", "name", 18, None), ("Segment", "seg", 10, None),
    ("Track", "track", 15, None), ("Risk label", "risk", 10, None), ("Evidence sources (of 5)", "sources", 9, "0"),
    ("Main driver", "main", 13, None), ("Also flagged for", "also", 12, None), ("Playbook", "pb", 9, None),
    ("Playbook name", "pbname", 24, None), ("Value at stake (Rs)", "value", 12, "#,##0"),
    ("Days to renewal", "renew", 9, "0"), ("Urgent (skips queue)", "urgent", 9, None),
    ("Reason (plain language)", "reason", 80, None),
]

# ================================================================ Ranked list (all 100)
rl = wb.create_sheet("Ranked list", 2)
hdr(rl, 1, 1, "Rank", NAVY, 6)
hdr(rl, 1, 2, "Row in Scoring (helper)", "A0A7B4", 8)
for j, (h, k, w, fmt) in enumerate(LIST_FIELDS, start=3):
    hdr(rl, 1, j, h, NAVY, w)
rl.row_dimensions[1].height = 36
for n in range(1, N + 1):
    rr = n + 1
    rl.cell(row=rr, column=1, value=n).font = f_base
    rl.cell(row=rr, column=2, value=f"=MATCH(A{rr},{scol('rank')},0)").font = f_note
    for j, (h, k, w, fmt) in enumerate(LIST_FIELDS, start=3):
        c = rl.cell(row=rr, column=j, value=f"=INDEX({scol(k)},$B{rr})")
        c.font = f_base
        if fmt:
            c.number_format = fmt
track_cf(rl, f"F2:F{N + 1}")
rl.freeze_panes = "D2"
rl.auto_filter.ref = f"A1:{get_column_letter(2 + len(LIST_FIELDS))}{N + 1}"

# ================================================================ CSM top 15 / Light-touch
def list_tab(title, sheet_name, intro, start_rank_formula, count_formula, extra_cols):
    ws = wb.create_sheet(sheet_name)
    ws["A1"] = title
    ws["A1"].font = f_title
    ws["A2"] = intro
    ws["A2"].font = f_note
    ws["A3"] = "Merchants on this list:"
    ws["A3"].font = f_bold
    ws["C3"] = count_formula
    ws["C3"].font = f_bold
    hr = 5
    hdr(ws, hr, 1, "#", NAVY, 5)
    hdr(ws, hr, 2, "Rank", NAVY, 6)
    hdr(ws, hr, 3, "Row in Scoring (helper)", "A0A7B4", 8)
    fields = [f for f in LIST_FIELDS if f[1] != "track"]
    for j, (h, k, w, fmt) in enumerate(fields, start=4):
        hdr(ws, hr, j, h, NAVY, w)
    base = 4 + len(fields)
    for j, (h, w) in enumerate(extra_cols):
        hdr(ws, hr, base + j, h, "B7791F", w)
    ws.row_dimensions[hr].height = 36
    for n in range(1, 41):
        rr = hr + n
        ws.cell(row=rr, column=1, value=n).font = f_base
        ws.cell(row=rr, column=2, value=f'=IF(A{rr}<=$C$3,{start_rank_formula}+A{rr},"")').font = f_base
        ws.cell(row=rr, column=3, value=f'=IF(B{rr}="","",MATCH(B{rr},{scol("rank")},0))').font = f_note
        for j, (h, k, w, fmt) in enumerate(fields, start=4):
            c = ws.cell(row=rr, column=j, value=f'=IF($C{rr}="","",INDEX({scol(k)},$C{rr}))')
            c.font = f_base
            if fmt:
                c.number_format = fmt
        for j, (h, w) in enumerate(extra_cols):
            c = ws.cell(row=rr, column=base + j)
            c.font = f_base
            c.border = border
            if h.startswith("CSM approved"):
                c.fill, c.font = fill_input, f_input
    ws.freeze_panes = ws.cell(row=hr + 1, column=5)
    return ws, hr, base

csm, csm_hr, csm_base = list_tab(
    "CSM outreach list (top of the ranking, up to the cap)", "CSM top 15",
    "Urgent renewals first, then by priority score. Nothing is sent until a CSM approves it (yellow column).",
    "0", f'=COUNTIF({scol("track")},"CSM outreach")',
    [("Draft outreach message (from Messages tab; CSM reviews)", 70), ("CSM approved? (Y/N)", 12)])
dv_yn2 = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
csm.add_data_validation(dv_yn2)
dv_yn2.add(f"{get_column_letter(csm_base + 1)}{csm_hr + 1}:{get_column_letter(csm_base + 1)}{csm_hr + 40}")
MSG_FIRST = 5
MSG_LAST = MSG_FIRST + (len(MSGS["csm"]) if MSGS else 1) - 1
for n in range(1, 41):
    rr = csm_hr + n
    c = csm.cell(row=rr, column=csm_base,
                 value=f'=IF($C{rr}="","",IFERROR(VLOOKUP($D{rr},Messages!$A${MSG_FIRST}:$E${MSG_LAST},5,FALSE),'
                       f'"No draft yet: write one before outreach"))')
    c.font, c.alignment = f_base, wrap_top

lt, lt_hr, lt_base = list_tab(
    "Light-touch list (flagged, below the CSM cap)", "Light-touch",
    "Templated email, plus WhatsApp only if the merchant has opted in. A CSM approves each batch before sending.",
    "'CSM top 15'!$C$3", f'=COUNTIF({scol("track")},"Light-touch")',
    [("Draft message (template filled from the data; CSM reviews)", 70), ("Channel", 26),
     ("Escalate to CSM list next month if", 44), ("CSM approved? (Y/N)", 12)])
SIGN = "[Your name], Customer Success, Razorpay"
for n in range(1, 41):
    rr = lt_hr + n
    g = lambda k: f"INDEX({scol(k)},$C{rr})"
    perf = ('"Hi [contact name], "&IF(' + g("srchg") + '<=-SRDrop,"your payment success rate on Razorpay was "&TEXT('
            + g("sr2") + ',"0.0")&"% over the last three months, down from "&TEXT(' + g("sr1")
            + ',"0.0")&"% in the three months before. ","you recently told us about payment failures. ")'
            '&"We would like to help fix this. Reply to this email and our payments team will review failures by payment '
            'method and bank with you, then share a fix plan within 3 days of that review. ' + SIGN + '"')
    serv = ('"Hi [contact name], "&IF(' + g("unres") + '>=UnresolvedMin,"you have "&' + g("unres")
            + '&" support tickets still open with us. ","you told us our support has been slow. ")'
            '&"I am sorry this has taken longer than it should. I am now your single point of contact. Reply with a good '
            'time to talk and I will go through each open issue with you and confirm where it stands. ' + SIGN + '"')
    pri = ('"Hi [contact name], your Razorpay contract is due for renewal in "&' + g("renew")
           + '&" days. Before then, your account manager can walk you through your current rate card, volume tiers and '
             'renewal terms. Reply to this email to set up a time. ' + SIGN + '"')
    dis = ('"Hi [contact name], we want to make sure we are in touch with the right person at "&' + g("name")
           + '&" for your Razorpay account. Could you reply with the best contact for payments and reporting? We would also '
             'be glad to run a short session on the reports and products on your account. ' + SIGN + '"')
    main = g("main")
    c = lt.cell(row=rr, column=lt_base,
                value=f'=IF($C{rr}="","",IF({main}="Performance",{perf},IF({main}="Service",{serv},'
                      f'IF({main}="Pricing",{pri},{dis}))))')
    c.font, c.alignment = f_base, wrap_top
    lt.cell(row=rr, column=lt_base + 1, value=f'=IF($C{rr}="","","Email; WhatsApp only if opted in")').font = f_base
    lt.cell(row=rr, column=lt_base + 2,
            value=f'=IF($C{rr}="","","Merchant replies, signals worsen, or renewal within "&EscalateDays&" days")').font = f_base
dv_yn3 = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
lt.add_data_validation(dv_yn3)
dv_yn3.add(f"{get_column_letter(lt_base + 3)}{lt_hr + 1}:{get_column_letter(lt_base + 3)}{lt_hr + 40}")

# ================================================================ Messages
ms = wb.create_sheet("Messages")
ms["A1"] = "Outreach drafts for the CSM list, and light-touch templates"
ms["A1"].font = f_title
ms["A2"] = ("Drafted by Claude from each merchant's row. Every number is checked automatically against the data (Fact check "
            "column; after editing a draft, re-run it with scripts/recheck_drafts.py). Nothing is sent until a CSM "
            "reviews and approves it.")
ms["A2"].font = f_note
ms["A3"] = ("If the CSM list changes after you edit Settings, a new merchant shows 'No draft yet' on the CSM top 15 tab: "
            "add a row here for it.")
ms["A3"].font = f_note
mheads = [("Merchant", 10), ("Business name", 18), ("Main driver", 13), ("Subject", 30), ("Draft email", 80),
          ("Before sending, the CSM should", 45), ("Numbers used (all from the data)", 40), ("Fact check", 12),
          ("Words", 7), ("Drafted by", 14)]
for j, (h, w) in enumerate(mheads, start=1):
    hdr(ms, 4, j, h, NAVY, w)
ms.row_dimensions[4].height = 30
names_by_id = dict(zip(merchants["merchant_id"], merchants["business_name"]))
for i, m in enumerate(MSGS["csm"] if MSGS else []):
    rr = MSG_FIRST + i
    fc = FACTS.get(m["id"], {})
    vals = [m["id"], names_by_id[m["id"]], fc.get("driver", ""), m["subject"], m["body"], m["check"],
            fc.get("numbers", ""), fc.get("status", "Not checked"), fc.get("words", ""),
            "Claude draft, edited by you" if m.get("edited") else "Claude (AI draft)"]
    for j, v in enumerate(vals, start=1):
        c = ms.cell(row=rr, column=j, value=v)
        c.font, c.alignment, c.border = f_base, wrap_top, border
ms.conditional_formatting.add(f"H{MSG_FIRST}:H{MSG_LAST}", CellIsRule(operator="equal", formula=['"Passed"'],
                                                                       fill=PatternFill("solid", fgColor="C6E0B4")))
trow = MSG_LAST + 3
ms.cell(row=trow, column=1, value="Light-touch templates (filled automatically on the Light-touch tab)").font = f_sub
for j, (h, w) in enumerate([("Driver", 10), ("", 18), ("", 13), ("Template ({fields} come from the data)", 30), ("", 80),
                            ("Notes", 45)], start=1):
    if h:
        hdr(ms, trow + 1, j, h, NAVY)
for i, t in enumerate(MSGS["templates"] if MSGS else []):
    rr = trow + 2 + i
    ms.cell(row=rr, column=1, value=t["driver"]).font = f_base
    ms.merge_cells(start_row=rr, start_column=4, end_row=rr, end_column=5)
    c = ms.cell(row=rr, column=4, value=t["text"])
    c.font, c.alignment = f_base, wrap_top
    n_ = ms.cell(row=rr, column=6, value=t["notes"])
    n_.font, n_.alignment = f_base, wrap_top
    ms.row_dimensions[rr].height = 75
ms.freeze_panes = "B5"

# ================================================================ Test vs labels
ts = wb.create_sheet("Test vs labels")
labws = openpyxl.load_workbook(LAB, data_only=True)["Labelling"]
labels = []
for rr in range(3, 23):
    labels.append([labws.cell(rr, c).value for c in (1, 21, 22, 23, 24, 25, 26)])
ts["A1"] = "System vs your 20 hand labels"
ts["A1"].font = f_title
ts["A2"] = ("Your labels are copied as values from P2_Labelling.xlsx (9 Nov 2026). They are your judgement, not the hidden "
            "truth. The system columns recalculate when Settings change.")
ts["A2"].font = f_note
TH = 10
heads = [("Merchant", 10), ("Your call", 12), ("Your confidence", 10), ("Your driver", 13), ("Your not-at-risk type", 16),
         ("Your playbook", 9), ("Your why", 50), ("System call", 12), ("System driver", 13), ("System playbook", 9),
         ("System track", 15), ("Call matches?", 9), ("Driver matches? (churn only)", 11), ("Playbook matches?", 10)]
for j, (h, w) in enumerate(heads, start=1):
    hdr(ts, TH, j, h, NAVY if j <= 7 else GREEN_HDR, w)
ts.row_dimensions[TH].height = 36
for i, lab in enumerate(labels):
    rr = TH + 1 + i
    mid, call, conf, drv, typ, pb, why = lab
    for j, v in enumerate([mid, call, conf, drv, typ, pb, why], start=1):
        c = ts.cell(row=rr, column=j, value=v)
        c.font = f_base
        c.alignment = Alignment(vertical="top", wrap_text=(j == 7))
    m = f"MATCH($A{rr},{scol('id')},0)"
    ts.cell(row=rr, column=8, value=f'=IF(INDEX({scol("flag")},{m}),"Likely churn","Not at risk")')
    ts.cell(row=rr, column=9, value=f'=INDEX({scol("main")},{m})')
    ts.cell(row=rr, column=10, value=f'=INDEX({scol("pb")},{m})')
    ts.cell(row=rr, column=11, value=f'=INDEX({scol("track")},{m})')
    ts.cell(row=rr, column=12, value=f'=IF(B{rr}=H{rr},"Yes","No")')
    ts.cell(row=rr, column=13, value=f'=IF(AND(B{rr}="Likely churn",H{rr}="Likely churn"),IF(D{rr}=I{rr},"Yes","No"),"n/a")')
    ts.cell(row=rr, column=14, value=f'=IF(F{rr}=J{rr},"Yes","No")')
    for j in range(8, 15):
        ts.cell(row=rr, column=j).font = f_base
first_t, last_t = TH + 1, TH + len(labels)
for col in ("L", "M", "N"):
    rngt = f"{col}{first_t}:{col}{last_t}"
    ts.conditional_formatting.add(rngt, CellIsRule(operator="equal", formula=['"Yes"'], fill=PatternFill("solid", fgColor="C6E0B4")))
    ts.conditional_formatting.add(rngt, CellIsRule(operator="equal", formula=['"No"'], fill=PatternFill("solid", fgColor="F4CCCC")))
track_cf(ts, f"K{first_t}:K{last_t}")
summary_rows = [
    ("Calls that match your label", f'=COUNTIF(L{first_t}:L{last_t},"Yes")&" of "&COUNTA(A{first_t}:A{last_t})'),
    ("You: churn, system: churn (true positives vs your labels)",
     f'=COUNTIFS(B{first_t}:B{last_t},"Likely churn",H{first_t}:H{last_t},"Likely churn")'),
    ("You: churn, system: not at risk (missed)",
     f'=COUNTIFS(B{first_t}:B{last_t},"Likely churn",H{first_t}:H{last_t},"Not at risk")'),
    ("You: not at risk, system: churn (false alarms)",
     f'=COUNTIFS(B{first_t}:B{last_t},"Not at risk",H{first_t}:H{last_t},"Likely churn")'),
    ("Drivers that match (where both say churn)",
     f'=COUNTIF(M{first_t}:M{last_t},"Yes")&" of "&(COUNTIF(M{first_t}:M{last_t},"Yes")+COUNTIF(M{first_t}:M{last_t},"No"))'),
    ("Your churn labels that reach the CSM list",
     f'=COUNTIFS(B{first_t}:B{last_t},"Likely churn",K{first_t}:K{last_t},"CSM outreach")&" of "&COUNTIF(B{first_t}:B{last_t},"Likely churn")'),
    ("Playbooks that match", f'=COUNTIF(N{first_t}:N{last_t},"Yes")&" of "&COUNTA(A{first_t}:A{last_t})'),
]
for i, (label, f) in enumerate(summary_rows):
    ts.cell(row=3 + i, column=1, value=label).font = f_base
    c = ts.cell(row=3 + i, column=5, value=f)
    c.font = f_bold
ts.freeze_panes = ts.cell(row=TH + 1, column=2)
TEST_SUMMARY = {label: f"'Test vs labels'!$E${3 + i}" for i, (label, _) in enumerate(summary_rows)}

# ================================================================ Summary
sm = wb.create_sheet("Summary", 2)
sm["A1"] = "Summary"
sm["A1"].font = f_title
sm["A2"] = "Everything below recalculates from Settings."
sm["A2"].font = f_note
sm.column_dimensions["A"].width = 52
for col in "BCDE":
    sm.column_dimensions[col].width = 16
row = 4
def line(label, formula, fmt=None, bold=False, note=None):
    global row
    sm.cell(row=row, column=1, value=label).font = f_bold if bold else f_base
    c = sm.cell(row=row, column=2, value=formula)
    c.font = f_bold if bold else f_base
    if fmt:
        c.number_format = fmt
    if note:
        n = sm.cell(row=row, column=3, value=note)
        n.font = f_note
    row += 1
def section(text):
    global row
    row += 1
    sm.cell(row=row, column=1, value=text).font = f_sub
    row += 1

section("Merchants")
line("Merchants scored", f"=COUNTA({scol('id')})", "0")
line("Flagged (any driver)", f"=COUNTIF({scol('flag')},TRUE)", "0", True)
line("CSM outreach (within cap)", f'=COUNTIF({scol("track")},"CSM outreach")', "0")
line("   of which urgent renewals that skipped the queue", f'=COUNTIFS({scol("track")},"CSM outreach",{scol("urgent")},TRUE)', "0")
line("Light-touch (flagged, below cap)", f'=COUNTIF({scol("track")},"Light-touch")', "0")
line("Expansion candidates (PB-5, no slot used)", f'=COUNTIF({scol("track")},"Expansion (PB-5)")', "0")
line("Monitor only (PB-6)", f'=COUNTIF({scol("track")},"Monitor")', "0")
line("High risk label (flagged, 3+ evidence sources)", f'=COUNTIF({scol("risk")},"High")', "0")
line("Medium risk label (flagged, 1-2 sources)", f'=COUNTIF({scol("risk")},"Medium")', "0")
csm_count_ref = "B6"

section("Value at stake (12 months x 1% of latest TPV)")
r_val_csm = row
line("CSM list (Rs)", f'=SUMIF({scol("track")},"CSM outreach",{scol("value")})', "#,##0", True)
line("Light-touch list (Rs)", f'=SUMIF({scol("track")},"Light-touch",{scol("value")})', "#,##0")
r_val_all = row
line("All flagged merchants (Rs)", f"=SUMIF({scol('flag')},TRUE,{scol('value')})", "#,##0")
line("Share of flagged value covered by the CSM list", f"=IF(B{r_val_all}=0,0,B{r_val_csm}/B{r_val_all})", "0.0%")

section("Cost vs value of the CSM outreach")
r_cost = row
line("Cost of the CSM outreaches (Rs) = number of calls x cost per outreach",
     f'=COUNTIF({scol("track")},"CSM outreach")*FalseAlarmCost', "#,##0", True,
     "Worst case: if every call were a false alarm, this is all we lose.")
line("Break-even save rate (cost / value on the CSM list)", f"=IF(B{r_val_csm}=0,0,B{r_cost}/B{r_val_csm})", "0.00%",
     note="Share of the CSM list's value we must save to cover the calls.")
line("Scenario save rate (Settings, an assumption)", "=SaveRate", "0%")
r_exp = row
line("Expected value saved in the scenario (Rs)", f"=SaveRate*B{r_val_csm}", "#,##0")
line("Net value after outreach cost (Rs)", f"=B{r_exp}-B{r_cost}", "#,##0", True)

section("Flagged merchants by main driver")
for j, h in enumerate(["Driver", "Flagged", "CSM list", "Light-touch", "Value at stake (Rs)"], start=1):
    hdr(sm, row, j, h, NAVY)
row += 1
for d in ["Performance", "Service", "Pricing", "Disengagement"]:
    sm.cell(row=row, column=1, value=d).font = f_base
    sm.cell(row=row, column=2, value=f'=COUNTIF({scol("main")},"{d}")').font = f_base
    sm.cell(row=row, column=3, value=f'=COUNTIFS({scol("main")},"{d}",{scol("track")},"CSM outreach")').font = f_base
    sm.cell(row=row, column=4, value=f'=COUNTIFS({scol("main")},"{d}",{scol("track")},"Light-touch")').font = f_base
    c = sm.cell(row=row, column=5, value=f'=SUMIF({scol("main")},"{d}",{scol("value")})')
    c.font, c.number_format = f_base, "#,##0"
    row += 1

section("Check against your 20 hand labels")
for label, ref in TEST_SUMMARY.items():
    line(label, f"={ref}")

# ================================================================ Assumptions
asm = wb.create_sheet("Assumptions")
asm["A1"] = "Assumptions and decisions behind this sheet"
asm["A1"].font = f_title
for j, (h, w) in enumerate([("#", 5), ("Assumption / decision", 95), ("Source", 34)], start=1):
    hdr(asm, 3, j, h, NAVY, w)
items = [
    ("'Churn' means the merchant moves most of its volume away within the next 90 days.", "Data Pack, P2_Assumptions"),
    ("Value of a saved merchant = 12 months of net revenue at 1% of the latest month's TPV.", "Data Pack, P2_Assumptions"),
    ("Cost of one outreach / false alarm = Rs 1,500 (1.5 hours of CSM time).", "Data Pack, P2_Assumptions"),
    ("The cap of 15 applies to CSM-led outreach. Light-touch templated messages do not use it; a CSM approves each batch in minutes.", "Your decision"),
    ("WhatsApp is used only for merchants who have opted in; everyone else gets email. The P2 data has no opt-in field.", "Your decision (mirrors the brief's rule in Problem 3)"),
    ("No message is sent, and no action is taken, until a CSM has reviewed and approved it.", "Your decision; brief, ground rules (safety first)"),
    ("Flagged merchants renewing within 30 days (before next month's re-check) skip the queue into the CSM list. Those renewing in 31-60 days are caught at next month's review.", "Your decision"),
    ("A light-touch merchant moves into next month's CSM list if it replies, its signals worsen, or its renewal is within 60 days.", "Your decision"),
    ("TPV alarm line of -10.5% was tuned on this same data (it excludes C2094, a seasonal merchant at -10.1%).", "Your decision; say so in the pitch"),
    ("The TPV alarm is required only for the disengagement flag; the other three flags fire early, before volume moves.", "Your decision"),
    ("Missing CSAT is ignored: it never counts for or against a merchant.", "Your decision"),
    ("Expansion (PB-5) merchants use none of the 15 CSM slots.", "Your decision"),
    ("Savability weights, urgency boosts and likelihood multipliers: Claude proposed the starting numbers; you tested "
     "how much they move the CSM 15 and kept them on 9 November 2026. Editable in Settings.", "Your decision (starting numbers from Claude)"),
    ("Seasonal industries are Gifting, Education and Travel.", "Data Pack, P2_Assumptions"),
    ("Note tags were suggested by Claude from keyword rules. You reviewed all 22 against the merchants' numbers "
     "and kept them on 9 November 2026.", "AI-assisted (Claude); your review"),
    ("The 25% save rate on the Summary tab is a scenario, not a number from the data.", "Assumption for illustration"),
    ("CSM-list drafts were written by Claude from each merchant's row and checked automatically: every number matches the data, no other merchant is named, and there are no promises of discounts, waivers, refunds or dates. A CSM still reviews each one.", "AI-assisted (Claude); automated fact check; CSM review"),
    ("Light-touch messages are fixed templates filled by formulas from the data, so they cannot state a fact the data does not hold.", "Rules (Excel formulas)"),
    ("Messages use [contact name] and [Your name] placeholders: the data has no contact names.", "Data limitation"),
]
for i, (text, srcs) in enumerate(items, start=1):
    rr = 3 + i
    asm.cell(row=rr, column=1, value=i).font = f_base
    a = asm.cell(row=rr, column=2, value=text)
    a.font, a.alignment = f_base, wrap_top
    b = asm.cell(row=rr, column=3, value=srcs)
    b.font, b.alignment = f_base, wrap_top

# ================================================================ Read me
readme["A1"] = "P2 churn early-warning: scoring workbook"
readme["A1"].font = f_title
readme.column_dimensions["A"].width = 26
readme.column_dimensions["B"].width = 110
lines = [
    ("What this is", "Your churn rules for Problem 2, written as live Excel formulas over all 100 merchants. Change a setting and every tab updates."),
    ("How to use it", "1) Edit the yellow cells (blue numbers) on Settings. 2) Review the tags on Note tags. 3) Read the CSM top 15 and Light-touch tabs. 4) Check Test vs labels and Summary."),
    ("Colour legend", "Yellow cell with blue text = input you can change. Black text = formula, do not overwrite. Track colours: red = CSM outreach, amber = light-touch, green = expansion (PB-5), grey = monitor."),
    ("Example input", "Settings > 'Urgent renewal ... skip the queue' = 30 means flagged merchants renewing within 30 days (before next month's re-check) go to the top. Try 60 and watch the CSM list change."),
    ("", ""),
    ("Settings", "All thresholds, weights and the value/cost assumptions."),
    ("Scoring", "One row per merchant: raw data (grey headers), calculations (blue), flags (purple), outputs (green), helper text (light grey)."),
    ("Ranked list", "All 100 merchants in priority order with track, risk label, driver, playbook, value at stake and a plain-language reason."),
    ("Summary", "Counts by track, value at stake, cost vs value of the outreach, drivers, and agreement with your labels."),
    ("CSM top 15", "The merchants for CSM outreach this month, with their draft message (pulled from Messages) and a CSM approval column."),
    ("Light-touch", "Flagged merchants below the cap: a template message filled from each merchant's data, channel, escalation rule and approval column."),
    ("Messages", "The drafted email for each CSM-list merchant (AI-drafted, fact-checked against the data, CSM approves) and the four light-touch templates."),
    ("Test vs labels", "The system compared with your 20 hand labels (calls, drivers, playbooks)."),
    ("Note tags", "The 22 distinct account manager notes and their tags (Performance, Service, Pricing, Disengagement, Seasonal, Positive, Neutral)."),
    ("Playbooks", "The six CSM playbooks, copied from the Data Pack."),
    ("Assumptions", "Every assumption and decision, with its source."),
    ("", ""),
    ("How the logic works", "1) TPV alarm: latest month at or below the alarm line vs the merchant's own 6-month average. "
                            "2) Driver flags: Performance (success rate down 5+ points or performance note), Service (2+ unresolved or service note), "
                            "Pricing (pricing note and renewal within 60 days), Disengagement (TPV alarm and low logins or disengagement note). "
                            "3) Main driver order: Performance, Service, Pricing, Disengagement. "
                            "4) Priority score = value at stake ^ value weight x savability weight x urgency boost x likelihood (optional). "
                            "5) Urgent renewals skip the queue; the top ranks up to the cap go to CSM outreach, the rest of the flagged go light-touch."),
]
for i, (a, b) in enumerate(lines, start=3):
    ca = readme.cell(row=i, column=1, value=a)
    ca.font = f_bold
    cb = readme.cell(row=i, column=2, value=b)
    cb.font, cb.alignment = f_base, wrap_top
readme["A3"].fill = fill_input

# ================================================================ names, fonts, order
for name, ref in names.items():
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

order = ["Read me", "Settings", "Summary", "CSM top 15", "Light-touch", "Messages", "Ranked list", "Scoring", "Test vs labels",
         "Note tags", "Playbooks", "Assumptions"]
wb._sheets = [wb[n] for n in order]
for ws in wb.worksheets:
    for row_ in ws.iter_rows():
        for c in row_:
            f = c.font
            if f is None or f.name != FONT:
                c.font = Font(name=FONT, size=(f.size if f and f.size else 10), bold=(f.bold if f else False),
                              italic=(f.italic if f else False), color=(f.color if f else None))
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)

with open(COLMAP, "w") as fh:
    json.dump({"scoring": L, "first": FIRST, "last": LAST}, fh, indent=1)
print("Saved", OUT)
print("Scoring columns:", {k: L[k] for k in ("flag", "main", "rank", "track", "pb", "reason", "risk", "value", "score")})
