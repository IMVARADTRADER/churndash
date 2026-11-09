"""Styling and small HTML pieces. Colours follow razorpay.com (checked 9 Nov 2026); no Razorpay logo or marks."""
import html

import streamlit as st

BLUE, NAVY, MUTED, LINE, SOFT = "#305EFF", "#192839", "#5D6D86", "#DFE3E9", "#F4F7FB"

TRACK = {  # (background, text) for the four tracks
    "CSM outreach": ("#EAF0FF", "#2341C8"),
    "Light-touch": ("#FFF4DE", "#8A5A00"),
    "Expansion (PB-5)": ("#E6F5EE", "#00704A"),
    "Monitor": ("#F0F4F6", "#40566D"),
}
RISK = {"High": ("#FDECEA", "#B3261E"), "Medium": ("#FFF4DE", "#8A5A00"),
        "Low (watch)": ("#F0F4F6", "#40566D"), "Low": ("#F0F4F6", "#40566D")}

CSS = f"""
<style>
:root {{ --blue:{BLUE}; --navy:{NAVY}; --muted:{MUTED}; --line:{LINE}; --soft:{SOFT}; }}
.block-container {{ max-width: 1180px; padding-top: 4.6rem; }}
h1, h2, h3 {{ letter-spacing: -0.01em; color: var(--navy); }}
[data-testid="stHeading"] h1 {{ color: var(--blue) !important; }}
[data-testid="stMetricValue"] {{ font-variant-numeric: tabular-nums; }}

/* gentle entrance for each page, and calm hover states */
@keyframes rise {{ from {{ opacity: 0; transform: translateY(6px); }} to {{ opacity: 1; transform: none; }} }}
[data-testid="stMainBlockContainer"] > [data-testid="stVerticalBlock"] > * {{
  animation: rise .38s cubic-bezier(.2,.7,.2,1) both; }}
[data-testid="stMainBlockContainer"] > [data-testid="stVerticalBlock"] > :nth-child(4) {{ animation-delay: .04s; }}
[data-testid="stMainBlockContainer"] > [data-testid="stVerticalBlock"] > :nth-child(5) {{ animation-delay: .08s; }}
[data-testid="stMainBlockContainer"] > [data-testid="stVerticalBlock"] > :nth-child(n+6) {{ animation-delay: .12s; }}
[data-testid="stVerticalBlockBorderWrapper"], .card, .step {{ transition: box-shadow .2s ease, border-color .2s ease, transform .2s ease; }}
[class*="st-key-card"]:hover, .card:hover, .step:hover {{ box-shadow: 0 6px 24px rgba(25,40,57,.07); border-color: #C9D3F5 !important; }}
button, a {{ transition: background-color .18s ease, color .18s ease, box-shadow .18s ease, transform .18s ease !important; }}
[data-testid="stBaseButton-primary"]:hover {{ box-shadow: 0 4px 14px rgba(48,94,255,.28); transform: translateY(-1px); }}
@media (prefers-reduced-motion: reduce) {{
  *, *::before, *::after {{ animation: none !important; transition: none !important; scroll-behavior: auto !important; }}
}}

/* top bar */
.topbar {{ display:flex; align-items:center; justify-content:space-between; gap:1rem; flex-wrap:wrap;
  padding: .2rem 0 1rem; border-bottom: 1px solid var(--line); margin-bottom: .4rem; }}
.brand {{ font-family: 'TASA Orbiter', Inter, sans-serif; font-weight: 600; font-size: 1.15rem; color: var(--navy); }}
.brand b {{ color: var(--blue); font-weight: 600; }}
.brand .sub {{ font-family: Inter, sans-serif; font-weight: 500; font-size: .9rem; color: var(--muted); }}
.proto {{ font-size: .78rem; color: #8A5A00; background: #FFF4DE; border: 1px solid #F3DDB0; border-radius: 999px;
  padding: .22rem .7rem; white-space: nowrap; }}

/* pills, cards, steps */
.pill {{ display:inline-block; font-size:.78rem; font-weight:600; border-radius:999px; padding:.18rem .62rem; margin:0 .3rem .3rem 0; }}
.eyebrow {{ font-size:.78rem; font-weight:600; letter-spacing:.06em; text-transform:uppercase; color: var(--muted); margin-bottom:.2rem; }}
.lead {{ font-size:1.08rem; color: var(--muted); max-width: 46rem; line-height:1.55; }}
.card {{ border:1px solid var(--line); border-radius:14px; padding:1rem 1.1rem; background:#fff; height:100%; }}
.card h4 {{ margin:0 0 .35rem; font-size:1rem; }}
.card p {{ margin:0; color: var(--muted); font-size:.92rem; line-height:1.5; }}
.steps {{ display:grid; grid-template-columns: repeat(auto-fill, minmax(250px, 1fr)); gap:.75rem; margin-top:.4rem; }}
.step {{ border:1px solid var(--line); border-radius:14px; padding:.9rem 1rem; background:#fff; }}
.step .n {{ display:inline-flex; width:1.6rem; height:1.6rem; border-radius:50%; align-items:center; justify-content:center;
  background: var(--soft); color: var(--blue); font-weight:700; font-size:.85rem; margin-right:.45rem; }}
.step .t {{ font-weight:600; color: var(--navy); }}
.step p {{ margin:.45rem 0 0; color: var(--muted); font-size:.9rem; line-height:1.45; }}
.step .tier {{ float:right; font-size:.72rem; font-weight:600; border-radius:999px; padding:.12rem .5rem; }}
.badge-ok {{ display:inline-flex; align-items:center; gap:.5rem; background:#E6F5EE; color:#00704A; border:1px solid #BFE6D3;
  border-radius:999px; padding:.35rem .9rem; font-weight:600; font-size:.92rem; }}
.badge-bad {{ display:inline-flex; align-items:center; gap:.5rem; background:#FDECEA; color:#B3261E; border:1px solid #F5C2BD;
  border-radius:999px; padding:.35rem .9rem; font-weight:600; font-size:.92rem; }}
.badge-na {{ display:inline-flex; align-items:center; gap:.5rem; background: var(--soft); color: var(--muted); border:1px solid var(--line);
  border-radius:999px; padding:.35rem .9rem; font-weight:500; font-size:.92rem; }}
.ev {{ display:grid; grid-template-columns: repeat(auto-fill, minmax(170px, 1fr)); gap:.6rem; }}
.ev div {{ border:1px solid var(--line); border-radius:12px; padding:.7rem .8rem; font-size:.88rem; color: var(--muted); }}
.ev div.on {{ border-color:#F5C2BD; background:#FFF8F7; }}
.ev b {{ display:block; color: var(--navy); margin-bottom:.15rem; }}
.ev .dot {{ display:inline-block; width:.5rem; height:.5rem; border-radius:50%; margin-right:.35rem; background:#C4CCD6; }}
.ev .on .dot {{ background:#D52B1E; }}
.formula {{ font-variant-numeric: tabular-nums; color: var(--muted); font-size:.92rem; line-height:1.9; }}
.formula b {{ color: var(--navy); }}
.reason {{ font-size:1.12rem; line-height:1.5; color: var(--navy); }}
/* the one call to action on the overview */
.st-key-cta {{ background: linear-gradient(135deg, #EEF3FF 0%, #F7F9FF 100%); border: 1px solid #C9D6FF !important;
  border-radius: 16px !important; padding: .4rem .6rem .8rem; }}
.st-key-cta [data-testid="stBaseButton-primary"] {{ min-height: 2.9rem; font-weight: 600; font-size: 1rem; }}
.cta-eyebrow {{ font-size:.75rem; font-weight:700; letter-spacing:.08em; text-transform:uppercase; color: var(--blue); }}
.cta-title {{ font-family:'TASA Orbiter', Inter, sans-serif; font-size:1.45rem; font-weight:600; color: var(--navy); margin:.15rem 0 .2rem; }}
.cta-text {{ color: var(--muted); margin:0 0 .3rem; font-size:.95rem; }}

/* what each tag does */
table.fx {{ width:100%; border-collapse:separate; border-spacing:0; border:1px solid var(--line); border-radius:12px;
  overflow:hidden; font-size:.9rem; }}
table.fx th {{ text-align:left; font-weight:600; color: var(--muted); background: var(--soft); padding:.55rem .8rem; }}
table.fx td {{ padding:.5rem .8rem; border-top:1px solid #EEF1F5; color: var(--navy); vertical-align:middle; line-height:1.45; }}
table.fx .n {{ text-align:right; white-space:nowrap; font-variant-numeric: tabular-nums; }}
table.fx td .pill {{ margin:0; }}

/* journey stepper on each step page */
.journey {{ display:flex; gap:.4rem; flex-wrap:wrap; margin:.2rem 0 .6rem; }}
.journey .j {{ display:flex; align-items:center; gap:.45rem; font-size:.84rem; color: var(--muted); padding:.3rem .75rem .3rem .35rem;
  border:1px solid var(--line); border-radius:999px; background:#fff; }}
.journey .j .d {{ width:1.35rem; height:1.35rem; border-radius:50%; display:inline-flex; align-items:center; justify-content:center;
  font-size:.75rem; font-weight:700; background: var(--soft); color: var(--muted); }}
.journey .j.done .d {{ background:#E6F5EE; color:#00704A; }}
.journey .j.todo-act {{ border-color:#F3DDB0; background:#FFF9EC; color:#8A5A00; }}
.journey .j.todo-act .d {{ background:#FFE9B8; color:#8A5A00; }}
.journey .j.now {{ border-color: var(--blue); color: var(--navy); font-weight:600; box-shadow: 0 0 0 3px rgba(48,94,255,.10); }}
.journey .j.now .d {{ background: var(--blue); color:#fff; }}
.journey .sep {{ align-self:center; width:14px; height:1px; background: var(--line); }}
.foot {{ color: var(--muted); font-size:.8rem; border-top:1px solid var(--line); margin-top:2.5rem; padding-top:.8rem; }}
</style>
"""


def esc(s) -> str:
    return html.escape(str(s))


def pill(text: str, colours: tuple[str, str]) -> str:
    bg, fg = colours
    return f'<span class="pill" style="background:{bg};color:{fg}">{esc(text)}</span>'


def topbar():
    st.html(CSS)
    st.html('<div class="topbar"><div class="brand">Churn<b>Dash</b> <span class="sub">· merchant churn early-warning</span></div>'
            '<div class="proto">Student build-a-thon prototype · not a Razorpay product</div></div>')


def footer():
    st.html('<div class="foot">Prototype built for the Razorpay Campus AI Build-a-thon (Problem 2) by a student entrant, '
            'with Claude (AI) writing the code. Synthetic data only. The Excel workbook is the source of truth; '
            'this app runs the same rules. Nothing here sends a message.</div>')
