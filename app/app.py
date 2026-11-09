"""ChurnDash: merchant churn early-warning, a Streamlit prototype for Razorpay Build-a-thon Problem 2.

Run from the project folder:  python -m streamlit run app/app.py
Pages follow the CSM's process: Overview -> 1 Upload -> 2 Tag notes -> 3 Ranking -> 4 Merchant (draft + approve).
Tier 2: AI suggests tags (DeepSeek by default, Jev as the switch) and DeepSeek drafts; every draft is fact-checked live.
"""
import hashlib
import os
from datetime import datetime

import altair as alt
import pandas as pd
import streamlit as st

import ai
import core
import look
from look import esc, pill

st.set_page_config(page_title="ChurnDash (student prototype)", page_icon=":material/radar:", layout="wide")

# On a host, keys come from its secrets store; locally, scripts/apikeys.py reads .env. Keys are never shown.
try:
    for _name in ai.KEY_NAMES.values():
        if _name in st.secrets and not os.environ.get(_name):
            os.environ[_name] = str(st.secrets[_name])
except Exception:  # no secrets file: fine locally
    pass

ss = st.session_state
ss.setdefault("upload", None)       # dict(name, data, checked) for the file in use; memory only
ss.setdefault("confirmed", {})      # note text -> tag a person confirmed this session
ss.setdefault("tag_log", {})        # note text -> how its tag was reached (AI suggestion kept or changed)
ss.setdefault("suggest", {})        # note text -> AI suggestion
ss.setdefault("scored", None)       # cached result for (file, tags)
ss.setdefault("approvals", {})      # (file digest, merchant) -> approved message
ss.setdefault("selected", None)     # merchant shown on the merchant page
ss.setdefault("ai_drafts", {})      # (file digest, merchant) -> DeepSeek draft
ss.setdefault("checks", {})         # hash of (file, merchant, text) -> fact-check verdict
ss.setdefault("ai_calls", 0)        # AI calls this session (cap: ai.SESSION_CAP)
ESCALATE_DAYS = 60                  # workbook Settings: EscalateDays


def keep_one(key: str, fallback: str):
    """on_change for a single-choice segmented control: clicking the chosen option again keeps it chosen."""
    def cb():
        if ss[key] is None:
            ss[key] = ss.get(f"{key}__last", fallback)
        ss[f"{key}__last"] = ss[key]
    ss.setdefault(f"{key}__last", ss.get(key) or fallback)
    return cb


def first_default(key: str, value):
    """Pass a widget's default only on first render; later the value lives in session state (no Streamlit warning)."""
    return {} if key in ss else {"default": value}


def ai_budget(n: int):
    if ss.ai_calls + n > ai.SESSION_CAP:
        raise ai.AIError(f"This session's AI limit ({ai.SESSION_CAP} calls) is used up. Tag or edit by hand.")
    ss.ai_calls += n


@st.cache_resource
def workbook():
    return core.workbook_snapshot()


@st.cache_resource
def playbooks():
    return core.playbooks_table()


def file_state():
    """'none', 'error', 'needs_tags' or 'ready', plus the list of untagged notes."""
    up = ss.upload
    if not up:
        return "none", []
    if not up["checked"].ok:
        return "error", []
    todo = core.pending_notes(up["checked"].df, ss.confirmed)
    return ("needs_tags" if todo else "ready"), todo


def load(name: str, data: bytes):
    checked = core.check_file(data)
    ss.upload = dict(name=name, data=data, checked=checked)
    ss.scored, ss.selected = None, None


def scored():
    """Score once per (file, confirmed tags) and keep it in this session only."""
    up = ss.upload
    df = up["checked"].df
    used = {n: ss.confirmed[n] for n in core.unseen_notes(df)}
    key = (up["checked"].digest, tuple(sorted(used.items())))
    if ss.scored and ss.scored["key"] == key:
        return ss.scored
    out, order = core.score(up["data"], used)
    table = core.explain(df, out, order, playbooks())
    wb_raw, wb_ranks, wb_tracks = workbook()
    same = core.same_data(df, wb_raw)
    ss.scored = dict(key=key, out=out, order=order, table=table, same=same,
                     match=core.compare_ranks(out, wb_ranks, wb_tracks) if same else None, used=used)
    return ss.scored


def badge(res):
    m = res["match"]
    if m is None:
        st.html('<span class="badge-na">Different data from the submitted workbook, so there are no saved ranks to '
                'compare against</span>')
    elif m["ranks"] == m["n"]:
        st.html(f'<span class="badge-ok">✓ Matches the submitted workbook: {m["ranks"]} of {m["n"]} ranks</span>')
    else:
        st.html(f'<span class="badge-bad">✕ Differs from the submitted workbook: {m["ranks"]} of {m["n"]} ranks match '
                f'({", ".join(m["differ"][:6])})</span>')


def journey(now: int):
    """The four steps, so the user always sees where they are and what's left."""
    state, todo = file_state()
    n_unseen = len(core.unseen_notes(ss.upload["checked"].df)) if state in ("needs_tags", "ready") else 0
    tag_label = ("Tag notes" if state in ("none", "error") else f"Tag notes · {len(todo)} to confirm" if todo
                 else "Tag notes · all recognised" if not n_unseen else f"Tag notes · {n_unseen} confirmed")
    steps = [("Upload", "done" if state in ("needs_tags", "ready") else "todo-act" if state == "error" else ""),
             (tag_label, "done" if state == "ready" else "todo-act" if state == "needs_tags" else ""),
             ("Ranking", "done" if state == "ready" and ss.scored else ""),
             ("Merchant, draft and approval", "done" if ss.approvals else "")]
    parts = []
    for i, (label, cls) in enumerate(steps, start=1):
        cls = (cls + " now").strip() if i == now else cls
        mark = "✓" if "done" in cls and i != now else str(i)
        parts.append(f'<div class="j {cls}"><span class="d">{mark}</span>{esc(label)}</div>')
    st.html('<div class="journey">' + '<span class="sep"></span>'.join(parts) + "</div>")


def need_file(msg="Upload a merchant file first."):
    st.info(msg, icon=":material/upload_file:")
    st.page_link(P_UPLOAD, label="Go to upload", icon=":material/arrow_forward:")


# ================================================================ Overview and downloads
def overview():
    st.html('<div class="eyebrow">Problem 2 · merchant churn early warning</div>')
    st.title("Who to call first, and why")
    st.html('<p class="lead">My rules flag merchants at risk of moving their volume away, name the likely reason, '
            'and rank who a CSM should contact this month (cap: 15). Every message waits for a person to approve it, '
            'and nothing is ever sent from here.</p>')

    with st.container(border=True, key="cta"):
        st.html('<div class="cta-eyebrow">Try it live</div>'
                '<div class="cta-title">Run the rules on a merchant file</div>'
                '<p class="cta-text">Upload → tag notes → ranked list → merchant page, draft and approval. '
                'About two minutes. Synthetic data only.</p>')
        b1, b2, _ = st.columns([1.25, 1.6, 2])
        if b1.button("Upload a merchant file", type="primary", icon=":material/upload_file:", width="stretch"):
            st.switch_page(P_UPLOAD)
        if core.ORIGINAL_DATA.exists() and b2.button("Or use the original file (demo)", icon=":material/bolt:",
                                                      width="stretch"):
            load(core.ORIGINAL_DATA.name, core.ORIGINAL_DATA.read_bytes())
            st.switch_page(P_TAGS)

    st.subheader("Download the submission")
    c1, c2 = st.columns(2)
    with c1, st.container(border=True, key="card_wb"):
        st.markdown("**Scoring workbook** · `P2_Churn_Scoring.xlsx`")
        st.caption("The working solution and source of truth: live Excel formulas, Settings you can edit, "
                   "the CSM 15, light-touch list, drafts and the test against my labels.")
        if core.WORKBOOK.exists():
            st.download_button("Download workbook", core.WORKBOOK.read_bytes(), core.WORKBOOK.name,
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               icon=":material/download:")
        else:
            st.caption("Not bundled with this copy of the app.")
    with c2, st.container(border=True, key="card_pdf"):
        st.markdown("**Written answers** · PDF")
        st.caption("The 11 written answers (100–150 words each), including how the rules were tested "
                   "and where they fail.")
        if core.ANSWERS_PDF.exists():
            st.download_button("Download written answers", core.ANSWERS_PDF.read_bytes(), core.ANSWERS_PDF.name,
                               "application/pdf", icon=":material/download:")
        else:
            st.caption("Not bundled with this copy of the app.")

    st.subheader("How it works")
    steps = [
        ("Upload", "Checks the sheet, columns and types. Rejects a wrong file with a clear reason. Synthetic data only; "
                   "nothing is stored.", "Built"),
        ("Tag notes", "Notes from my reviewed table get my tag. For an unseen note, DeepSeek (or Jev) suggests a tag "
                      "and a person confirms it, so a Pricing flag can't drop silently.", "Built"),
        ("Score and rank", "Runs scripts/reference.py unchanged: the same rules as the workbook. Shows a badge when "
                           "the ranks match the submitted workbook.", "Built"),
        ("Ranked list", "All merchants, with the CSM 15 and the light-touch list highlighted.", "Built"),
        ("Merchant page", "Why it's flagged, the evidence behind it, the priority maths and the playbook.", "Built"),
        ("Draft", "The 15 stored drafts, a DeepSeek email and WhatsApp message written from that merchant's row and "
                  "playbook only, or the rules template. The system prompt is on the page.", "Built"),
        ("Fact check", "scripts/factcheck.py runs live on every draft and every edit: numbers, other merchants' names, "
                       "promise words, driver fit. A failed draft is blocked.", "Built"),
        ("CSM edits and approves", "A person edits and approves. WhatsApp only with an opt-in toggle. "
                                   "Nothing is sent, ever.", "Built"),
    ]
    tiers = {"Built": ("#E6F5EE", "#00704A"), "Tier 2": ("#F0F4F6", "#5D6D86")}
    cells = "".join(
        f'<div class="step"><span class="tier" style="background:{tiers[t][0]};color:{tiers[t][1]}">{t}</span>'
        f'<span class="n">{i}</span><span class="t">{esc(a)}</span><p>{esc(b)}</p></div>'
        for i, (a, b, t) in enumerate(steps, start=1))
    st.html(f'<div class="steps">{cells}</div>')


# ================================================================ 1. Upload
def upload_page():
    journey(1)
    st.title("Upload merchant data")
    st.warning("**Synthetic data only.** Never upload real merchant data: the build-a-thon brief forbids it. "
               "The file is read in memory for your session and isn't saved anywhere.", icon=":material/shield:")

    def on_upload():  # runs before the page redraws, so the steps above show the new file at once
        f = ss.uploader
        if f is not None and ss.get("last_file_id") != f.file_id:  # only a newly chosen file replaces the one in use
            ss.last_file_id = f.file_id
            load(f.name, f.getvalue())

    st.file_uploader("Excel file with a `P2_Merchants` sheet (same columns as the Data Pack)", type=["xlsx"],
                     key="uploader", on_change=on_upload)
    if core.ORIGINAL_DATA.exists():
        st.button("Use the original Problem Statement 2.xlsx from this folder (local demo)",
                  icon=":material/folder_open:",
                  on_click=lambda: load(core.ORIGINAL_DATA.name, core.ORIGINAL_DATA.read_bytes()))

    up = ss.upload
    if not up:
        st.caption("Required columns: " + ", ".join(f"`{c}`" for c in core.REQUIRED) +
                   ". CSAT and the account manager note may be blank.")
        return
    chk = up["checked"]
    st.markdown(f"**File:** {esc(up['name'])}")
    if not chk.ok:
        st.error("This file can't be scored. Fix it and upload it again:\n\n" +
                 "\n".join(f"- {e}" for e in chk.errors), icon=":material/error:")
        return
    df = chk.df
    unseen = core.unseen_notes(df)
    st.success(f"Checked: {len(df)} merchants, all required columns present, types OK.", icon=":material/check_circle:")
    for w in chk.warnings:
        st.caption(f"Note: {w}")
    nt = core.note_table(df)
    known = len(nt) - len(unseen)
    c1, c2, c3 = st.columns(3)
    c1.metric("Merchants", len(df))
    c2.metric("Notes in my reviewed table", known)
    c3.metric("Unseen notes", len(unseen))
    state, todo = file_state()
    label = (f"Next: tag {len(todo)} unseen note{'s' if len(todo) > 1 else ''}" if todo
             else "Next: check the note tags")
    if st.button(label, type="primary", icon=":material/arrow_forward:"):
        st.switch_page(P_TAGS)


# ================================================================ 2. Tag notes
def tags_page():
    journey(2)
    st.title("Tag account manager notes")
    up = ss.upload
    if not up or not up["checked"].ok:
        return need_file()
    df = up["checked"].df
    nt = core.note_table(df)
    unseen = core.unseen_notes(df)
    st.html("<p class='lead'><b>Why this step:</b> the rules can't read free text. Each account manager note is "
            "turned into one of seven tags, and the rules use the tag. A Pricing note is the only way a Pricing flag "
            "can fire, so a note the rules can't read would silently drop that merchant from the list.</p>")

    # what each tag does, and how many merchants in this file carry it
    tag_of = {**ss.confirmed, **core.REVIEWED_TAGS}
    per_tag = df[core.NOTE_COL].map(lambda n: tag_of.get(n, "Not tagged yet") if isinstance(n, str) else None)
    counts = per_tag.value_counts()
    rows = "".join(
        f'<tr><td>{pill(t, ("#FDECEA", "#B3261E") if t in core.RISK_TAGS else ("#F0F4F6", "#40566D"))}</td>'
        f'<td>{esc(core.TAG_EFFECT[t])}</td><td class="n">{int(counts.get(t, 0))}</td></tr>'
        for t in core.TAG_CHOICES)
    waiting = int(counts.get("Not tagged yet", 0))
    if waiting:
        rows += (f'<tr><td>{pill("Not tagged yet", ("#FFF4DE", "#8A5A00"))}</td><td>Waiting for a person below. '
                 f'Until then these notes count for nothing, so scoring is held.</td><td class="n">{waiting}</td></tr>')
    st.html('<table class="fx"><tr><th>Tag</th><th>What it does in the rules</th><th class="n">Merchants in this '
            f'file</th></tr>{rows}</table>')
    st.caption("In the tests, rewording every note without this step dropped all 6 Pricing flags and changed 6 of "
               "the CSM 15. With the step, a person confirms the new tags and the same 15 come back.")

    source = {n: ("My reviewed table (9 Nov)" if n in core.REVIEWED_TAGS else
                  (f"Confirmed by a person ({ss.tag_log[n]['by']} suggestion kept)" if ss.tag_log.get(n, {}).get("kept")
                   else "Confirmed by a person") if n in ss.confirmed else "Waiting for a person")
              for n in nt["note"]}
    with st.expander(f"All {len(nt)} notes in this file and their tags", icon=":material/list:",
                     expanded=not unseen):
        st.dataframe(pd.DataFrame({"Note": nt["note"], "Merchants": nt["merchants"],
                                   "Tag": [tag_of.get(n, "") for n in nt["note"]],
                                   "Where the tag came from": [source[n] for n in nt["note"]]}),
                     hide_index=True, width="stretch")

    if not unseen:
        st.success(f"All {len(nt)} notes in this file are in my reviewed table, so there's nothing to confirm. "
                   "Notes with new wording would appear here for a person to tag.", icon=":material/check_circle:")
        if st.button("Next: see the ranking", type="primary", icon=":material/arrow_forward:"):
            st.switch_page(P_RANK)
        return

    st.subheader(f"{len(unseen)} unseen note{'s' if len(unseen) > 1 else ''}")
    pick_key = {n: "pick_" + hashlib.md5(n.encode("utf-8")).hexdigest()[:10] for n in unseen}
    c1, c2 = st.columns([1, 2], vertical_alignment="bottom")
    tagger = c1.segmented_control(
        "AI suggestions from", ["DeepSeek", "Jev"], key="tagger", **first_default("tagger", "DeepSeek"),
        on_change=keep_one("tagger", "DeepSeek"), help="DeepSeek (deepseek-flash) is the default: 94 of 94 in the tagging test. Jev (TypeSafe) is the tested "
             "alternative: 91 of 94, and it gives a confidence score.") or "DeepSeek"
    todo = [n for n in unseen if not ss.suggest.get(n, {}).get("tag") or ss.suggest[n]["by"] != tagger]
    if c2.button(f"Suggest tags with {tagger}" + (f" ({len(todo)} notes)" if todo else ""), icon=":material/auto_awesome:",
                 disabled=not todo or not ai.available(tagger)):
        try:
            ai_budget(len(todo))
            with st.spinner(f"Asking {tagger} about {len(todo)} notes…"):
                got = ai.suggest_tags(todo, tagger)
            ss.suggest.update(got)
            for n, s in got.items():  # pre-select the suggestion; the person still confirms below
                if s["tag"]:
                    ss[pick_key[n]] = s["tag"]
            st.rerun()
        except ai.AIError as e:
            st.error(str(e), icon=":material/cloud_off:")
    if not ai.available(tagger):
        st.caption(f"No {tagger} key is set, so pick each tag yourself.")
    st.caption("AI only suggests. Nothing is scored until a person confirms every tag below. The prompt is the one "
               "from the 94-note tagging test (scripts/test_tagging.py).")
    with st.expander("What each tag means (the definitions the AI is given)", icon=":material/help:"):
        for t in core.TAG_CHOICES:
            st.markdown(f"**{t}:** {core.DEFINITIONS[t]}")

    ids_by_note = df.groupby(core.NOTE_COL)["merchant_id"].apply(list).to_dict()
    with st.form("tags"):
        picks = {}
        for note in unseen:
            k = pick_key[note]
            if k not in ss and ss.confirmed.get(note):
                ss[k] = ss.confirmed[note]
            with st.container(border=True):
                st.markdown(f"“{esc(note)}”")
                st.caption("Merchants: " + ", ".join(ids_by_note.get(note, [])))
                s = ss.suggest.get(note)
                if s and s["tag"]:
                    conf = s.get("confidence")
                    unsure = isinstance(conf, (int, float)) and conf < 0.6
                    st.html(pill(f"{s['by']} suggests {s['tag']}" +
                                 (f" · confidence {conf:.2f}" if isinstance(conf, (int, float)) else ""),
                                 ("#FFF4DE", "#8A5A00") if unsure else ("#EAF0FF", "#2341C8")) +
                            (pill("Please check", ("#FDECEA", "#B3261E")) if unsure else ""))
                elif s and s.get("error"):
                    st.caption(f"{s['by']} couldn't suggest a tag ({s['error']}). Pick one yourself.")
                picks[note] = st.selectbox("Tag", core.TAG_CHOICES, key=k, placeholder="Choose a tag",
                                           label_visibility="collapsed", **({} if k in ss else {"index": None}))
        done = st.form_submit_button("Confirm tags and score", type="primary", icon=":material/check:")
    if done:
        missing = [n for n, t in picks.items() if not t]
        if missing:
            st.error(f"{len(missing)} note{'s are' if len(missing) > 1 else ' is'} still untagged. Tag every note "
                     "before scoring.", icon=":material/error:")
        else:
            for n, t in picks.items():
                s = ss.suggest.get(n) or {}
                ss.tag_log[n] = dict(tag=t, suggested=s.get("tag"), by=s.get("by"),
                                     kept=bool(s.get("tag")) and s.get("tag") == t)
            ss.confirmed.update(picks)
            ss.scored = None
            st.switch_page(P_RANK)


# ================================================================ 3. Ranking
def rank_page():
    journey(3)
    st.title("Ranked list")
    state, todo = file_state()
    if state in ("none", "error"):
        return need_file("Upload a valid merchant file first." if state == "error" else "Upload a merchant file first.")
    if state == "needs_tags":
        st.warning(f"**{len(todo)} unseen note{'s' if len(todo) > 1 else ''} in this file.** Scoring waits until a "
                   "person tags them. Scoring now would treat them as untagged and silently drop any Pricing flag.",
                   icon=":material/pending_actions:")
        st.page_link(P_TAGS, label=f"Tag {len(todo)} note{'s' if len(todo) > 1 else ''}",
                     icon=":material/arrow_forward:")
        return

    res = scored()
    t = res["table"]
    badge(res)
    if res["used"]:
        logs = [ss.tag_log.get(n, {}) for n in res["used"]]
        kept = sum(bool(lg.get("kept")) for lg in logs)
        changed = sum(bool(lg.get("suggested")) and not lg.get("kept") for lg in logs)
        st.caption(f"Includes {len(res['used'])} note tag{'s' if len(res['used']) > 1 else ''} confirmed by a person "
                   f"in step 2: {kept} AI suggestion{'s' if kept != 1 else ''} kept, {changed} changed, "
                   f"{len(logs) - kept - changed} picked without a suggestion.")
    counts = t.track.value_counts()
    flagged = t[t.track.isin(["CSM outreach", "Light-touch"])]
    csm = t[t.track == "CSM outreach"]
    cols = st.columns(5)
    cols[0].metric("Flagged", len(flagged))
    cols[1].metric("CSM outreach", int(counts.get("CSM outreach", 0)), help="Cap: 15 this month")
    cols[2].metric("Light-touch", int(counts.get("Light-touch", 0)), help="Templated email; no CSM slot used")
    cols[3].metric("CSM list value (Rs lakh)", f"{csm.value.sum() / 1e5:.1f}",
                   help="Value at stake: 12 months × 1% net revenue × month-6 TPV, summed over the CSM list")
    cols[4].metric("CSM call cost (Rs)", f"{len(csm) * 1500:,}", help="Rs 1,500 per outreach (from the assumptions)")

    names = {"csm": f"CSM list ({len(csm)})", "lt": f"Light-touch ({int(counts.get('Light-touch', 0))})",
             "all": f"All {len(t)}"}
    tracks = {"csm": ["CSM outreach"], "lt": ["Light-touch"], "all": list(look.TRACK)}
    pick = st.segmented_control("Show", list(names), format_func=names.get, key="show",
                                **first_default("show", "csm"), on_change=keep_one("show", "csm"),
                                label_visibility="collapsed") or "csm"
    view = t[t.track.isin(tracks[pick])]
    digest = ss.upload["checked"].digest
    shown = pd.DataFrame({
        "Rank": view["rank"], "Merchant": view.merchant_id, "Name": view.business_name,
        "Risk": view.risk, "Evidence": view.sources.map(lambda s: f"{s} of 5"),
        "Driver": view.main.replace("", "–"), "Reason": view.reason,
        "Value (Rs lakh)": view.value / 1e5, "Renewal (days)": view.renewal, "Urgent": view.urgent,
        "Playbook": view.playbook, "Track": view.track,
        "CSM": [("Approved" if (digest, m) in ss.approvals else "To review") if tr in ("CSM outreach", "Light-touch")
                else "" for m, tr in zip(view.merchant_id, view.track)],
    })

    def tint(row):
        bg, _ = look.TRACK[row["Track"]]
        return [f"background-color: {bg}" if row["Track"] in ("CSM outreach", "Light-touch") else ""] * len(row)

    ev = st.dataframe(
        shown.style.apply(tint, axis=1).format({"Value (Rs lakh)": "{:.2f}"}),
        hide_index=True, width="stretch", height=min(3 + 33 * (len(shown) + 1), 640),
        on_select="rerun", selection_mode=["single-row", "single-cell"], key=f"tbl_{pick}",
        column_config={
            "Rank": st.column_config.NumberColumn(width=55),
            "Merchant": st.column_config.TextColumn(width=80),
            "Reason": st.column_config.TextColumn(width=420),
            "Urgent": st.column_config.CheckboxColumn(width=65, help="Renews within 30 days: skips the queue"),
            "Value (Rs lakh)": st.column_config.NumberColumn(help="Value at stake: 12 × 1% × month-6 TPV"),
            "Evidence": st.column_config.TextColumn(help="How many of 5 sources see a problem: volume, payments, "
                                                        "support, engagement, account manager note"),
        })
    st.caption("Click any merchant to open its page. Urgent renewals (within 30 days) go first; then priority "
               "score = value × savability weight × renewal urgency × likelihood.")
    sel = ev.selection
    row = (sel.rows[0] if sel.rows else sel.cells[0][0] if getattr(sel, "cells", None) else None)
    if row is not None:  # a click on the row's tick box or on any cell opens that merchant
        ss.selected = view.iloc[row].merchant_id
        st.switch_page(P_MERCHANT)

    export = t[["rank", "merchant_id", "business_name", "track", "risk", "sources", "main", "also", "playbook",
                "playbook_name", "value", "renewal", "urgent", "reason"]]
    st.download_button("Download this ranking (CSV)", export.to_csv(index=False).encode("utf-8"),
                       "ranking.csv", "text/csv", icon=":material/download:")


# ================================================================ 4. Merchant page, draft and approval
def merchant_page():
    journey(4)
    state, _ = file_state()
    if state != "ready":
        st.title("Merchant")
        if state == "needs_tags":
            st.warning("Tag the unseen notes first, then the ranking and merchant pages open.",
                       icon=":material/pending_actions:")
            st.page_link(P_TAGS, label="Tag notes", icon=":material/arrow_forward:")
            return
        return need_file()
    res = scored()
    t = res["table"]
    ids = t.merchant_id.tolist()
    if ss.selected not in ids:
        ss.selected = ids[0]
    name_of = dict(zip(t.merchant_id, t.business_name))
    rank_of = dict(zip(t.merchant_id, t["rank"]))
    ss.selected = st.selectbox("Merchant", ids, index=ids.index(ss.selected),
                               format_func=lambda m: f"#{rank_of[m]} · {m} · {name_of[m]}")
    m = t[t.merchant_id == ss.selected].iloc[0].to_dict()

    st.title(f"{m['business_name']}")
    chips = pill(m["track"], look.TRACK[m["track"]]) + pill(f"Risk: {m['risk']}", look.RISK[m["risk"]])
    if m["main"]:
        chips += pill(f"Driver: {m['main']}" + (f" (also {m['also']})" if m["also"] else ""), ("#F0F4F6", "#192839"))
    if m["urgent"]:
        chips += pill(f"Urgent: renews in {m['renewal']} days", ("#FDECEA", "#B3261E"))
    st.html(f'<div>{pill("Rank " + str(m["rank"]), ("#EAF0FF", "#2341C8"))}{chips}</div>'
            f'<div style="color:#5D6D86;font-size:.9rem;margin-top:.2rem">{esc(m["merchant_id"])} · {esc(m["segment"])} · '
            f'{esc(m["industry"])} · {m["tenure"]} months with Razorpay</div>')

    left, right = st.columns([3, 2], gap="large")
    with left:
        st.subheader("Why")
        st.html(f'<div class="reason">{esc(m["reason"])}</div>')
        st.write("")
        st.markdown("**Evidence: how many of 5 sources see a problem**")
        e = m["evidence"]
        csat = f"CSAT {core.num(m['csat'])}/5" if core.has(m["csat"]) else "no CSAT (ignored)"
        detail = {
            "Volume": f"Latest month {m['latest'] * 100:+.1f}% vs 6-month average (alarm at {core.S['AlarmLine'] * 100:.1f}%)",
            "Payments": f"Success rate {m['sr1']:.1f}% → {m['sr2']:.1f}% ({m['srchg']:+.1f} pts; signal at −{core.S['SRDrop']})",
            "Support": f"{m['unresolved']} unresolved of {m['tickets']} tickets (90 days), {csat}",
            "Engagement": f"{m['logins']} logins in 30 days (signal at {core.S['LowLogins']} or fewer)",
            "Account manager note": (f"“{esc(m['note'])}” → {m['tag']}" if m["note"] else "No note"),
        }
        st.html('<div class="ev">' + "".join(
            f'<div class="{"on" if e[k] else ""}"><b><span class="dot"></span>{esc(k)}</b>{detail[k] if k == "Account manager note" else esc(detail[k])}</div>'
            for k in detail) + "</div>")
        st.write("")
        rules = {"Performance": "success rate down 5+ points, or a performance note",
                 "Service": "2+ unresolved tickets, or a service note",
                 "Pricing": "pricing note plus renewal within 60 days",
                 "Disengagement": "TPV alarm plus 5 or fewer logins or a disengagement note"}
        fired = [f"**{d}** ({rules[d]})" for d, on in m["flags"].items() if on]
        st.markdown("**Driver flags:** " + ("; ".join(fired) if fired else "none fired"))
        if not fired and m["sources"]:
            st.caption("A source sees something, but no driver rule fired. For example, low logins count as "
                       "Disengagement only together with the TPV alarm (rule 3).")
        tpv = pd.DataFrame({"Month": [f"m{k}" for k in range(1, 7)], "TPV": m["tpv"]})
        st.markdown("**TPV, last 6 months (Rs lakh)**")
        bars = alt.Chart(tpv).mark_bar(color="#305EFF", cornerRadiusTopLeft=3, cornerRadiusTopRight=3, size=34).encode(
            x=alt.X("Month:N", axis=alt.Axis(labelAngle=0, title=None)),
            y=alt.Y("TPV:Q", axis=alt.Axis(title=None, grid=True, tickCount=4)),
            tooltip=["Month", alt.Tooltip("TPV:Q", format=".1f")])
        avg_line = alt.Chart(pd.DataFrame({"avg": [m["avg"]]})).mark_rule(color="#192839", strokeDash=[4, 3]).encode(y="avg:Q")
        st.altair_chart((bars + avg_line).properties(height=170), width="stretch")
        st.caption("Dashed line: 6-month average.")

    with right:
        st.subheader("Priority")
        if m["track"] in ("CSM outreach", "Light-touch"):
            st.html(
                f'<div class="formula">Value at stake <b>Rs {m["value"]:,.0f}</b> (12 × 1% × Rs {m["tpv"][5]:.1f} lakh)<br>'
                f'× savability ({esc(m["main"])}) <b>{m["weight"]:.2f}</b><br>'
                f'× renewal urgency ({m["renewal"]} days) <b>{m["urgency"]:.2f}</b><br>'
                f'× likelihood ({m["sources"]} of 5 sources) <b>{m["likelihood"]:.2f}</b><br>'
                f'= priority score <b>{m["score"]:,.0f}</b></div>')
            if m["urgent"]:
                st.caption("Renews within 30 days, before next month's re-check, so it skips the queue.")
        else:
            st.caption(f"Not flagged: {m['kind'].lower()}. No priority score.")
        st.subheader("Playbook")
        pb = playbooks().set_index("playbook_id").loc[m["playbook"]]
        st.markdown(f"**{m['playbook']} · {pb['name']}**")
        st.caption(f"{pb['what_happens']}  \nWhen it applies: {pb['when_it_applies']}")

    st.divider()
    draft_section(m, res["same"])


def run_checks(m: dict, items: list[dict]) -> list[dict]:
    """Live fact check (scripts/factcheck.py), cached per exact text so each version is checked once."""
    up, res = ss.upload, ss.scored
    keys = [hashlib.sha256(repr((up["checked"].digest, m["merchant_id"], it["kind"], it["track"], it["subject"],
                                 it["body"])).encode("utf-8")).hexdigest() for it in items]
    todo = [(k, it) for k, it in zip(keys, items) if k not in ss.checks]
    if todo:
        with st.spinner("Fact-checking against this merchant's row…"):
            for (k, _), v in zip(todo, core.fact_check(up["data"], res["out"], [it for _, it in todo])):
                ss.checks[k] = v
    return [ss.checks[k] for k in keys]


def write_ai_draft(m: dict, pb: dict, digest: str):
    try:
        ai_budget(1)
        with st.spinner("DeepSeek is drafting from this merchant's row and playbook…"):
            got = ai.draft(m, pb)
        prev = ss.ai_drafts.get((digest, m["merchant_id"]), {})
        ss.ai_drafts[(digest, m["merchant_id"])] = got | dict(n=prev.get("n", 0) + 1,
                                                              at=datetime.now().strftime("%H:%M"))
        ss.approvals.pop((digest, m["merchant_id"]), None)
        return True
    except ai.AIError as e:
        st.error(str(e), icon=":material/cloud_off:")
        return False


def draft_section(m: dict, same: bool):
    st.subheader("Draft and approval")
    base = core.draft_for(m, same)
    if base is None:
        if m["playbook"] == "PB-5":
            st.info("Expansion check-in (PB-5). Uses no CSM slot and needs no risk outreach this month.",
                    icon=":material/trending_up:")
        else:
            st.info("Monitor only (PB-6). No outreach; re-check next month.", icon=":material/visibility:")
        return
    digest, mid = ss.upload["checked"].digest, m["merchant_id"]
    pb = playbooks().set_index("playbook_id").loc[m["playbook"]].to_dict()

    # 6. Where the draft comes from. The CSM list can use DeepSeek; light-touch keeps the rules template,
    #    as in the workbook, so its messages can't state anything the data doesn't hold.
    sources = (["Stored draft"] if base["kind"] == "stored" else []) + (
        ["DeepSeek draft"] if m["track"] == "CSM outreach" else []) + ["Template"]
    k_src = f"src_{digest}_{mid}"
    src = (st.segmented_control("Draft from", sources, key=k_src, **first_default(k_src, sources[0]),
                                on_change=keep_one(k_src, sources[0])) or sources[0]
           if len(sources) > 1 else "Template")
    if m["track"] == "Light-touch":
        st.caption("Light-touch merchants get the rules template only, filled from the row, as in the workbook.")

    if src == "DeepSeek draft":
        with st.expander("What DeepSeek is told (system prompt) and what it sees (this merchant only)",
                         icon=":material/visibility:"):
            st.code(ai.system_prompt(m["main"]), language=None, wrap_lines=True)
            st.code(ai.brief(m, pb), language=None, wrap_lines=True)
        aid = ss.ai_drafts.get((digest, mid))
        if not aid:
            if not ai.available("DeepSeek"):
                st.info("No DeepSeek key is set, so AI drafting is off. Use the template.", icon=":material/cloud_off:")
            elif st.button("Write a draft with DeepSeek", type="primary", icon=":material/auto_awesome:"):
                if write_ai_draft(m, pb, digest):
                    st.rerun()
            st.caption(f"One call writes an email and a WhatsApp message. AI calls this session: "
                       f"{ss.ai_calls} of {ai.SESSION_CAP}.")
            return
        d = dict(kind="ai", subject=aid["subject"], body=aid["email"], whatsapp=aid["whatsapp"],
                 by=f"DeepSeek AI draft · {aid['model']} · version {aid['n']}",
                 check="Read it against the evidence above. The AI saw only this merchant's row and playbook.")
        ver = f"ai{aid['n']}"
    elif src == "Template":
        d = core.draft_for(m, same, prefer_template=True) | {}
        d["whatsapp"] = d["body"]
        ver = "t"
    else:
        d, ver = dict(base, whatsapp=None), "s"

    st.html(pill(d["by"], ("#EAF0FF", "#2341C8") if d["kind"] != "template" else ("#F0F4F6", "#192839")))
    st.markdown(f"**Before sending, the CSM should:** {d['check']}")
    k_sub, k_body, k_wa = (f"{p}_{digest}_{mid}_{ver}" for p in ("sub", "body", "watext"))
    ss.setdefault(k_sub, d["subject"])
    ss.setdefault(k_body, d["body"])

    st.text_input("Email subject", key=k_sub)
    st.text_area("Email", key=k_body, height=190)
    wa_on = st.toggle("Merchant has opted in to WhatsApp", key=f"wa_{digest}_{mid}",
                      help="The data has no opt-in column, so this stays off unless the CSM knows the merchant opted in.")
    if wa_on and d.get("whatsapp"):
        ss.setdefault(k_wa, d["whatsapp"])
        st.text_area("WhatsApp message", key=k_wa, height=110)
    elif wa_on:
        st.caption("The stored draft has no WhatsApp version. Choose 'DeepSeek draft' or 'Template' for one.")
    use_wa = wa_on and bool(d.get("whatsapp"))

    # 7. Live fact check on exactly what would be approved
    items = [dict(id=mid, subject=ss[k_sub], body=ss[k_body], kind=d["kind"], track=m["track"])]
    if use_wa:
        items.append(dict(id=mid, subject="", body=ss[k_wa], kind=d["kind"], track=m["track"]))
    verdicts = run_checks(m, items)
    blocked = False
    for label, v in zip(["Email", "WhatsApp"], verdicts):
        if v["ok"]:
            st.html(pill(f"{label} · fact check (live): Passed", ("#E6F5EE", "#00704A")) +
                    f'<span style="color:#5D6D86;font-size:.85rem">Numbers used: {esc(v["numbers"])}</span>')
        else:
            blocked = True
            st.error(f"**{label} blocked by the fact check.** It can't be approved until this is fixed:\n\n" +
                     "\n".join(f"- {p}" for p in v["problems"]), icon=":material/block:")
    if src == "DeepSeek draft":  # specific, not generic: the driver's own numbers must be in the message
        for label, text in [("Email", ss[k_body])] + ([("WhatsApp", ss[k_wa])] if use_wa else []):
            miss = ai.missing_facts(m, text)
            if miss:
                st.warning(f"**{label} is vague:** it doesn't state " + "; ".join(miss) +
                           ". Add it, or regenerate.", icon=":material/info:")
            else:
                st.html(pill(f"{label} · states the key facts", ("#EAF0FF", "#2341C8")) +
                        f'<span style="color:#5D6D86;font-size:.85rem">{esc("; ".join(ai.must_include(m)))}</span>')
    if src == "DeepSeek draft":
        b1, b2, _ = st.columns([1.2, 1.5, 3])
        if b1.button("Regenerate", icon=":material/refresh:", help=f"AI calls this session: {ss.ai_calls} of "
                                                                     f"{ai.SESSION_CAP}"):
            if write_ai_draft(m, pb, digest):
                st.rerun()
        if blocked:
            b2.button("Use the template instead", icon=":material/description:",
                      on_click=lambda: ss.update({k_src: "Template", f"{k_src}__last": "Template"}))

    channel = "email + WhatsApp" if use_wa else "email"
    if m["track"] == "Light-touch":
        st.caption(f"Channel: {channel}. Escalate to next month's CSM list if the merchant replies, its signals "
                   f"worsen, or renewal is within {ESCALATE_DAYS} days.")
    else:
        st.caption(f"Channel: {channel}. Pricing changes stay with the account manager.")

    # 8. The CSM approves. Nothing is sent.
    current = dict(subject=ss[k_sub], body=ss[k_body], whatsapp=ss[k_wa] if use_wa else None, source=src)
    edited = (current["subject"] != d["subject"] or current["body"] != d["body"]
              or (use_wa and current["whatsapp"] != d["whatsapp"]))
    appr = ss.approvals.get((digest, mid))
    same_as_appr = appr and all(appr.get(k) == current[k] for k in current)

    def approve():
        ss.approvals[(digest, mid)] = dict(current, at=datetime.now().strftime("%d %b %Y, %H:%M"), edited=edited)

    def revert():
        ss[k_sub], ss[k_body] = d["subject"], d["body"]
        if use_wa:
            ss[k_wa] = d["whatsapp"]
        ss.approvals.pop((digest, mid), None)

    c1, c2, _ = st.columns([1.3, 1.2, 3])
    c1.button("Approve (doesn't send)", type="primary", icon=":material/task_alt:", on_click=approve,
              disabled=blocked, help="Blocked drafts can't be approved." if blocked else None)
    c2.button("Revert to draft", icon=":material/undo:", on_click=revert, disabled=not edited and not appr)

    if same_as_appr:
        st.success(f"Approved by the CSM on {appr['at']}{' (edited)' if appr['edited'] else ''}, {channel}. "
                   "**Not sent:** this prototype never sends anything. Copy it into your usual tool.",
                   icon=":material/task_alt:")
        text = f"Subject: {appr['subject']}\n\n{appr['body']}\n" + (
            f"\n--- WhatsApp ---\n{appr['whatsapp']}\n" if appr.get("whatsapp") else "")
        st.download_button("Download approved message (.txt)", text.encode("utf-8"), f"{mid}_approved.txt",
                           "text/plain", icon=":material/download:")
    elif appr:
        st.warning("Changed since it was approved. Approve again before using it.", icon=":material/edit:")
    elif edited and not blocked:
        st.warning("Edited and not yet approved. The fact check above ran on your edited text.", icon=":material/edit:")


# ================================================================ navigation
P_HOME = st.Page(overview, title="Overview", icon=":material/home:", default=True)
P_UPLOAD = st.Page(upload_page, title="1 · Upload", icon=":material/upload_file:", url_path="upload")
P_TAGS = st.Page(tags_page, title="2 · Tag notes", icon=":material/sell:", url_path="tags")
P_RANK = st.Page(rank_page, title="3 · Ranking", icon=":material/format_list_numbered:", url_path="ranking")
P_MERCHANT = st.Page(merchant_page, title="4 · Merchant", icon=":material/storefront:", url_path="merchant")

look.topbar()
nav = st.navigation([P_HOME, P_UPLOAD, P_TAGS, P_RANK, P_MERCHANT], position="top")
nav.run()
look.footer()
