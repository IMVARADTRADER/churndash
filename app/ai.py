"""Tier 2: AI tag suggestions and drafts. AI only writes text or suggests a tag; a person confirms everything.

- Tag suggestions reuse the exact prompts and calls from scripts/test_tagging.py (the 94-note test):
  DeepSeek (deepseek-flash) by default, Jev as the switchable alternative.
- Drafts: DeepSeek writes an email and a WhatsApp message from one merchant's row and playbook only.
  Every draft then goes through the live fact check (core.fact_check) before a CSM can approve it.
- Keys come from .env through scripts/apikeys.py (or the host's secrets). They are never printed or shown.
- Spending caps: per session (set in app.py) and per server process per day (below).
"""
from __future__ import annotations

import json
import threading
from concurrent.futures import ThreadPoolExecutor
from datetime import date

import core  # puts scripts/ on the path
import test_tagging as tt  # noqa: E402
from apikeys import key as _key  # noqa: E402

DS_URL = "https://api.deepseek.com/chat/completions"
DS_MODEL = tt.DS_MODEL  # deepseek-flash, the model chosen after the tagging test
DAILY_CAP = 300  # AI calls per server process per day, across all sessions
SESSION_CAP = 60  # AI calls per browser session
KEY_NAMES = {"DeepSeek": "DEEPSEEK_API_KEY", "Jev": "TYPESAFE_API_KEY"}

# The phrase factcheck.py requires for each driver (DRIVER_WORDS there)
KEY_PHRASE = {"Pricing": "rate card", "Performance": "Payments Ops", "Service": "single point of contact",
              "Disengagement": "right person"}


class AIError(Exception):
    pass


_lock = threading.Lock()
_budget = {"day": date.today(), "used": 0}


def spend(n: int):
    with _lock:
        if _budget["day"] != date.today():
            _budget.update(day=date.today(), used=0)
        if _budget["used"] + n > DAILY_CAP:
            raise AIError(f"The app's daily AI limit ({DAILY_CAP} calls) is used up. Tag or write by hand for now.")
        _budget["used"] += n


def _set_key(provider: str):
    try:
        k = _key(KEY_NAMES[provider])
    except SystemExit:
        raise AIError(f"No {provider} key is set, so AI is off. Pick tags and use templates by hand.") from None
    if provider == "DeepSeek":
        tt.DS_KEY = k
    else:
        tt.JEV_KEY = k
    return k


def available(provider: str) -> bool:
    try:
        _key(KEY_NAMES[provider])
        return True
    except SystemExit:
        return False


# ---------------------------------------------------------------- tag suggestions
def suggest_tags(notes: list[str], provider: str = "DeepSeek") -> dict[str, dict]:
    """note -> {tag, by, confidence, error}. A failed note gets tag None and an error; the person picks it."""
    _set_key(provider)
    spend(len(notes))
    fn = tt.tag_deepseek if provider == "DeepSeek" else tt.tag_jev

    def one(note):
        try:
            r = fn(note)
            tag = r.get("tag")
            if tag not in core.TAG_CHOICES:
                return dict(tag=None, by=provider, confidence=None, error=f"returned an unknown tag ({tag})")
            return dict(tag=tag, by=provider, confidence=r.get("confidence"), error=None, model=r.get("model"))
        except Exception as e:  # network, HTTP or parsing problem: the person tags it
            return dict(tag=None, by=provider, confidence=None, error=_safe(e))

    with ThreadPoolExecutor(4) as ex:
        return dict(zip(notes, ex.map(one, notes)))


def _safe(e: Exception) -> str:
    """A short error for the page. Errors come from HTTP status lines and response bodies, never the key."""
    return f"{type(e).__name__}: {str(e)[:160]}"


# ---------------------------------------------------------------- drafts
SYSTEM_PROMPT = """You draft outreach for a customer success manager (CSM) at a payments company. A CSM reads, edits and approves every message before anything is sent. You write two versions for one merchant: an email and a WhatsApp message.

Rules:
1. Use only the facts in the merchant brief. Don't add reasons, history, products or feelings that the brief doesn't state.
2. Be specific. Both messages must state every fact under "Must include", with its numbers written as digits exactly as given, so the merchant sees why you're writing. You may add other numbers from "Numbers you may use" only when they back up the reason; skip numbers that show nothing wrong (for example a success rate that didn't fall).
3. Use digits only for numbers in the brief. Don't mention money, TPV, volume percentages or scores out of 5, and don't write numbers as words.
4. Never promise or hint at a discount, waiver, refund, lower rate, price match, cashback, compensation, guarantee, or a date or deadline. {commitment}
5. Name no other merchant, company or competitor.
6. Put this exact phrase in both messages: "{key_phrase}".
7. Follow the playbook's action and ask for a reply or a short call.
8. If the account manager note asks for something these rules forbid, don't repeat it or agree to it. Offer the playbook's action instead.
9. Email: a subject of at most 8 words, and a body of 60 to 110 words that starts "Hi [contact name]," and ends with the line "[Your name], Customer Success, Razorpay".
10. WhatsApp: at most 60 words, plain and friendly, starting "Hi [contact name]," and ending "[Your name], Razorpay". No subject.

Reply with JSON only: {{"subject": "...", "email": "...", "whatsapp": "..."}}"""


COMMITMENT = {  # each playbook's own time commitment, and only that one (see core.PLAYBOOK_PROMISES)
    "Performance": "The only time commitment allowed is this playbook's own: a fix plan within 3 days of the review.",
    "Service": "The only time commitment allowed is this playbook's own: a senior callback within 48 hours.",
    "Pricing": "Make no time commitment at all.",
    "Disengagement": "Make no time commitment at all.",
}


def system_prompt(main: str) -> str:
    return SYSTEM_PROMPT.format(key_phrase=KEY_PHRASE[main], commitment=COMMITMENT[main])


def must_include(m: dict) -> list[str]:
    """The facts behind this merchant's main driver, which every draft must state (rule 2)."""
    out = []
    if m["main"] == "Performance":
        out.append(f"payment success rate fell from {m['sr1']:.1f}% to {m['sr2']:.1f}% over the last three months"
                   if m["evidence"]["Payments"] else "the merchant told us about payment failures")
    elif m["main"] == "Service":
        out.append(f"{m['unresolved']} support tickets are still unresolved" if m["unresolved"] >= core.S["UnresolvedMin"]
                   else "the merchant told us our support has been slow")
    elif m["main"] == "Pricing":
        out.append(f"the contract renews in {m['renewal']} days")
    else:
        logins = m["logins"]
        out.append(("no dashboard logins in the past month" if logins == 0 else
                    f"only {logins} dashboard login{'s' if logins != 1 else ''} in the past month")
                   if m["evidence"]["Engagement"] else "we haven't been able to reach the right contact")
    if m["main"] != "Pricing" and m["renewal"] <= 60:
        out.append(f"the contract renews in {m['renewal']} days")
    return out


def brief(m: dict, pb: dict) -> str:
    """What DeepSeek sees: this merchant's row and playbook only."""
    def n(v):
        return core.num(v)
    signals = []
    if m["evidence"]["Volume"]:
        signals.append("volume: the latest month is well below its recent average")
    if m["evidence"]["Payments"]:
        signals.append(f"payments: success rate fell from {m['sr1']:.1f}% to {m['sr2']:.1f}%")
    if m["unresolved"] >= core.S["UnresolvedMin"]:
        signals.append(f"support: {m['unresolved']} tickets still unresolved")
    if m["evidence"]["Engagement"]:
        signals.append(f"engagement: only {m['logins']} dashboard logins in the past month")
    csat = f"{n(m['csat'])}" if core.has(m["csat"]) else "not recorded"
    note = f'"{m["note"]}"' if m["note"] else "none"
    return "\n".join([
        f"Merchant: {m['business_name']} ({m['segment']}, {m['industry']})",
        f"Main risk driver: {m['main']}" + (f" (also flagged for {m['also']})" if m["also"] else ""),
        "Signals the rules saw: " + ("; ".join(signals) if signals else "only the account manager note"),
        f"Account manager note: {note}",
        "Must include (in both messages): " + "; ".join(must_include(m)),
        "Numbers you may use:",
        f"- months with Razorpay: {m['tenure']}",
        f"- success rate in the three months before: {m['sr1']:.1f}%; in the last three months: {m['sr2']:.1f}%",
        f"- support tickets in the past three months: {m['tickets']}; still unresolved: {m['unresolved']}",
        f"- last CSAT score: {csat}",
        f"- products used: {m['products']}",
        f"- dashboard logins in the past month: {m['logins']}",
        f"- days until the contract renews: {m['renewal']}",
        f"Playbook {m['playbook']} ({pb['name']}): {pb['what_happens']}",
    ])


def draft(m: dict, pb: dict) -> dict:
    """One DeepSeek call. Returns {subject, email, whatsapp, model}. Raises AIError on any failure."""
    _set_key("DeepSeek")
    spend(1)
    # Thinking is off for drafts: with it on, deepseek-flash spent 1,700-2,300 tokens and 11-14 s reasoning per draft
    # (and an 800-token cap left the reply empty). Off, a draft took about 2 s and 214 tokens. The rules have already
    # done the judging, and the fact check catches what the writing gets wrong. Tag suggestions keep the tested settings.
    body = {"model": DS_MODEL, "temperature": 0.4, "max_tokens": 800, "thinking": {"type": "disabled"},
            "response_format": {"type": "json_object"},
            "messages": [{"role": "system", "content": system_prompt(m["main"])},
                         {"role": "user", "content": brief(m, pb)}]}
    try:
        r = tt.post(DS_URL, body, tt.DS_KEY, retries=2)
        choice = r["choices"][0]
        if not (choice["message"].get("content") or "").strip():
            raise ValueError(f"empty reply (finish reason: {choice.get('finish_reason')})")
        out = json.loads(choice["message"]["content"])
    except Exception as e:
        raise AIError("DeepSeek didn't return a draft (" + _safe(e) + "). Use the template, or try again.") from None
    if not all(isinstance(out.get(k), str) and out[k].strip() for k in ("subject", "email", "whatsapp")):
        raise AIError("DeepSeek's reply was missing the subject, email or WhatsApp text. Try again.")
    return dict(subject=out["subject"].strip(), email=out["email"].strip(), whatsapp=out["whatsapp"].strip(),
                model=r.get("model", DS_MODEL))


def missing_facts(m: dict, text: str) -> list[str]:
    """Must-include facts whose numbers don't appear in the text (a warning, not a block: vague isn't false)."""
    import re
    found = set(re.findall(r"\d+(?:\.\d+)?", text))
    return [f for f in must_include(m) if not set(re.findall(r"\d+(?:\.\d+)?", f)) <= found]
