"""Compare three ways to tag account manager notes the lookup table hasn't seen.

Taggers: keyword rules (no AI), Jev (TypeSafe), DeepSeek. The answer key is the user's reviewed tags.
Usage (from D:\\Master CV\\Razorpay):  python -I -B scripts/test_tagging.py
Writes scripts/tagging_results.json. Reads API keys from .env and never prints them.
"""
import json
import os
import re
import sys
import time
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
import urllib.error
import urllib.request

sys.dont_write_bytecode = True
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from apikeys import key  # noqa: E402

TAGS = ["Performance", "Service", "Pricing", "Disengagement", "Seasonal", "Positive", "Neutral"]
RISK = {"Performance", "Service", "Pricing", "Disengagement"}
DEFINITIONS = {
    "Performance": "Payment success problems: failed or declined payments, UPI or bank drop-offs, a request to review success rates. Not about support speed.",
    "Service": "Support problems still open: slow support replies, unresolved tickets, escalations. Not a complaint that has already been fixed or resolved.",
    "Pricing": "Price or switching risk: comparing quotes or providers, a competitor offering lower fees or MDR, asking for the rate card or fees before renewal, exploring other options.",
    "Disengagement": "The merchant is drifting away: no response to outreach, unclear contact, owner or manager change, the dashboard rarely used.",
    "Seasonal": "A volume dip explained by the season or a festive or peak cycle, expected to recover.",
    "Positive": "Healthy or growing: growth, interest in more products, a check-in with no concerns, many questions that were all resolved while growing.",
    "Neutral": "Routine or already resolved matters, a small low-touch merchant, never raised tickets; no risk and no growth signal.",
}
INSTRUCTIONS = ("This is a customer success manager's note about a merchant on a payments platform. Choose the tag "
                "that best describes the note's main signal. If the note shows any churn risk (Performance, Service, "
                "Pricing or Disengagement), choose that risk over Seasonal, Positive or Neutral.")


# ---------------------------------------------------------------- 1) keyword rules (written from the 22 original notes)
KEYWORDS = [
    ("Pricing", r"competitor|quote|rate card|\bmdr\b|pricing|cheaper|exploring options|other providers?|\bfees?\b"),
    ("Performance", r"success rate|\bupi\b|drop-?offs?|failed payments?|failures?|declin"),
    ("Disengagement", r"no response|not respond|unclear contact|new manager|handed|rarely used|unreachable"),
    ("Seasonal", r"season|festive|peak|last year|slow quarter"),
    ("Positive", r"growing|interested|adding|no concerns|expand"),
    ("Service", r"support|ticket|escalat|not closed|unresolved"),
]


def tag_keywords(text):
    low = text.lower()
    for tag, pattern in KEYWORDS:
        if re.search(pattern, low):
            if tag == "Service" and (re.search(r"resolved|fixed|closed now|all closed|never", low)
                                     and "not closed" not in low):
                continue
            return {"tag": tag}
    return {"tag": "Neutral"}


# ---------------------------------------------------------------- 2) Jev
def post(url, body, auth, retries=3):
    data = json.dumps(body).encode()
    for attempt in range(retries):
        req = urllib.request.Request(url, data=data, headers={"Authorization": "Bearer " + auth,
                                                              "Content-Type": "application/json"})
        try:
            return json.load(urllib.request.urlopen(req, timeout=60))
        except urllib.error.HTTPError as e:
            if e.code in (429, 500, 502, 503) and attempt < retries - 1:
                time.sleep(2 * (attempt + 1))
                continue
            raise RuntimeError(f"HTTP {e.code}: {e.read()[:300]!r}")
        except (urllib.error.URLError, TimeoutError):
            if attempt < retries - 1:
                time.sleep(2 * (attempt + 1))
                continue
            raise


JEV_KEY = None


def tag_jev(text):
    body = {"model": "jev-latest", "state": text,
            "questions": {"tag": {"type": "choice", "instructions": INSTRUCTIONS, "criteria": DEFINITIONS}}}
    r = post("https://api.typesafe.ai/v1/systemone", body, JEV_KEY)
    a = r["answers"]["tag"]
    return {"tag": a["choice"], "confidence": a.get("confidence"), "probabilities": a.get("probabilities"),
            "model": r.get("model")}


# ---------------------------------------------------------------- 3) DeepSeek
DS_KEY, DS_MODEL = None, "deepseek-flash"


def tag_deepseek(text):
    system = (INSTRUCTIONS + " Tags:\n" + "\n".join(f"- {t}: {d}" for t, d in DEFINITIONS.items()) +
              '\nReply with JSON only: {"tag": "<one of the tags>"}')
    body = {"model": DS_MODEL, "temperature": 0, "response_format": {"type": "json_object"},
            "messages": [{"role": "system", "content": system}, {"role": "user", "content": text}]}
    r = post("https://api.deepseek.com/chat/completions", body, DS_KEY)
    content = r["choices"][0]["message"]["content"]
    tag = json.loads(content).get("tag", "")
    return {"tag": tag if tag in TAGS else f"INVALID:{tag}", "model": r.get("model")}


def run_all(fn, notes, workers=4):
    def one(n):
        try:
            return fn(n["text"])
        except Exception as e:  # keep going; an error counts as a miss
            return {"tag": "ERROR", "error": str(e)[:200]}
    with ThreadPoolExecutor(workers) as ex:
        return list(ex.map(one, notes))


def score(name, notes, results):
    ok = [r["tag"] == n["label"] for n, r in zip(notes, results)]
    by_group = {g: (sum(o for o, n in zip(ok, notes) if n["group"] == g), sum(n["group"] == g for n in notes))
                for g in ("original", "reworded", "edge")}
    pricing = [(r["tag"] == "Pricing") for n, r in zip(notes, results) if n["label"] == "Pricing"]
    risk_missed = sum(1 for n, r in zip(notes, results) if n["label"] in RISK and r["tag"] not in RISK)
    false_risk = sum(1 for n, r in zip(notes, results) if n["label"] not in RISK and r["tag"] in RISK)
    errors = sum(r["tag"] in ("ERROR",) or r["tag"].startswith("INVALID") for r in results)
    print(f"\n{name}: {sum(ok)}/{len(ok)} correct  | original {by_group['original'][0]}/{by_group['original'][1]}"
          f"  reworded {by_group['reworded'][0]}/{by_group['reworded'][1]}  edge {by_group['edge'][0]}/{by_group['edge'][1]}")
    print(f"   Pricing notes caught: {sum(pricing)}/{len(pricing)} | risk notes tagged no-risk: {risk_missed}"
          f" | no-risk notes tagged as risk: {false_risk} | errors: {errors}")
    for n, r, o in zip(notes, results, ok):
        if not o:
            conf = f" (confidence {r['confidence']:.2f})" if isinstance(r.get("confidence"), (int, float)) else ""
            print(f"   miss [{n['group']}] {n['text']!r}: expected {n['label']}, got {r['tag']}{conf}"
                  + (f" | {r['error']}" if r.get("error") else ""))
    return {"correct": sum(ok), "total": len(ok), "by_group": by_group, "pricing_caught": [sum(pricing), len(pricing)],
            "risk_missed": risk_missed, "false_risk": false_risk, "errors": errors}


if __name__ == "__main__":
    notes = json.load(open(os.path.join(HERE, "note_test_set.json"), encoding="utf-8"))["notes"]
    originals = [n for n in notes if n["group"] == "original"]
    JEV_KEY, DS_KEY = key("TYPESAFE_API_KEY"), key("DEEPSEEK_API_KEY")
    out = {"run_at": time.strftime("%Y-%m-%d %H:%M"), "notes": notes, "taggers": {}}
    for name, fn in (("Keyword rules", tag_keywords), ("Jev", tag_jev), ("DeepSeek", tag_deepseek)):
        res = run_all(fn, notes) if fn is not tag_keywords else [fn(n["text"]) for n in notes]
        summary = score(name, notes, res)
        if fn is not tag_keywords:
            again = run_all(fn, originals)
            first = [r for n, r in zip(notes, res) if n["group"] == "original"]
            same = sum(a["tag"] == b["tag"] for a, b in zip(first, again))
            summary["repeat_same_on_originals"] = [same, len(originals)]
            models = Counter(r.get("model") for r in res if r.get("model"))
            summary["models"] = dict(models)
            print(f"   repeat run on the 22 originals: {same}/{len(originals)} same tag | model versions: {dict(models)}")
        out["taggers"][name] = {"summary": summary, "results": res}
    json.dump(out, open(os.path.join(HERE, "tagging_results.json"), "w", encoding="utf-8"), indent=1, ensure_ascii=False)
    print("\nSaved scripts/tagging_results.json")
