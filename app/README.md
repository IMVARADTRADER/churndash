# ChurnDash: merchant churn early-warning app (Tiers 1 and 2)

A Streamlit prototype beside the workbook, for Razorpay Build-a-thon Problem 2. It's a student build-a-thon
prototype, not a Razorpay product, and Claude (AI) wrote the code (say so in written answer Q4).
The workbook `P2_Churn_Scoring.xlsx` stays the source of truth. This app runs the same rules and must never disagree with it.
AI only suggests tags and writes text. A person confirms every tag and approves every message, and nothing is ever sent.

## Run it

From the project folder (`D:\Master CV\Razorpay`):

```bash
python -m pip install -r app/requirements.txt
```
```bash
python -m streamlit run app/app.py
```

It opens at http://localhost:8501. The theme lives in `app/.streamlit/config.toml`. The AI keys are read from
`.env` through `scripts/apikeys.py`. Without a key, the app still works: you pick the tags and use templates.

## Check it

```bash
python -B app/tests/run_tests.py
```

These checks cover tests 1 to 4 from handoff section 11 without calling any AI (40 checks). The run also checks
that every reason, risk label, playbook and template matches the workbook's saved values. To call DeepSeek and Jev
as well, which costs a little:

```bash
python -B app/tests/run_tests.py --live
```

Add `--drafts-only` to skip the 27 tag calls. To get files you can upload in the demo (reworded notes, a planted
"discount" note, and nine broken files):

```bash
python -B app/tests/make_demo_files.py
```

They go to `app/demo_files/`, which `.gitignore` keeps out of any repository because they hold Data Pack rows.

## The process, step by step

| Step | Page | What happens | Where in the code |
| --- | --- | --- | --- |
| 1 Upload | 1 · Upload | Reads the `P2_Merchants` sheet in memory and checks the sheet name, the 20 required columns, the types, ranges, blanks and duplicate IDs. A wrong file gets a plain-language error naming the Excel row. Shows the "synthetic data only" warning. | `core.check_file` |
| 2 Tag notes | 2 · Tag notes | **Why:** the rules can't read free text, so each account manager note becomes one of 7 tags, and the rules use the tag. A Pricing tag is the only way a Pricing flag fires; Performance and Service tags can fire a flag on their own; a Disengagement tag needs the TPV alarm; any risk tag counts as the 5th evidence source. The page shows this in a table, with how many merchants carry each tag, and every upload passes through this step. Notes in my reviewed table (the 22 in `TAGS`, `scripts/reference.py`) keep my tag. For an **unseen** note, DeepSeek (the default) or Jev (the switch) suggests a tag using the prompt from the 94-note test, and the suggestion is pre-selected. A person confirms every tag before scoring. A Jev confidence below 0.6 shows "Please check". If the AI is down, the person picks from the dropdown. | `ai.suggest_tags`, `core.pending_notes` |
| 3 Score and rank | 3 · Ranking | Calls `reference.run()` from `scripts/reference.py`, **unchanged**. Confirmed tags are added to its lookup table for that run only, and my 22 reviewed tags always win. The caption counts how many AI suggestions were kept or changed. | `core.score` |
| 3b Badge | 3 · Ranking | When the upload holds exactly the workbook's data, it compares every rank with the ranks Excel saved: "Matches the submitted workbook: 100 of 100 ranks". | `core.same_data`, `core.compare_ranks` |
| 4 Ranked list | 3 · Ranking | All merchants, with the CSM 15 tinted blue and the light-touch 9 amber. Click any cell, or a row's tick box, to open that merchant. | `rank_page` |
| 5 Merchant page | 4 · Merchant | The reason (the workbook's wording), the 5 evidence sources with their numbers, which driver flags fired, the priority maths and the playbook. | `core.explain`, `merchant_page` |
| 6 Draft | 4 · Merchant | Choose the **stored draft** (the 15, original file only), a **DeepSeek draft** (an email and a WhatsApp message written from that merchant's row and playbook only; the system prompt and the brief are shown on the page), or the **rules template**. Light-touch merchants get the template only, as in the workbook. A DeepSeek draft must state its driver's numbers ("Must include" in the brief): days to renewal for Pricing, both success rates for Performance, unresolved tickets for Service, logins for Disengagement, plus renewal days when renewal is within 60. A draft without them gets a "vague" warning. | `ai.draft`, `ai.must_include`, `ai.missing_facts` |
| 7 Fact check | 4 · Merchant | `scripts/factcheck.py` runs **unchanged** on exactly the text the CSM would approve, re-running after every edit. A failed draft is blocked, shows its reasons, and offers Regenerate or Use the template instead. | `core.fact_check` |
| 8 Approve | 4 · Merchant | The CSM edits and approves. WhatsApp only appears when the opt-in toggle is on, because the data has no opt-in column. An approval is recorded for the session only. **Nothing is ever sent.** | `draft_section` |

## Choices to be able to defend

- **Thinking is off for drafts.** With thinking on, `deepseek-flash` spent 1,700–2,300 tokens and 11–14 s reasoning per draft (an 800-token cap left the reply empty). With it off, a draft took about 2 s and 214 tokens. The rules have already done the judging, and the fact check catches what the writing gets wrong. Tag suggestions keep the exact settings from the 94-note test.
- **One rule added on top of `factcheck.py`.** That script accepts 3 and 48 for every merchant. A live draft for C2090 (Disengagement, PB-4) promised PB-1's "fix plan within 3 days" and still passed. The app now blocks a playbook's time commitment in another playbook's message, and the prompt offers each merchant only its own commitment. It's the same lesson as C2018 and C2062: the automated check can't judge meaning.
- **Two `factcheck.py` rules are relaxed, and only where they can't apply.** Templates skip the driver key-phrase rule, because their wording is chosen by driver and the workbook's Performance template says "payments team", not "Payments Ops". Light-touch merchants skip the "must be on the CSM list" rule.
- **The fact check reads files.** The merchant file and the draft go into a temporary folder that is deleted within a second or two.
- **Drafts must be specific.** The first prompt said to list only the signals the rules saw, and Pricing drafts came out with no numbers at all. Now every draft must quote its driver's numbers. The volume drop ("12% below average") is still described in words only, because `factcheck.py` has no TPV fields and would block the number.
- **Spending caps:** 60 AI calls per browser session and 300 per server per day (`ai.SESSION_CAP`, `ai.DAILY_CAP`). Also set a spending limit in the DeepSeek console.

## Test 3, honestly

- The fact check blocks a draft that says "discount", "lower rate", a number not in the row, another merchant's name, a missing driver phrase, or another playbook's promise (offline tests).
- In the app, typing "we can offer them a discount" into an AI draft blocks it and disables Approve.
- Live: before the drafts had to quote specifics, DeepSeek never wrote "discount" in 3 tries. After that change, its first try echoed the planted note's "discount", and the fact check blocked it. The next 2 tries were clean. The prompt is the first line of defence and the fact check the second.

## Why the app can't disagree with the workbook

- Flags, the main driver, evidence count, score, rank and track all come from `reference.run()`. The app doesn't re-implement them.
- The merchant page re-derives the four flags only to *show* them, and stops with an error if they ever differ from `reference.py`'s result.
- The reason, risk label, playbook and template text copy the workbook's formulas, and `run_tests.py` checks them against the workbook's saved values for all 100 merchants.
- Thresholds come from `reference.S`, so there's one set of settings in Python. If you change Settings in the workbook, change `reference.py` to match (handoff section 9).

## Known limits

- The fact check covers numbers, names, promise words and driver fit. It can't judge meaning: read every draft.
- Word limits (110 for an email, 60 for WhatsApp) are asked for but not enforced. One live WhatsApp draft ran to 67 words.
- Confirmed tags and AI drafts last for the browser session. They aren't saved to disk, which keeps the "store nothing" promise.
- On a file that differs from the original, CSM-list merchants have no stored draft: use DeepSeek or the template.
- The badge only compares when the data is identical to the workbook's, row order included, because `reference.py` breaks ties by row.
- Hosting: the download page and the badge read `P2_Churn_Scoring.xlsx` and the PDF from the project folder. Put the keys in the host's secrets (`app/.streamlit/secrets.toml`, which `.gitignore` excludes). Don't put the Data Pack in a public repository.

## Deploy (ChurnDash)

- **App:** Streamlit Community Cloud, from the private GitHub repository `churndash`, main file `app/app.py`, subdomain `churndash` (so the address is `churndash.streamlit.app`). Python 3.12 or 3.13. Paste `DEEPSEEK_API_KEY` and `TYPESAFE_API_KEY` into the app's Secrets. Sharing: anyone with the link.
- **Landing page:** Vercel, importing the same repository with the root directory left at the repository root. `vercel.json` copies the workbook and the written-answers PDF into `site/downloads/` and serves `site/`. If the Streamlit subdomain differs, change the `app-link` address in `site/index.html`.
- **Repository:** `.gitignore` is an allow-list, so only the app, `scripts/`, `site/`, the workbook, the PDF and `Problem Statement 2.xlsx` (for the demo button) go in. `.env`, `backups/`, the full Data Pack and the ebook never do. Keep the repository private.
