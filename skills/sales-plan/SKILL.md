# /sales plan — End-to-End Account Business Plan Generator

Trigger on `/sales plan <url>` to run all six `/sales` research sub-skills for a target account and auto-populate the Adobe Lite DX Business Plan PowerPoint template. Use when preparing an account plan, QBR deck, or new account review. Outputs land in the account's deal folder under the configured deals root.

## When to Use

- User types `/sales plan <url>` — always invoke this skill
- User asks to "build an account plan", "prep a business plan deck", or "create a QBR deck" for a named account
- User wants to run all sales research and get a populated PPT in one command
- Refresh: user re-runs `/sales plan` on an existing deal folder to update research and regenerate the PPT

## When NOT to Use

- Running a single sub-skill in isolation (`/sales prospect`, `/sales qualify`, `/sales contacts`, etc.)
- Pipeline reporting or forecasting (`/sales report`, `/sales report-pdf`)
- Outreach-only tasks where no PPT or full research is needed
- When no company URL is provided — prospect URL is required to proceed

## Limitations

- Requires all 6 sub-skills installed: `sales-prospect`, `sales-qualify`, `sales-contacts`, `sales-competitors`, `sales-prep`, `sales-outreach`
- Requires Python 3.8+ with `python-pptx` installed
- Requires the Lite DX Business Plan Template `.pptx` synced from OneDrive
- Research sub-skills require internet access; PPT generation does not
- Slides 4 and 8 cannot be auto-populated without Clari/Panorama data — flag these explicitly

## Invocation syntax

```
/sales plan <url> [--arr=<amount>] [--renewal=<quarter>] [--stage=<stage>] [--close-date=<date>] [--products=<list>]
```

**Examples:**
```
/sales plan https://www.appliedmaterials.com
/sales plan https://www.snowflake.com --arr=150000 --renewal=Q3 --stage=POC --close-date=2026-08-31 --products="ALM,Marketo"
```

---

## Step 1 — Parse arguments

Extract from the invocation string:
- `url` (required)
- `arr` (optional) — current or target ARR in dollars
- `renewal` (optional) — renewal quarter, e.g. `Q3` or `Q3 FY26`
- `stage` (optional) — deal stage, e.g. `POC`, `Eval`, `Negotiate`
- `close_date` (optional) — target close date
- `products` (optional) — comma-separated Adobe products in scope

---

## Step 2 — Run /sales prospect (blocking, first)

Invoke the `sales-prospect` skill for `<url>`. This is the only blocking step — do not proceed to Step 3 until it completes.

Save the output as:
```
__DEALS_ROOT__/<account>-<year>/PROSPECT-ANALYSIS.md
```

From the output, extract:
- **Company name** — the account name as it appears on line 1 of the output
- **Account slug** — lowercase, hyphens only, with current year appended (e.g. "Applied Materials" → `applied-materials-2026`)

The deal folder name is always `<slug>-<current-year>`. If a folder with that name already exists under the deals root, use it (this is a refresh run).

If a `deal.yaml` file already exists in the deal folder root, read it first and treat its contents as confirmed facts — incorporate ARR, stage, renewal date, and products from `deal.yaml` into the `content_map.json` fields, overriding any extracted values.

---

## Step 3 — Create deal folder structure

```
__DEALS_ROOT__/<account>-<year>/
__DEALS_ROOT__/<account>-<year>/content/
__DEALS_ROOT__/<account>-<year>/artifacts/
```

If the folder already exists (active deal in progress), treat this run as a refresh — research files will be overwritten, prior PPT will be timestamped before replacement. Preserve any `deal.yaml` or `brief.md` already in the folder root — do not overwrite them.

---

## Step 4 — Run 5 sub-skills in parallel

Using the Agent tool, spawn **5 parallel agents**. Each follows the corresponding sales skill for `<url>` and saves its output to `content/`:

| Agent | Skill to invoke | Output file |
|-------|----------------|-------------|
| A | `sales-qualify` | `__DEALS_ROOT__/<account>-<year>/content/LEAD-QUALIFICATION.md` |
| B | `sales-contacts` | `__DEALS_ROOT__/<account>-<year>/content/DECISION-MAKERS.md` |
| C | `sales-competitors` | `__DEALS_ROOT__/<account>-<year>/content/COMPETITIVE-INTEL.md` |
| D | `sales-prep` | `__DEALS_ROOT__/<account>-<year>/content/MEETING-PREP.md` |
| E | `sales-outreach` | `__DEALS_ROOT__/<account>-<year>/content/OUTREACH-SEQUENCE.md` |

Wait for all 5 agents to complete before proceeding.

---

## Step 5 — Build content_map.json

Read all 6 markdown files from `content/`. For each field below, extract the relevant section and distil it to fit the character budget. **Output only the distilled text — no markdown syntax, no bullet prefixes unless explicitly noted, no preamble.**

Apply these rules when distilling:
- Preserve complete sentences and specific detail; the PPT shapes use normAutofit so space is not a hard constraint
- Keep named people, specific numbers, dollar amounts, competitor names, dates, and concrete facts — these are the most valuable signals
- Remove filler phrases ("it is important to note", "additionally", "in conclusion", "it is worth mentioning")
- Truncate only when the field genuinely exceeds the budget after removing filler — never truncate to hit the budget artificially
- For bullet lists: use `•` as prefix, one item per line; include all substantive bullets up to the stated count

### content_map.json schema

```json
{
  "company_name": "<exact company name, 60 chars max>",
  "date": "<YYYY-MM-DD today>",

  "slide_2": {
    "business_issue":    "<core business problem 2-3 sentences, 600 chars max>",
    "big_idea":          "<value statement from top talking point, 400 chars max>",
    "company_objectives":"<4-6 strategic objectives as bullet list, 600 chars max>",
    "challenges":        "<top 4-6 challenges as bullet list, 600 chars max>",
    "adobe_solution":    "<Adobe differentiated solution positioning, 600 chars max>"
  },

  "slide_3": {
    "account_background":    "<Revenue · Employees · Business Focus · Industry · HQ — 4-6 lines, 700 chars max>",
    "account_background_lob":"<secondary division or LOB overview, or repeat account_background, 600 chars max>",
    "account_intel":         "<why this account, whitespace opportunity, Tier 1 rationale, 600 chars max>",
    "adobe_strengths":       "<Adobe advantages specific to this account, 5 bullets, 600 chars max>",
    "opportunities":         "<specific opportunity areas to create value, 5 bullets, 600 chars max>"
  },

  "slide_4": {
    "upcoming_renewals": "<if --arr and --renewal provided: '• [Products] | $[ARR] | [Renewal quarter] | Renewal'; else '[FILL: Bullet renewals — Solution | $Value | Renewal Date | Driver. Source: Clari/Panorama.]'>",
    "renewal_strategy":  "[FILL: Renewal strategy and risk. Source: Panorama / Clari.]"
  },

  "slide_5": {
    "market_trends":          "<broad industry/market developments, 5-6 bullets, 800 chars max>",
    "company_goals":          "<strategic goals, 5-6 bullets, 600 chars max>",
    "digital_priorities":     "<digital and technology priorities, 5-6 bullets, 600 chars max>",
    "market_opportunities":   "<specific value/growth opportunity areas, 5 bullets, 600 chars max>",
    "customer_challenges":    "<key obstacles in customer's own language, 5 bullets, 600 chars max>",
    "partner_strategy":       "<partner ecosystem and relationships, 3-4 bullets, 400 chars max>",
    "implementation_partners":"<known implementation/consulting partners, 400 chars max>"
  },

  "slide_6": {
    "buying_committee": [
      {"name": "<full name>", "title": "<job title>", "role": "<Economic Buyer|Champion|Technical Evaluator|End User|Blocker|Coach>", "attitude": "<Positive|Neutral|Negative|Unknown>"},
      "... up to 10 contacts from DECISION-MAKERS.md Buying Committee Map ..."
    ]
  },

  "slide_7": {
    "big_idea":       "<big idea in customer language, 3-4 sentences, 650 chars max>",
    "tagline":        "<one powerful sentence capturing the opportunity, 250 chars max>",
    "business_issue": "<business issue impacting performance and goals, 600 chars max>",
    "portfolio_plays":"<Adobe portfolio/sales plays to pitch with alignment rationale, 600 chars max>",
    "path_to_value":  "<path to value and budget alignment; include pipeline estimate if known, 600 chars max>"
  },

  "slide_8": {
    "opportunities_summary": "<if --arr, --stage, --close-date provided: format as 'Opp: [Company] ALM | ARR: $[X] | Stage: [stage] | Close: [date] | Next: [next step]'; else '[FILL: Complete in Clari. Paste pipeline table: Opp Name | Solution | ARR | Stage | Close Date | Next Step.]'>"
  },

  "slide_9": {
    "touchpoints_h1": "<key touchpoints Dec–Jul — events, renewal meetings, CEC; include --renewal and --close-date if provided; 500 chars max>",
    "touchpoints_h2": "<key touchpoints Aug–Nov — MAX, renewal close, EBC; 500 chars max>"
  },

  "slide_11": {
    "goals":            "<company transformational goals, 2-3 sentences, 400 chars max>",
    "challenges":       "<selected customer challenges, 2-3 sentences, 400 chars max>",
    "initiatives":      "<company priorities and how Adobe aligns, 2-3 sentences, 400 chars max>",
    "impact":           "<digital initiatives and expected Adobe impact, 2-3 sentences, 400 chars max>",
    "big_idea_paragraph":"<full Big Idea paragraph in customer language, 900 chars max>",
    "tagline":           "<tagline, 200 chars max>"
  }
}
```

**Source mapping per field:**

| Field | Primary source | Secondary source |
|-------|---------------|-----------------|
| slide_2.business_issue | PROSPECT-ANALYSIS §Executive Summary | LEAD-QUALIFICATION §Need Analysis |
| slide_2.big_idea | MEETING-PREP §Talking Points (top 1) | — |
| slide_2.company_objectives | PROSPECT-ANALYSIS §Company Profile | MEETING-PREP §Business Situation |
| slide_2.challenges | PROSPECT-ANALYSIS §Executive Summary (red flags) | LEAD-QUALIFICATION §Red Flags |
| slide_2.adobe_solution | COMPETITIVE-INTEL §Competitive Positioning Statements | MEETING-PREP §Competitive Context |
| slide_3.account_background | PROSPECT-ANALYSIS §Prospect Snapshot table | PROSPECT-ANALYSIS §Company Profile |
| slide_3.account_intel | PROSPECT-ANALYSIS §Opportunity Assessment | LEAD-QUALIFICATION §Opportunity Quality |
| slide_3.adobe_strengths | COMPETITIVE-INTEL §Win Patterns | MEETING-PREP §Competitive Context |
| slide_3.opportunities | LEAD-QUALIFICATION §Opportunity Quality Score | PROSPECT-ANALYSIS §Executive Summary |
| slide_5.market_trends | PROSPECT-ANALYSIS §Company Profile | MEETING-PREP §Business Situation |
| slide_5.company_goals | PROSPECT-ANALYSIS §Company Profile | MEETING-PREP §Business Situation |
| slide_5.digital_priorities | MEETING-PREP §Business Situation | PROSPECT-ANALYSIS §Company Profile |
| slide_5.customer_challenges | MEETING-PREP §Discovery Questions (Listen For) | LEAD-QUALIFICATION §Need Analysis |
| slide_5.partner_strategy | DECISION-MAKERS §Multi-Threading Strategy | COMPETITIVE-INTEL §Current Solutions |
| slide_6.buying_committee | DECISION-MAKERS §Buying Committee Map | — |
| slide_7.big_idea | MEETING-PREP §Talking Points | LEAD-QUALIFICATION §Pain Point Analysis |
| slide_7.portfolio_plays | COMPETITIVE-INTEL §Recommended Competitive Strategy | MEETING-PREP §Talking Points |
| slide_7.path_to_value | LEAD-QUALIFICATION §BANT (Budget Analysis) | MEETING-PREP §Success Metrics |
| slide_9.touchpoints_h1 | MEETING-PREP §Success Metrics | LEAD-QUALIFICATION §Timeline Analysis |
| slide_11.goals | PROSPECT-ANALYSIS §Company Profile | MEETING-PREP §Business Situation |
| slide_11.big_idea_paragraph | All sources — synthesized | — |

Write the completed JSON to:
```
__DEALS_ROOT__/<account>-<year>/content/content_map.json
```

---

## Step 6 — Run PPT generator

Run this exact command:

```bash
__PYTHON_CMD__ "__PPT_SCRIPT__" \
  --content-map "__DEALS_ROOT__/<account>-<year>/content/content_map.json" \
  --template "__TEMPLATE_PATH__" \
  --output "__DEALS_ROOT__/<account>-<year>/artifacts"
```

Capture stdout/stderr. If the script exits with an error, diagnose and fix before reporting completion.

---

## Step 7 — Print completion summary

```
✓ /sales plan complete — <Company Name>
─────────────────────────────────────────────────────────
Deal folder:   <slug>/
PPT:           artifacts/<Company>-Business-Plan-<date>.pptx
Research:      content/ (6 files + content_map.json)

Auto-populated:  Slides 1, 2, 3, 5, 6, 7, 9, 11
Manual fill:     Slide 4 (Performance History — source: Panorama/Clari)
                 Slide 8 (FY Opportunities — source: Clari)
                 Slide 9 (add specific event dates)
```

If `--arr`, `--renewal`, or `--stage` were passed, remove those slides from the manual fill list as appropriate.

---

## Error handling

| Scenario | Action |
|----------|--------|
| URL unreachable | Report to user, suggest alternate URL, stop |
| `sales-qualify` fails | Write `[FILL: /sales qualify unavailable]` for slide_2.challenges, slide_3.account_intel, slide_7.path_to_value, slide_9.touchpoints — continue |
| `sales-contacts` fails | Write `[FILL: /sales contacts unavailable]` for slide_6.buying_committee — continue |
| `sales-competitors` fails | Write `[FILL: /sales competitors unavailable]` for slide_2.adobe_solution, slide_3.adobe_strengths, slide_7.portfolio_plays — continue |
| `sales-prep` fails | Write `[FILL: /sales prep unavailable]` for slide_2.big_idea, slide_5.digital_priorities, slide_7.big_idea, slide_9.touchpoints — continue |
| `sales-outreach` fails | Log failure, no slide zones affected — continue |
| Python script fails | Print full error, diagnose root cause, attempt fix (e.g. missing python-pptx — run `pip install python-pptx`) |
| Template file not found | Report exact path checked, ask user to confirm file exists at `__TEMPLATE_PATH__` |
| content_map.json field exceeds budget after summarization | Truncate to budget and append `…` |

---

## Notes

- The source PPT template is **never modified** — the script always clones it.
- If re-running for the same account, prior PPT is renamed with a timestamp (e.g. `Company-Business-Plan-2026-05-07_1430.pptx`) before the new one is written.
- The `content_map.json` is a useful audit trail — it shows exactly what Claude extracted and compressed for each slide zone.
- To re-run only the PPT step (after manually editing markdown files), update `content_map.json` and re-run Step 6 directly.
