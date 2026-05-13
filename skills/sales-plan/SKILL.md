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

Expected `deal.yaml` keys (all optional):
```yaml
arr: 150000               # current/target ARR in dollars
renewal: "Q3 FY26"        # renewal quarter
stage: POC                # deal stage
close_date: 2026-08-31    # ISO date
products: "ALM,Marketo"   # comma-separated Adobe products
champion: "Marie Gabriel" # confirmed champion name
notes: |                  # free-text AE notes — treat as ground truth
  …
```

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

Using the Agent tool, spawn **5 parallel agents**. Each follows the corresponding sales skill for `<url>` and saves its output to `content/`. Brief each agent with the content areas its output will feed in Step 5:

| Agent | Skill | Output file | Content_map fields this feeds |
|-------|-------|-------------|-------------------------------|
| A | `sales-qualify` | `…/content/LEAD-QUALIFICATION.md` | slide_2.challenges, slide_2.business_issue, slide_3.account_intel, slide_7.path_to_value, slide_9.h1/h2; focus on BANT, pain points, timeline triggers |
| B | `sales-contacts` | `…/content/DECISION-MAKERS.md` | slide_6.buying_committee, slide_3.account_background_lob; map full buying committee with name/title/role/reason/attitude; capture org context for the L&D/HR division |
| C | `sales-competitors` | `…/content/COMPETITIVE-INTEL.md` | slide_2.adobe_solution, slide_3.adobe_strengths, slide_7.portfolio_plays, slide_7.like_customers; identify current LMS/LXP/HCM stack, peer companies that deployed ALM, ALM differentiators |
| D | `sales-prep` | `…/content/MEETING-PREP.md` | slide_2.big_idea, slide_5.digital_priorities, slide_7.big_idea, slide_7.tagline, slide_7.business_issue; extract the big idea narrative, digital transformation priorities, and timeline pressures |
| E | `sales-outreach` | `…/content/OUTREACH-SEQUENCE.md` | No direct content_map fields; used as supplemental context for tone and messaging in Step 5 |

> Note: replace `…` with `__DEALS_ROOT__/<account>-<year>`. Wait for all 5 agents to complete before proceeding.

---

## Step 5 — Build content_map.json

Read all 6 markdown files from `content/`. For each field below, extract and synthesize content through the lens of **where Adobe (primarily Adobe Learning Manager / ALM, and the broader Adobe DX portfolio) creates value for this specific account**. Do not summarize research neutrally — extract and frame content as it would be used in a live customer conversation or executive briefing. Every field should read as if an experienced AE wrote it, not as a research summary.

Apply these rules when distilling:
- Preserve complete sentences and specific detail; the PPT shapes use normAutofit so space is not a hard constraint
- Keep named people, specific numbers, dollar amounts, competitor names, dates, and concrete facts — these are the most valuable signals
- Remove filler phrases ("it is important to note", "additionally", "in conclusion", "it is worth mentioning")
- Truncate only when the field genuinely exceeds the budget after removing filler — never truncate to hit the budget artificially
- For bullet lists: use `•` as prefix, one item per line; include all substantive bullets up to the stated count
- For fields without a clear match in the source, synthesize from the closest available content — never leave a field blank or write [FILL] unless the schema explicitly says to

### content_map.json schema

```json
{
  "company_name":   "<exact company name as it appears on their website, 60 chars max>",
  "company_domain": "<company domain for logo fetch, e.g. 'appliedmaterials.com'. Extract from the prospect URL. No protocol, no path, no www prefix.>",
  "date": "<YYYY-MM-DD today>",

  "slide_2": {
    "business_issue":    "<The primary business challenge creating urgency to buy. Frame from the economic buyer's perspective — specific, quantified where possible, tied to business outcomes. Do NOT mention Adobe. 2-3 sentences, 600 chars max>",
    "big_idea":          "<The single strongest reason this account should partner with Adobe — in customer language, connecting their business goals to Adobe's capabilities. One powerful assertion, not a feature list. 400 chars max>",
    "company_objectives":"<Strategic objectives where Adobe DX / ALM can accelerate progress — prioritize objectives tied to workforce development, digital skills, content operations, or employee experience. Include objective name + brief context. 4-6 bullets, 600 chars max>",
    "challenges":        "<Top challenges Adobe can directly address — prioritize talent development gaps, L&D technology debt, skills deficits, compliance training needs. Use the account's own language where available. 4-6 bullets, 600 chars max>",
    "adobe_solution":    "<How Adobe's portfolio (ALM, Experience Cloud, Marketo) directly addresses the identified challenges. Name specific products and the capability that maps to each pain point. Include competitive differentiators vs. known incumbent. 600 chars max>"
  },

  "slide_3": {
    "account_background":    "<MUST follow this exact 10-line labeled format — one field per line, no blank lines:\nAnnual Revenue: [exact figure with source year]\nOnline Revenue: [figure or N/A]\nSize: [X employees]\nBusiness Focus & Major Divisions: [key segments]\nIndustry & Position: [industry + market rank]\nGeneral digital maturity: [low/medium/high + 1-sentence context]\nGeography: [HQ + key operating regions]\nCompetitive Footprint: [key tech vendors/platforms in use]\nAzure or AWS Commitments: [cloud provider + contract scale if known]\nPartner Footprint: [key SI/consulting partners]\n700 chars max>",
    "account_background_lob":"<The specific LOB, division, or workforce segment most relevant to an ALM or Adobe DX sale (e.g., HR/L&D org, AGS field services, IT, Marketing COE). MUST use the same 10-line labeled format as account_background, scoped to this specific division:\nAnnual Revenue: [segment revenue or N/A]\nOnline Revenue: [N/A if internal L&D]\nSize: [employees in this division]\nBusiness Focus & Major Divisions: [this LOB's mandate]\nIndustry & Position: [how this LOB ranks vs peers]\nGeneral digital maturity: [maturity of this specific org]\nGeography: [where this division operates]\nCompetitive Footprint: [LMS/LXP/HCM tools in use by this LOB]\nAzure or AWS Commitments: [relevant cloud commitments for this LOB]\nPartner Footprint: [vendors or SIs serving this LOB]\n600 chars max>",
    "account_intel":         "<Sales rationale for pursuing this account now — whitespace (what Adobe doesn't have yet), trigger events (leadership change, transformation initiative, contract event), and account tier/priority rationale. Be specific. 600 chars max>",
    "adobe_strengths":       "<Adobe's specific advantages for winning at this account — tailored to their industry, tech stack, and known pain points. Include proof points (customer references in same industry, specific ALM capabilities they need). Not generic Adobe strengths. 5 bullets, 600 chars max>",
    "opportunities":         "<Specific Adobe product plays and expansion opportunities. Each bullet names the product, the use case, and the business impact — e.g., '• ALM — replace legacy LMS for global manufacturing workforce, 15k+ learners'. 5 bullets, 600 chars max>"
  },

  "slide_4": {
    "upcoming_renewals": "<if --arr and --renewal provided: '• [Products] | $[ARR] | [Renewal quarter] | Renewal'; else '[FILL: Bullet renewals — Solution | $Value | Renewal Date | Driver. Source: Clari/Panorama.]'>",
    "renewal_strategy":  "[FILL: Renewal strategy and risk. Source: Panorama / Clari.]"
  },

  "slide_5": {
    "market_trends":          "<Industry-level trends creating tailwinds for Adobe's solutions — AI in L&D, skills-based organizations, digital workforce transformation, compliance mandates. Each bullet = trend + why it matters for this account specifically. 5-6 bullets, 800 chars max>",
    "company_goals":          "<This account's stated strategic goals for the current fiscal year — from earnings calls, investor presentations, or public announcements. Prioritize goals where Adobe can demonstrate measurable impact. 5-6 bullets, 600 chars max>",
    "digital_priorities":     "<Specific digital and technology initiatives this account is investing in — systems being modernized, platforms being evaluated, transformation programs underway. Flag any where Adobe competes or complements. 5-6 bullets, 600 chars max>",
    "market_opportunities":   "<White-space and expansion opportunities Adobe can capture — untapped business units, geographies, or use cases not currently served by Adobe. Anchor each to a specific Adobe product and estimated impact. 5 bullets, 600 chars max>",
    "customer_challenges":    "<Obstacles this account faces in their own words — from earnings calls, press releases, job postings, or analyst coverage. Frame as pain statements an economic buyer would recognize, not symptoms. 5 bullets, 600 chars max>",
    "partner_strategy":       "<Known SI, consulting, or technology partners this account works with. Flag Adobe partners (Accenture, Deloitte, Infosys) vs. competitive partners. Include partner name and engagement context. 3-4 bullets, 400 chars max>",
    "implementation_partners":"<Specific consulting/SI firms engaged at this account for relevant projects (HCM, L&D, digital transformation). Include partner name and engagement context. 400 chars max>"
  },

  "slide_6": {
    "buying_committee": [
      {
        "name":     "<full name>",
        "title":    "<exact job title>",
        "role":     "<Economic Buyer|Champion|Technical Evaluator|End User|Blocker|Coach>",
        "reason":   "<1-2 sentences: why this person matters to the deal — their specific pain point, their relationship to the decision, their key question to answer, or their influence over budget/selection>",
        "attitude": "<Positive|Neutral|Negative|Unknown>"
      },
      "... up to 10 contacts from DECISION-MAKERS.md Buying Committee Map ..."
    ]
  },

  "slide_7": {
    "big_idea":       "<The central value thesis — 3-4 sentences connecting this account's most urgent business challenge to Adobe's differentiated solution, with a clear outcome statement. Written in language the economic buyer would use, not product marketing language. 650 chars max>",
    "tagline":        "<A single memorable sentence capturing the opportunity — sharp enough to be a meeting title or email subject line. 250 chars max>",
    "business_issue": "<The specific business issue driving urgency for a decision — quantified impact, named stakeholders affected, timeline pressure if known. This is the 'so what' that justifies executive attention. 600 chars max>",
    "portfolio_plays":"<The specific Adobe products/plays to lead with and why — mapped to the identified business issue. Include the sales motion (new logo, upsell, competitive displacement) and the proof point or use case to anchor the pitch. 600 chars max>",
    "path_to_value":  "<How Adobe delivers ROI for this account — include budget indicators, decision timeline, potential ARR, and the first value milestone (quick win). Connect to the economic buyer's success metrics. 600 chars max>",
    "like_customers": "<2-3 peer companies (same industry or use case) that have deployed ALM or Adobe DX in comparable scenarios — include company name, the specific use case, and the outcome or proof point. Anchor credibility for the pitch. 400 chars max>"
  },

  "slide_8": {
    "opportunities_summary": "<if --arr, --stage, --close-date provided: format as 'Opp: [Company] ALM | ARR: $[X] | Stage: [stage] | Close: [date] | Next: [next step]'; else '[FILL: Complete in Clari. Paste pipeline table: Opp Name | Solution | ARR | Stage | Close Date | Next Step.]'>"
  },

  "slide_9": {
    "h1": {
      "december":  "<1-2 short touchpoints for December — e.g. 'CEC attendance; renewal strategy kick-off with champion'. 80 chars max>",
      "january":   "<1-2 short touchpoints for January — e.g. 'Executive roundtable; QBR'. 80 chars max>",
      "february":  "<1-2 short touchpoints for February — e.g. 'Product demo; discovery session'. 80 chars max>",
      "march":     "<1-2 short touchpoints for March — e.g. 'Adobe Summit attendance; exec dinner'. 80 chars max>",
      "april":     "<1-2 short touchpoints for April — e.g. 'POC scope review; business case draft'. 80 chars max>",
      "may":       "<1-2 short touchpoints for May — e.g. 'Business case to CHRO; budget alignment'. 80 chars max>",
      "june":      "<1-2 short touchpoints for June — e.g. 'Contract review; procurement kick-off'. 80 chars max>",
      "july":      "<1-2 short touchpoints for July — e.g. 'Target close / renewal; EBC if needed'. 80 chars max>"
    },
    "h2": {
      "august":    "<1-2 short touchpoints for August. 80 chars max>",
      "september": "<1-2 short touchpoints for September — e.g. 'QBR; user adoption review'. 80 chars max>",
      "october":   "<1-2 short touchpoints for October — e.g. 'Adobe MAX attendance'. 80 chars max>",
      "november":  "<1-2 short touchpoints for November — e.g. 'FY budget alignment; expansion conversation'. 80 chars max>"
    }
  },

  "slide_11": {
    "goals":            "<The 2-3 most important company goals where Adobe creates measurable impact. Written as assertions, not list items. 2-3 sentences, 400 chars max>",
    "challenges":       "<The 2-3 most acute challenges this account faces that Adobe directly addresses. Use the account's own language. 2-3 sentences, 400 chars max>",
    "initiatives":      "<The 2-3 strategic initiatives underway where Adobe should be positioned as a key enabler — include the initiative name and Adobe's specific role. 2-3 sentences, 400 chars max>",
    "impact":           "<Concrete business outcomes Adobe drives for this account — quantified where possible (e.g., 'reduce time-to-competency by 40%', 'unify 50k learners on a single platform'). 2-3 sentences, 400 chars max>",
    "big_idea_paragraph":"<A fully written Big Idea paragraph for exec communications — combines goals, challenges, Adobe solution, and expected outcome into a cohesive narrative. Written in customer-facing language as if an AE is speaking to the CFO or CHRO. 900 chars max>",
    "tagline":           "<The sharpest possible single-sentence summary of why Adobe and this account should partner. Can be used as a meeting opener. 200 chars max>"
  }
}
```

**Source mapping per field:**

| Field | Primary source | Secondary source |
|-------|---------------|-----------------|
| company_domain | Prospect URL (extract domain) | — |
| slide_2.business_issue | PROSPECT-ANALYSIS §Executive Summary | LEAD-QUALIFICATION §Need Analysis |
| slide_2.big_idea | MEETING-PREP §Talking Points (top 1) | — |
| slide_2.company_objectives | PROSPECT-ANALYSIS §Company Profile | MEETING-PREP §Business Situation |
| slide_2.challenges | PROSPECT-ANALYSIS §Executive Summary (red flags) | LEAD-QUALIFICATION §Red Flags |
| slide_2.adobe_solution | COMPETITIVE-INTEL §Competitive Positioning Statements | MEETING-PREP §Competitive Context |
| slide_3.account_background | PROSPECT-ANALYSIS §Prospect Snapshot table | PROSPECT-ANALYSIS §Company Profile |
| slide_3.account_background_lob | DECISION-MAKERS §Buying Committee Map (org context) | PROSPECT-ANALYSIS §Company Profile |
| slide_3.account_intel | PROSPECT-ANALYSIS §Opportunity Assessment | LEAD-QUALIFICATION §Opportunity Quality |
| slide_3.adobe_strengths | COMPETITIVE-INTEL §Win Patterns | MEETING-PREP §Competitive Context |
| slide_3.opportunities | LEAD-QUALIFICATION §Opportunity Quality Score | PROSPECT-ANALYSIS §Executive Summary |
| slide_5.market_trends | PROSPECT-ANALYSIS §Company Profile | MEETING-PREP §Business Situation |
| slide_5.company_goals | PROSPECT-ANALYSIS §Company Profile | MEETING-PREP §Business Situation |
| slide_5.digital_priorities | MEETING-PREP §Business Situation | PROSPECT-ANALYSIS §Company Profile |
| slide_5.customer_challenges | MEETING-PREP §Discovery Questions (Listen For) | LEAD-QUALIFICATION §Need Analysis |
| slide_5.partner_strategy | DECISION-MAKERS §Multi-Threading Strategy | COMPETITIVE-INTEL §Current Solutions |
| slide_6.buying_committee | DECISION-MAKERS §Buying Committee Map | PROSPECT-ANALYSIS §Key Stakeholders |
| slide_7.big_idea | MEETING-PREP §Talking Points | LEAD-QUALIFICATION §Pain Point Analysis |
| slide_7.portfolio_plays | COMPETITIVE-INTEL §Recommended Competitive Strategy | MEETING-PREP §Talking Points |
| slide_7.path_to_value | LEAD-QUALIFICATION §BANT (Budget Analysis) | MEETING-PREP §Success Metrics |
| slide_7.like_customers | COMPETITIVE-INTEL §Current Solutions | MEETING-PREP §Proof Points |
| slide_9.h1/h2 | MEETING-PREP §Success Metrics | LEAD-QUALIFICATION §Timeline Analysis |
| slide_11.goals | PROSPECT-ANALYSIS §Company Profile | MEETING-PREP §Business Situation |
| slide_11.big_idea_paragraph | All sources — synthesized | — |

Write the completed JSON to:
```
__DEALS_ROOT__/<account>-<year>/content/content_map.json
```

Run this one-liner to validate the file before proceeding:
```bash
__PYTHON_CMD__ -c "import json,sys; data=json.load(open('__DEALS_ROOT__/<account>-<year>/content/content_map.json')); required=['company_name','company_domain','slide_2','slide_3','slide_5','slide_6','slide_7','slide_9','slide_11']; missing=[k for k in required if k not in data]; print('WARN missing:',missing) if missing else print('JSON OK')"
```

If any required keys are missing, fill them before continuing to Step 6.

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

Auto-populated:  Slides 1 (title + logo), 2, 3, 5, 6, 7, 9, 11
Manual fill:     Slide 4 (Performance History — source: Panorama/Clari)
                 Slide 8 (FY Opportunities — source: Clari)
```

If `--arr`, `--renewal`, or `--stage` were passed, remove those slides from the manual fill list as appropriate.

---

## Error handling

| Scenario | Action |
|----------|--------|
| URL unreachable | Report to user, suggest alternate URL, stop |
| `sales-qualify` fails | Write `[FILL: /sales qualify unavailable]` for slide_2.challenges, slide_3.account_intel, slide_7.path_to_value, slide_9.h1/h2 — continue |
| `sales-contacts` fails | Write `[FILL: /sales contacts unavailable]` for slide_6.buying_committee — continue |
| `sales-competitors` fails | Write `[FILL: /sales competitors unavailable]` for slide_2.adobe_solution, slide_3.adobe_strengths, slide_7.portfolio_plays — continue |
| `sales-prep` fails | Write `[FILL: /sales prep unavailable]` for slide_2.big_idea, slide_5.digital_priorities, slide_7.big_idea, slide_9.h1/h2 — continue |
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
