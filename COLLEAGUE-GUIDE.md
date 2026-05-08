# /sales plan — Colleague Guide

**What this is:** A Claude Code skill that researches any account and auto-populates the Adobe Lite DX Business Plan PowerPoint template. You type one command and get a populated, presentation-ready deck in your deal folder.

**Who this is for:** Adobe AEs using Claude Code who want to cut account plan prep time from hours to minutes.

---

## Table of Contents

1. [What You Get](#what-you-get)
2. [Before You Start — Prerequisites](#before-you-start)
3. [Installation Walkthrough](#installation-walkthrough)
4. [Running Your First Plan](#running-your-first-plan)
5. [Understanding the Output](#understanding-the-output)
6. [Slides Reference — What's Auto-Populated vs. Manual](#slides-reference)
7. [Using Optional Flags for Deal Context](#using-optional-flags)
8. [Re-Running for an Existing Account](#re-running)
9. [Best Practices](#best-practices)
10. [Troubleshooting](#troubleshooting)
11. [Quick Reference Card](#quick-reference-card)

---

## What You Get

One command produces:

```
deals/snowflake-2026/
├── PROSPECT-ANALYSIS.md          ← full account research
├── content/
│   ├── LEAD-QUALIFICATION.md     ← BANT + MEDDIC scorecard
│   ├── DECISION-MAKERS.md        ← buying committee map
│   ├── COMPETITIVE-INTEL.md      ← incumbent stack + competitive positioning
│   ├── MEETING-PREP.md           ← talking points, discovery questions
│   ├── OUTREACH-SEQUENCE.md      ← 3-email sequence, LinkedIn plays
│   └── content_map.json          ← audit trail of what went into each slide
└── artifacts/
    └── Snowflake-Business-Plan-2026-05-07.pptx   ← your deck
```

**The PowerPoint has 8 slides pre-populated from the research.** Two slides (Performance History and Opportunities) require internal data from Clari/Panorama — placeholders are inserted so you know exactly what to fill in.

The original template is never touched. The script always works from a clone.

---

## Before You Start

You need four things installed before `/sales plan` will work. Go through this checklist in order.

### Checklist

- [ ] **Claude Code** — the CLI tool (not the web app)
- [ ] **AI Sales Team skills** — the base `/sales` skills this orchestrates
- [ ] **Python 3.8+** — for PPT generation
- [ ] **This skill** — `/sales plan` itself

### 1. Claude Code

Download and install from **https://claude.ai/download**. You want the desktop app or the CLI — not claude.ai in a browser.

After installing, open a terminal and verify:
```
claude --version
```

If you see a version number, you're good.

### 2. AI Sales Team Skills

This is the base layer that provides `/sales prospect`, `/sales qualify`, `/sales contacts`, and the other research skills that `/sales plan` orchestrates.

**Install with one command (Mac / Git Bash on Windows):**
```bash
curl -fsSL https://raw.githubusercontent.com/zubair-trabzada/ai-sales-team-claude/main/install.sh | bash
```

**Or on Windows PowerShell:**
```powershell
git clone https://github.com/zubair-trabzada/ai-sales-team-claude.git
cd ai-sales-team-claude
.\install.ps1
```

### 3. Python 3.8+

**Mac:** Python is usually pre-installed. Check with `python3 --version`. If you need to install or upgrade: https://www.python.org/downloads/

**Windows:** Download from https://www.python.org/downloads/ — check "Add Python to PATH" during installation.

After installing, verify in terminal:
```
python3 --version   (Mac)
python --version    (Windows)
```

### 4. This Skill

See [Installation Walkthrough](#installation-walkthrough) below.

---

## Installation Walkthrough

### Mac / Linux / Git Bash on Windows

Open a terminal and run:

```bash
git clone https://github.com/YOUR_REPO_URL/sales-plan-skill.git
cd sales-plan-skill
chmod +x install.sh
./install.sh
```

### Windows PowerShell

```powershell
git clone https://github.com/YOUR_REPO_URL/sales-plan-skill.git
cd sales-plan-skill
.\install.ps1
```

### What the installer does (step by step)

**Step 1 — Checks prerequisites**

The installer checks for Claude Code and Python. If anything is missing, it tells you exactly what to install before continuing.

**Step 2 — Installs python-pptx**

The library that generates the PowerPoint. Installed automatically if not present.

**Step 3 — Asks for your template path**

```
Enter the full path to your Lite DX Business Plan Template .pptx file.
```

This is your personal copy of the template, synced from OneDrive. The installer tries to detect it automatically. If it finds it, it asks you to confirm. If not, paste the full path.

> **Finding your template:** Open File Explorer → OneDrive - Adobe → look for `Lite DX Business Plan Template - [YourName].pptx`. Right-click → Properties → copy the full path.

**Step 4 — Asks for your deals folder**

```
Enter your deals root directory (where account folders live).
Default: ~/Desktop/claude-workspace/deals
```

Press Enter to accept the default, or paste your own path. This is the folder where all your deal subfolders live (e.g., `deals/snowflake-2026/`, `deals/amat-2026/`, etc.).

**Step 5 — Installs and patches**

The skill and script are installed to `~/.claude/skills/sales-plan/` with your paths baked in. You won't need to edit anything manually.

**Step 6 — Checks for sub-skills**

The installer verifies that the AI Sales Team sub-skills are present. If any are missing, it shows you the install command.

### Successful install output

```
OK  Claude Code found
OK  Python found: python3 (Python 3.12.0)
OK  python-pptx installed
OK  Template found
OK  Deals root: /Users/yourname/Desktop/claude-workspace/deals

Installing...
OK  sales-plan skill installed
OK  sales_plan_ppt.py installed

Checking required sub-skills...
OK  sales-prospect
OK  sales-qualify
OK  sales-contacts
OK  sales-competitors
OK  sales-prep
OK  sales-outreach

Installation Complete!
```

---

## Running Your First Plan

Open Claude Code (the desktop app or terminal). Navigate to your working directory if you're using the CLI, or just open it.

### Basic run — full research

```
/sales plan https://www.snowflake.com
```

Claude will:
1. Research the company (expect 3–6 minutes for full parallel research)
2. Create `deals/snowflake-2026/` with subfolders
3. Build the content map
4. Generate the PPT

You'll see progress in the Claude Code chat as each research agent completes.

### With deal context

If you already know the ARR, renewal quarter, or deal stage, pass them as flags. These pre-fill the internal slides that can't be sourced from public research:

```
/sales plan https://www.snowflake.com --arr=150000 --renewal=Q3 --stage=POC --close-date=2026-08-31
```

### End-of-run output

```
✓ /sales plan complete — Snowflake, Inc.
────────────────────────────────────────────────────────
Deal folder:   deals/snowflake-2026/
PPT:           artifacts/Snowflake-Inc-Business-Plan-2026-05-07.pptx
Research:      content/ (6 files + content_map.json)

Auto-populated:  Slides 1, 2, 3, 5, 6, 7, 9, 11
Manual fill:     Slide 4 (Performance History — source: Panorama/Clari)
                 Slide 8 (FY Opportunities — source: Clari)
```

Open the `.pptx` from `artifacts/` and you're ready to review.

---

## Understanding the Output

### The research files (content/)

Six markdown files covering every angle of the account. Read them before customer meetings — they contain more depth than the PPT slides. The PPT is a compressed, presentation-ready version.

| File | What it contains |
|------|-----------------|
| `PROSPECT-ANALYSIS.md` | Full company profile, exec summary, decision maker map, opportunity assessment, score breakdown |
| `LEAD-QUALIFICATION.md` | BANT scorecard, MEDDIC analysis, specific pain evidence, timeline triggers |
| `DECISION-MAKERS.md` | Buying committee table, verified contact details, org chart, personalization anchors, email patterns |
| `COMPETITIVE-INTEL.md` | Current tech stack, incumbent LMS hypothesis, competitor threat ratings, ALM differentiation |
| `MEETING-PREP.md` | Talking points, discovery questions, competitive context, success metrics |
| `OUTREACH-SEQUENCE.md` | 3-email sequence, LinkedIn DM cadence, channel strategy, send timing |

### content_map.json

The intermediary between research and PPT. Claude reads all 6 files, compresses the relevant sections to the character budget for each slide zone, and writes this JSON. If a slide looks thin or off, this is the first place to check — you can edit it directly and re-run just the PPT step.

### The PPT (artifacts/)

A clone of your template with research content populated by shape name. Fonts, colors, and layout are untouched. What you get:

- **Ready to present:** Slides 1, 2, 3, 5, 6, 7, 9, 11 are fully populated
- **Labeled placeholders:** Slides 4 and 8 show `[FILL: ...]` labels where internal data goes
- **Buying committee:** Slide 6 has a clean table (Name · Title · Role · Attitude toward Adobe) with up to 10 contacts
- **Timestamp protection:** If you re-run, the prior PPT is renamed (e.g., `..._1430.pptx`) before the new one is written

---

## Slides Reference

### Auto-populated slides

**Slide 1 — Title**
Company name. Clean, no action needed.

---

**Slide 2 — Executive Summary**
Five zones populated from research:
- Business Issue — core problem the account faces
- Big Idea — lead value statement
- Company Objectives — 3–4 strategic goals as bullets
- Key Challenges — top obstacles (internal, competitive, or market)
- Adobe Differentiated Solution — why ALM specifically for this account

*Review this slide carefully — it's the most important and the most compressed. The full reasoning is in `PROSPECT-ANALYSIS.md` § Executive Summary.*

---

**Slide 3 — Company Overview**
Five text boxes:
- Account Background — revenue, headcount, industry, HQ, key segments
- LOB / Division Overview — major business unit or product line relevant to the deal
- Account Intelligence — whitespace rationale, why we're pursuing this account now
- Adobe Strengths — our specific advantages for this account
- Opportunities — concrete areas to create value

---

**Slide 4 — Performance History** *(manual fill required)*
Shows existing Adobe ARR, renewals, and deal history. Claude cannot source this from public research.

> **Where to get this data:** Panorama for current ARR and products. Clari for renewal dates and deal history. If you passed `--arr` and `--renewal` as flags, those are pre-populated here.

---

**Slide 5 — Market Trends & Goals**
Seven zones from research:
- Market Trends — industry / macro developments
- Company Goals — strategic priorities from earnings calls and investor material
- Digital Priorities — technology and digital transformation focus
- Market Opportunities — where we can create value
- Customer Challenges — pain in the customer's own language
- Partner Strategy — key partnerships and ecosystem
- Implementation Partners — known SIs and consulting partners

---

**Slide 6 — Org Chart (Buying Committee)**
A clean table inserted by the script:

| Name | Title | Role | Attitude toward Adobe |
|------|-------|------|-----------------------|
| Person A | SVP, Engineering | Economic Buyer | Unknown |
| Person B | Director, L&D | Champion | Positive |

Populated from `DECISION-MAKERS.md`. Up to 10 contacts. Roles are one of: Economic Buyer, Champion, Technical Evaluator, End User, Blocker, Coach.

> **Verify before presenting.** Contact data comes from public research (LinkedIn, company site, news). Cross-check against Sales Nav or your own knowledge before using contact names in customer conversations.

---

**Slide 7 — Value Strategy**
Five zones:
- Big Idea — the core value narrative in customer language
- Tagline — one sentence capturing the opportunity
- Business Issue — what's driving urgency
- Portfolio Plays — which Adobe products to lead with and why
- Path to Value — how we get to closed/won, with ARR estimate if known

---

**Slide 8 — FY Opportunities** *(manual fill required)*
Open pipeline from Clari. If you passed `--arr`, `--stage`, and `--close-date`, those are pre-populated here in the format:

```
Opp: Snowflake ALM | ARR: $150,000 | Stage: POC | Close: 2026-08-31 | Next: Technical demo
```

Otherwise the placeholder instructs you to paste from Clari.

---

**Slide 9 — FY Timeline**
H1 and H2 touchpoints — key events, renewal milestones, EBC/CEC dates. Populated from research where conference calendars and timing are publicly available. Add specific event dates you know from your account calendar.

---

**Slide 11 — Big Idea (Appendix)**
Full Big Idea narrative for deeper conversations:
- Goals, Challenges, Initiatives, Impact — four rows
- Full Big Idea paragraph in customer language
- Tagline

The most synthesis-heavy slide. Review the paragraph for tone — it should read like something the customer would say about their own situation, not an Adobe pitch.

---

## Using Optional Flags

Flags let you pre-populate the internal slides that Claude can't source from public research.

### --arr

The current or target annual recurring revenue for this account.

```
/sales plan https://www.snowflake.com --arr=150000
```

Populates the ARR field on Slide 4 (Performance History) and Slide 8 (Opportunities).

### --renewal

The renewal quarter. Use `Q1`, `Q2`, `Q3`, or `Q4`. Optionally include the fiscal year.

```
/sales plan https://www.snowflake.com --renewal=Q3
/sales plan https://www.snowflake.com --renewal="Q3 FY26"
```

Populates the renewal date on Slide 4 and the H1/H2 timeline on Slide 9.

### --stage

The current deal stage. Free text — use whatever your CRM uses.

```
/sales plan https://www.snowflake.com --stage=POC
/sales plan https://www.snowflake.com --stage="Stage 3 - Eval"
```

Populates Slide 8 (Opportunities table).

### --close-date

Target close date in YYYY-MM-DD format.

```
/sales plan https://www.snowflake.com --close-date=2026-08-31
```

Populates Slide 8 and the timeline on Slide 9.

### --products

Comma-separated list of Adobe products in scope. Useful when you're presenting a multi-product plan.

```
/sales plan https://www.snowflake.com --products="ALM,Marketo"
```

Populates the products field on Slide 4 (Upcoming Renewals).

### All flags together

```
/sales plan https://www.snowflake.com \
  --arr=150000 \
  --renewal=Q3 \
  --stage=POC \
  --close-date=2026-08-31 \
  --products="ALM,Marketo"
```

When all five flags are passed, Slides 4 and 8 are fully populated and removed from the manual fill list.

---

## Re-Running for an Existing Account

Running `/sales plan` on an account that already has a deal folder is safe and intentional — use it to refresh research before QBRs or after a significant trigger event (earnings, leadership change, new product announcement).

**What happens on a refresh run:**
- All 6 research files in `content/` are rewritten with fresh data
- The existing PPT is renamed with a timestamp: `Snowflake-Business-Plan-2026-05-07_1430.pptx`
- A new PPT is written: `Snowflake-Business-Plan-2026-05-07.pptx`
- `deal.yaml` and `brief.md` in the folder root are **never touched**

**To re-run only the PPT** (after manually editing a research file or `content_map.json`):

1. Edit the markdown file or `content_map.json` in `content/`
2. Open Claude Code and run:

```
python ~/.claude/skills/sales-plan/sales_plan_ppt.py \
  --content-map "/path/to/deals/snowflake-2026/content/content_map.json" \
  --output "/path/to/deals/snowflake-2026/artifacts"
```

This skips all research and regenerates only the PPT from the existing JSON. Useful when you've manually refined the content.

---

## Best Practices

**1. Run it the day before a prep session, not the morning of a customer meeting.**
Research takes 3–6 minutes. Give yourself time to review the output, spot hallucinations, and fill in the manual slides before you're on a call.

**2. Always read the DECISION-MAKERS.md before using contact names.**
The research is thorough but fallible. People change roles. Verify names and titles against LinkedIn or Sales Nav before any customer-facing use.

**3. Internal briefs beat web research.**
If you have an `*_Exec_Brief*.pdf` for the account (from the account team or a prior deal), read it before running `/sales plan`. It contains confirmed stakeholder names, Amplify spend, and in-flight motions that web research cannot produce. Add those facts to `content_map.json` before re-running the PPT step.

**4. Never fabricate contact confirmation.**
If the research flags a contact as "MEDIUM confidence" or "requires verification," treat it as a hypothesis — not a confirmed stakeholder. The confidence level is there for a reason.

**5. The PPT is a starting point, not a final draft.**
The skill gives you 70–80% of a polished deck. Slides 2, 7, and 11 especially benefit from a human pass — tighten the language, add anything you know from your own calls, and remove anything that doesn't feel right for this specific account relationship.

**6. Save the `content_map.json`.**
It's a record of exactly what Claude extracted for each slide zone. If a slide looks off in the PPT, the JSON shows you whether the problem was in the research or the PPT population step.

**7. Use `--products` to anchor the plan.**
If you're in an active deal, pass `--products="ALM"` or whatever's in scope. It keeps the narrative focused and makes the portfolio plays on Slide 7 more precise.

---

## Troubleshooting

### "Template not found" when running the skill

Your template path is stale — the file may have moved when OneDrive re-synced, or you got a new machine.

**Fix:** Re-run the installer. It will ask for the new path and patch the skill.

```bash
./install.sh     (Mac / Git Bash)
.\install.ps1    (Windows PowerShell)
```

---

### "ModuleNotFoundError: No module named 'pptx'"

python-pptx isn't installed in the Python environment Claude Code is using.

**Fix:**
```bash
pip install python-pptx lxml        (Mac / Linux)
py -m pip install python-pptx lxml  (Windows)
```

---

### "/sales prospect skill not found" or similar

The AI Sales Team base skills aren't installed.

**Fix:**
```bash
curl -fsSL https://raw.githubusercontent.com/zubair-trabzada/ai-sales-team-claude/main/install.sh | bash
```

---

### Some slides are blank or show "[FILL: ...]" unexpectedly

Two possible causes:

**Cause 1 — Shape names changed in a newer template version.**
The script finds shapes by name. If you're on a different version of the template than the one the script was built for, names may differ.

Run this to inspect your template's shape names:
```python
from pptx import Presentation
prs = Presentation('/path/to/your/template.pptx')
for i, slide in enumerate(prs.slides):
    print(f'--- Slide {i+1} ---')
    for shape in slide.shapes:
        print(f'  {shape.name}')
```
Compare against the shape names in `~/.claude/skills/sales-plan/sales_plan_ppt.py` and raise an issue in the repo with the discrepancy.

**Cause 2 — Research agent failed for that skill.**
If one of the 5 parallel agents failed (timeout, URL issue, rate limit), the corresponding content file will be empty or contain an error. Check the file in `content/`. Re-run that specific skill:

```
/sales qualify https://www.snowflake.com
```

Then manually update `content_map.json` with the new content and re-run the PPT step.

---

### The company name in the PPT filename looks wrong (e.g., "Applied-Materials-Inc" instead of "Applied-Materials")

The filename is derived from the `company_name` field in `content_map.json`. Edit `content_map.json`, change `company_name` to your preferred format, and re-run the PPT step.

---

### Research looks outdated or wrong

The research agents pull from public web sources at the time of the run. If the account has had a major event (acquisition, leadership change, earnings) since the run, refresh:

```
/sales plan https://www.snowflake.com
```

This overwrites all research files with fresh data. Your prior PPT is preserved with a timestamp.

---

### I accidentally used the wrong URL

The deal folder name is derived from the company name found in research. If the wrong URL produced research for the wrong company, delete the folder and re-run with the correct URL.

---

## Quick Reference Card

### Install

```bash
# Step 1 — Install base AI Sales Team skills
curl -fsSL https://raw.githubusercontent.com/zubair-trabzada/ai-sales-team-claude/main/install.sh | bash

# Step 2 — Install /sales plan
git clone https://github.com/YOUR_REPO_URL/sales-plan-skill.git
cd sales-plan-skill && ./install.sh          # Mac / Git Bash
cd sales-plan-skill && .\install.ps1         # Windows PowerShell
```

### Run

```
# Basic
/sales plan https://www.company.com

# With deal context
/sales plan https://www.company.com --arr=150000 --renewal=Q3 --stage=POC --close-date=2026-08-31 --products="ALM"
```

### Output location

```
~/Desktop/claude-workspace/deals/<account>-<year>/artifacts/<Company>-Business-Plan-<date>.pptx
```

### Slide summary

| Slide | Title | Status |
|-------|-------|--------|
| 1 | Title (company name) | Auto |
| 2 | Executive Summary | Auto |
| 3 | Company Overview | Auto |
| 4 | Performance History | **Manual** — source: Panorama / Clari |
| 5 | Market Trends & Goals | Auto |
| 6 | Org Chart / Buying Committee | Auto |
| 7 | Value Strategy | Auto |
| 8 | FY Opportunities | **Manual** — source: Clari |
| 9 | FY Timeline | Auto |
| 11 | Big Idea (Appendix) | Auto |

### Re-run PPT only (after manual edits to content_map.json)

```bash
python ~/.claude/skills/sales-plan/sales_plan_ppt.py \
  --content-map "/path/to/deals/<account>/content/content_map.json" \
  --output "/path/to/deals/<account>/artifacts"
```

### Uninstall

```bash
./uninstall.sh     (Mac / Git Bash)
```

---

*Built for Adobe AEs. Questions or issues → raise them in the repo or ping the author.*
