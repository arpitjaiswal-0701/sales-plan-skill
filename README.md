# /sales plan — Adobe Lite DX Business Plan Generator

A Claude Code skill that runs all `/sales` research workflows for a target account and automatically populates your Adobe **Lite DX Business Plan** PowerPoint template — formatting intact, slides ready to present.

One command. 6 research agents. A populated PPT in your deal folder.

---

## What It Does

```
/sales plan https://www.snowflake.com --arr=150000 --renewal=Q3 --stage=POC
```

1. Runs `/sales prospect` to establish company profile and deal folder
2. Runs 5 research agents in parallel: qualify, contacts, competitors, prep, outreach
3. Compresses all research into per-slide character budgets
4. Clones your PPT template and populates 8 slides by shape name
5. Drops the finished deck in `deals/<account>-<year>/artifacts/`

```
deals/snowflake-2026/
├── PROSPECT-ANALYSIS.md
├── content/
│   ├── LEAD-QUALIFICATION.md
│   ├── DECISION-MAKERS.md
│   ├── COMPETITIVE-INTEL.md
│   ├── MEETING-PREP.md
│   ├── OUTREACH-SEQUENCE.md
│   └── content_map.json        ← audit trail of what Claude extracted
└── artifacts/
    └── Snowflake-Business-Plan-2026-05-07.pptx
```

### Slides auto-populated

| Slide | Content | Source |
|-------|---------|--------|
| 1 | Company name | PROSPECT-ANALYSIS |
| 2 | Executive summary (business issue, big idea, objectives, challenges, Adobe solution) | All sources |
| 3 | Company overview (background, LOB, account intel, Adobe strengths, opportunities) | PROSPECT-ANALYSIS + LEAD-QUALIFICATION |
| 5 | Market trends, company goals, digital priorities, partners | PROSPECT-ANALYSIS + MEETING-PREP |
| 6 | Buying committee table (name, title, role, attitude) | DECISION-MAKERS |
| 7 | Value strategy (big idea, tagline, business issue, portfolio plays, path to value) | MEETING-PREP + COMPETITIVE-INTEL |
| 9 | FY timeline — H1 and H2 touchpoints | MEETING-PREP + LEAD-QUALIFICATION |
| 11 | Big Idea appendix — goals, challenges, initiatives, impact, full paragraph | All sources |

### Slides requiring manual fill

| Slide | Content | Source |
|-------|---------|--------|
| 4 | Performance History — renewal history, ARR trends | Panorama / Clari |
| 8 | FY Opportunities — open pipeline | Clari |

Pass `--arr`, `--renewal`, `--stage`, `--close-date` to pre-fill Slides 4 and 8 with what you know.

---

## Prerequisites

| Requirement | Version | Notes |
|-------------|---------|-------|
| [Claude Code](https://claude.ai/download) | Latest | The CLI that runs the skill |
| [AI Sales Team skills](https://github.com/zubair-trabzada/ai-sales-team-claude) | Latest | Provides the 6 sub-skills this orchestrates |
| Python | 3.8+ | For PPT generation |
| python-pptx | ≥0.6.23 | Auto-installed by the installer |
| Adobe Lite DX Business Plan Template | Your copy | Synced from OneDrive; installer asks for the path |

> **The template is never modified.** The script always clones it — your master copy stays clean.

---

## Installation

### Step 1 — Install the AI Sales Team (if you haven't already)

```bash
curl -fsSL https://raw.githubusercontent.com/zubair-trabzada/ai-sales-team-claude/main/install.sh | bash
```

### Step 2 — Install this skill

**Mac / Linux / Git Bash on Windows:**
```bash
git clone https://github.com/arpitjaiswal-0701/sales-plan-skill.git
cd sales-plan-skill
chmod +x install.sh
./install.sh
```

**Windows PowerShell:**
```powershell
git clone https://github.com/arpitjaiswal-0701/sales-plan-skill.git
cd sales-plan-skill
.\install.ps1
```

The installer will ask you for:
1. **Template path** — the full path to your `Lite DX Business Plan Template - <Name>.pptx` file (it tries to detect it from OneDrive automatically)
2. **Deals root** — where your account folders live (defaults to `~/Desktop/claude-workspace/deals`)

It patches both the skill and the Python script with your paths, so no manual editing is required.

### Optional: install Python dependencies manually

```bash
pip install -r requirements.txt
```

---

## Usage

### Basic — full research + PPT

```bash
/sales plan https://www.snowflake.com
```

### With deal context — pre-fills Slides 4 and 8

```bash
/sales plan https://www.snowflake.com \
  --arr=150000 \
  --renewal=Q3 \
  --stage=POC \
  --close-date=2026-08-31 \
  --products="ALM,Marketo"
```

### CLI argument reference

| Argument | Description | Example |
|----------|-------------|---------|
| `<url>` | Company website (required) | `https://www.snowflake.com` |
| `--arr=<amount>` | Current or target ARR in dollars | `--arr=150000` |
| `--renewal=<quarter>` | Renewal quarter | `--renewal=Q3` or `--renewal="Q3 FY26"` |
| `--stage=<stage>` | Deal stage | `--stage=POC` |
| `--close-date=<date>` | Target close date | `--close-date=2026-08-31` |
| `--products=<list>` | Adobe products in scope | `--products="ALM,Marketo"` |

---

## Refresh Runs

Running `/sales plan` on an account that already has a deal folder is safe:

- Research files in `content/` are overwritten with fresh data
- The existing PPT in `artifacts/` is **renamed with a timestamp** (e.g., `Snowflake-Business-Plan-2026-05-07_1430.pptx`) before the new one is written — you never lose a prior version
- `deal.yaml` and `brief.md` in the folder root are **never overwritten**

To re-run only the PPT step (after manually editing a markdown file):
1. Edit the relevant file in `content/`
2. Update `content/content_map.json` with the revised content
3. Run the script directly:
```bash
python ~/.claude/skills/sales-plan/sales_plan_ppt.py \
  --content-map "/path/to/deals/company/content/content_map.json" \
  --output "/path/to/deals/company/artifacts"
```

---

## Troubleshooting

**`Template not found` error**
The template path set during installation no longer resolves — OneDrive may have moved it. Re-run the installer to update the path.

**`ModuleNotFoundError: No module named 'pptx'`**
```bash
pip install python-pptx lxml
```

**`/sales prospect` skill not found**
Install the AI Sales Team first (see Prerequisites above).

**Shape names not matching — some slides blank**
Your template version may have different shape names. Run the shape inspector to check:
```bash
python -c "
from pptx import Presentation
prs = Presentation('/path/to/your/template.pptx')
for i, slide in enumerate(prs.slides):
    print(f'--- Slide {i+1} ---')
    for shape in slide.shapes:
        print(f'  {shape.shape_type:2d}  {shape.name}')
"
```
Compare the output against the shape names in `tools/sales_plan_ppt.py` and update as needed.

**Windows console encoding error (`charmap codec`)**
Already handled — the script reconfigures stdout to UTF-8 on import. If you see this on an older Python, upgrade to 3.8+.

---

## How It Works

```
/sales plan <url>
       │
       ▼
 /sales prospect (blocking)
 → Establishes company name and deal folder path
       │
       ▼
 5 parallel agents
 → qualify · contacts · competitors · prep · outreach
 → all write to deals/<account>/content/
       │
       ▼
 Claude compresses each section to per-slide character budgets
 → writes content_map.json (no API key needed in Python)
       │
       ▼
 sales_plan_ppt.py
 → clones template
 → populates shapes by name (XML-level, preserves font/size/color)
 → slide 6: inserts clean table (buying committee)
 → saves to artifacts/
```

The Python script requires no API key — all intelligence work happens in the Claude skill step. The script is purely mechanical: read JSON, find shape by name, replace text.

---

## Uninstall

```bash
./uninstall.sh
```

Your deal folders and generated PPTs are untouched.

---

## License

MIT
