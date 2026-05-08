#!/usr/bin/env python3
"""
sales_plan_ppt.py — Populate Adobe Lite DX Business Plan Template from content_map.json

Usage:
    python sales_plan_ppt.py \
        --content-map  "/path/to/deals/company/content/content_map.json" \
        --template     "/path/to/Lite DX Business Plan Template.pptx" \
        --output       "/path/to/deals/company/artifacts"
"""

import argparse
import json
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

# Windows console may be cp1252 — reconfigure to utf-8 so Unicode output works
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from lxml import etree


TEMPLATE_DEFAULT = "__TEMPLATE_PATH__"


# ── Shape helpers ─────────────────────────────────────────────────────────────

def find_shape(container, name):
    """Recursively find a shape by name in a slide or group."""
    shapes = getattr(container, 'shapes', [])
    for shape in shapes:
        if shape.name == name:
            return shape
        if shape.shape_type == 6:  # GROUP
            result = find_shape(shape, name)
            if result:
                return result
    return None


def set_text(shape, text):
    """Replace text in a shape while preserving first-run font formatting."""
    if not shape or not shape.has_text_frame:
        return
    txBody = shape.text_frame._txBody
    paras = txBody.findall(qn('a:p'))
    if not paras:
        return

    # Drop all paragraphs after the first
    for para in paras[1:]:
        txBody.remove(para)

    first_para = paras[0]
    runs = first_para.findall(qn('a:r'))

    if runs:
        for run in runs[1:]:
            first_para.remove(run)
        t = runs[0].find(qn('a:t'))
        if t is None:
            t = etree.SubElement(runs[0], qn('a:t'))
        t.text = text
    else:
        r = etree.SubElement(first_para, qn('a:r'))
        t = etree.SubElement(r, qn('a:t'))
        t.text = text


def set_cell_text(cell, text):
    """Set text in a table cell's text frame."""
    try:
        tf = cell.text_frame
        if tf.paragraphs:
            tf.paragraphs[0].text = text
        else:
            set_text(tf, text)
    except Exception:
        pass


def get_tables(slide):
    """Return all TABLE shapes on a slide in document order."""
    return [s for s in slide.shapes if s.shape_type == 19]


# ── Slide populators ──────────────────────────────────────────────────────────

def slide_1(slide, cm):
    """Title — company name."""
    shape = find_shape(slide, 'Title 1')
    if shape:
        set_text(shape, cm.get('company_name', '[FILL: Company Name]'))


def slide_2(slide, cm):
    """Executive Summary table + last-refreshed date."""
    s = cm.get('slide_2', {})
    tables = get_tables(slide)
    if tables:
        t = tables[0].table
        # Exec summary table has 8 cols; content is in row 1 at cols 0,1,3,5,7
        # (cols 2,4,6 are merged continuations — skip them)
        mapping = {
            0: s.get('business_issue',      '[FILL: Business Issue]'),
            1: s.get('big_idea',             '[FILL: Big Idea]'),
            3: s.get('company_objectives',   '[FILL: Company Objectives]'),
            5: s.get('challenges',           '[FILL: Key Challenges]'),
            7: s.get('adobe_solution',       '[FILL: Adobe Differentiated Solution]'),
        }
        for col, text in mapping.items():
            try:
                set_cell_text(t.cell(1, col), text)
            except Exception:
                pass

    date_shape = find_shape(slide, 'TextBox 15')
    if date_shape:
        set_text(date_shape, f"Last Refreshed on: {cm.get('date', datetime.today().strftime('%Y-%m-%d'))}")


def slide_3(slide, cm):
    """Company Overview — account background boxes."""
    s = cm.get('slide_3', {})
    targets = {
        'Rectangle 11': s.get('account_background',   '[FILL: Revenue · Employees · Business Focus · Industry · Position]'),
        'Rectangle 32': s.get('account_background_lob','[FILL: LOB/Department account overview]'),
        'Rectangle 34': s.get('account_intel',         '[FILL: Account intelligence — whitespace rationale, Tier 1 reasoning]'),
        'Rectangle 51': s.get('adobe_strengths',       '[FILL: Adobe advantages and capabilities for this account]'),
        'Rectangle 59': s.get('opportunities',         '[FILL: Specific opportunity areas to create value]'),
    }
    for name, text in targets.items():
        shape = find_shape(slide, name)
        if shape:
            set_text(shape, text)


def slide_4(slide, cm):
    """Performance History — renewal data (CLI args or placeholders)."""
    s = cm.get('slide_4', {})
    tables = get_tables(slide)
    # Template has 3 tables: [0] instructions, [1] Upcoming Renewals, [2] Renewal Strategy
    if len(tables) > 1:
        try:
            set_cell_text(tables[1].table.cell(1, 0),
                          s.get('upcoming_renewals',
                                '[FILL: Upcoming renewals — Solution | $Value | Renewal Date | Driver. Source: Clari/Panorama.]'))
        except Exception:
            pass
    if len(tables) > 2:
        try:
            set_cell_text(tables[2].table.cell(1, 0),
                          s.get('renewal_strategy',
                                '[FILL: Renewal strategy and risk. Source: Panorama / Clari.]'))
        except Exception:
            pass


def slide_5(slide, cm):
    """Market Trends, Goals, Priorities, Partners."""
    s = cm.get('slide_5', {})
    targets = {
        'Rectangle 49':  s.get('market_trends',       '[FILL: Broad industry/market developments — 3-4 bullets]'),
        'Rectangle 9':   s.get('company_goals',        '[FILL: Company strategic goals — 3-4 bullets]'),
        'Rectangle 15':  s.get('digital_priorities',   '[FILL: Digital priorities and technology strategy]'),
        'Rectangle 14':  s.get('market_opportunities', '[FILL: Specific areas to create value or solve problems]'),
        'Rectangle 4':   s.get('customer_challenges',  '[FILL: Key obstacles customer must overcome — in customer language]'),
        'Rectangle 183': s.get('partner_strategy',     '[FILL: Partner strategy and relationships]'),
        'Rectangle 187': s.get('implementation_partners', '[FILL: Historical implementation partners]'),
    }
    for name, text in targets.items():
        shape = find_shape(slide, name)
        if shape:
            set_text(shape, text)


def slide_6(slide, cm):
    """Org Chart — insert buying committee as a clean table."""
    contacts = cm.get('slide_6', {}).get('buying_committee', [])
    if not contacts:
        return

    rows = min(len(contacts), 10) + 1  # +1 for header row
    cols = 4
    left   = Inches(0.3)
    top    = Inches(1.15)
    width  = Inches(12.5)
    height = Pt(20) * rows

    tbl_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    tbl = tbl_shape.table

    tbl.columns[0].width = Inches(2.2)
    tbl.columns[1].width = Inches(3.0)
    tbl.columns[2].width = Inches(2.8)
    tbl.columns[3].width = Inches(4.5)

    def styled_cell(cell, text, bold=False, header=False):
        tf = cell.text_frame
        tf.word_wrap = True
        para = tf.paragraphs[0]
        for r in para.runs:
            r._r.getparent().remove(r._r)
        run = para.add_run()
        run.text = text
        run.font.size = Pt(8.5)
        run.font.bold = bold
        if header:
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            solidFill = etree.SubElement(tcPr, qn('a:solidFill'))
            srgbClr = etree.SubElement(solidFill, qn('a:srgbClr'))
            srgbClr.set('val', 'FA0F00')  # Adobe red

    headers = ['Name', 'Title', 'Role', 'Attitude toward Adobe']
    for ci, h in enumerate(headers):
        styled_cell(tbl.cell(0, ci), h, bold=True, header=True)

    for ri, contact in enumerate(contacts[:10]):
        styled_cell(tbl.cell(ri + 1, 0), str(contact.get('name',     '')))
        styled_cell(tbl.cell(ri + 1, 1), str(contact.get('title',    '')))
        styled_cell(tbl.cell(ri + 1, 2), str(contact.get('role',     '')))
        styled_cell(tbl.cell(ri + 1, 3), str(contact.get('attitude', '')))


def slide_7(slide, cm):
    """Value Strategy — big idea, business issue, portfolio plays."""
    s = cm.get('slide_7', {})
    targets = {
        'Rectangle 25':  s.get('big_idea',        '[FILL: Big Idea in customer language — 2-3 sentences]'),
        'Rectangle 26':  s.get('tagline',          '[FILL: One-sentence tagline]'),
        'Rectangle 69':  s.get('business_issue',   '[FILL: Business issue impacting performance and goals]'),
        'Rectangle 84':  s.get('portfolio_plays',  '[FILL: Adobe Portfolio / Sales Plays to pitch]'),
        'Rectangle 217': s.get('path_to_value',    '[FILL: Path to Value and Budget alignment. Potential pipeline: $X]'),
    }
    for name, text in targets.items():
        shape = find_shape(slide, name)
        if shape:
            set_text(shape, text)


def slide_8(slide, cm):
    """FY25 Opportunities — Clari data or placeholder."""
    s = cm.get('slide_8', {})
    tables = get_tables(slide)
    if tables:
        try:
            set_cell_text(tables[0].table.cell(0, 0),
                          s.get('opportunities_summary',
                                '[FILL: Complete in Clari. Paste pipeline table: Opp Name | Solution | ARR | Stage | Close Date | Next Step.]'))
        except Exception:
            pass


def slide_9(slide, cm):
    """FY25 Timeline — touchpoints rows."""
    s = cm.get('slide_9', {})
    tables = get_tables(slide)
    if tables:
        try:
            set_cell_text(tables[0].table.cell(1, 1),
                          s.get('touchpoints_h1', '[FILL: H1 touchpoints — events, renewal meetings, CEC]'))
        except Exception:
            pass
    if len(tables) > 1:
        try:
            set_cell_text(tables[1].table.cell(1, 1),
                          s.get('touchpoints_h2', '[FILL: H2 touchpoints — MAX, renewal close, EBC]'))
        except Exception:
            pass


def slide_11(slide, cm):
    """Big Idea appendix — Goals/Challenges/Initiatives/Impact + paragraph."""
    s = cm.get('slide_11', {})
    tables = get_tables(slide)
    if tables:
        t = tables[0].table
        rows_data = [
            ('goals',       'Goals\n'),
            ('challenges',  'Challenges\n'),
            ('initiatives', 'Initiatives\n'),
            ('impact',      'Impact\n'),
        ]
        for ri, (key, prefix) in enumerate(rows_data):
            val = s.get(key, f'[FILL: {key.title()} — from company goals/priorities slide]')
            try:
                set_cell_text(t.cell(ri, 0), prefix + val)
            except Exception:
                pass

    big_idea = find_shape(slide, 'TextBox 25')
    if big_idea:
        set_text(big_idea, s.get('big_idea_paragraph', '[FILL: Big Idea paragraph in customer language]'))

    tagline = find_shape(slide, 'TextBox 11')
    if tagline:
        set_text(tagline, s.get('tagline', '[FILL: Tagline]'))


# ── Slide dispatch ─────────────────────────────────────────────────────────────

POPULATORS = {
    1:  slide_1,
    2:  slide_2,
    3:  slide_3,
    4:  slide_4,
    5:  slide_5,
    6:  slide_6,
    7:  slide_7,
    8:  slide_8,
    9:  slide_9,
    11: slide_11,
}


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Populate Lite DX Business Plan PPT from content_map.json')
    parser.add_argument('--content-map', required=True, help='Path to content_map.json')
    parser.add_argument('--template', default=TEMPLATE_DEFAULT, help='Path to .pptx template')
    parser.add_argument('--output', required=True, help='Output directory')
    args = parser.parse_args()

    with open(args.content_map, 'r', encoding='utf-8') as f:
        cm = json.load(f)

    template_path = Path(args.template)
    if not template_path.exists():
        print(f"ERROR: Template not found: {template_path}")
        print("Check the path or re-run the installer to update it.")
        sys.exit(1)

    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    company  = re.sub(r'[^\w\s-]', '', cm.get('company_name', 'Unknown')).strip().replace(' ', '-')
    date_str = cm.get('date', datetime.today().strftime('%Y-%m-%d'))
    out_name = f"{company}-Business-Plan-{date_str}.pptx"
    out_path = output_dir / out_name

    if out_path.exists():
        ts = datetime.now().strftime('%H%M')
        backup = output_dir / f"{company}-Business-Plan-{date_str}_{ts}.pptx"
        shutil.move(str(out_path), str(backup))
        print(f"  Prior version renamed: {backup.name}")

    shutil.copy2(str(template_path), str(out_path))
    prs = Presentation(str(out_path))

    errors = []
    for i, slide in enumerate(prs.slides):
        n = i + 1
        if n in POPULATORS:
            try:
                POPULATORS[n](slide, cm)
            except Exception as e:
                errors.append(f"Slide {n}: {e}")

    prs.save(str(out_path))

    print(f"\nPPT saved: {out_path}")
    if errors:
        print("\nNon-fatal errors:")
        for e in errors:
            print(f"  {e}")

    fill_needed = []
    for slide_key, slide_num in [('slide_4', 4), ('slide_8', 8)]:
        vals = cm.get(slide_key, {}).values()
        if any(str(v).startswith('[FILL') for v in vals):
            fill_needed.append(slide_num)
    if fill_needed:
        print(f"\nManual fill needed on slides: {', '.join(str(s) for s in fill_needed)}")
        print("  Source: Clari / Panorama")


if __name__ == '__main__':
    main()
