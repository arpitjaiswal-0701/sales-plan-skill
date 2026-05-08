#!/usr/bin/env python3
"""sales_plan_ppt.py — Populate Adobe Lite DX Business Plan Template from content_map.json.

Reads a ``content_map.json`` produced by the ``/sales plan`` Claude Code skill
and populates a cloned copy of the Adobe Lite DX Business Plan PowerPoint
template. All intelligence work happens upstream in Claude; this script is
purely mechanical: read JSON, locate shapes by name, rebuild text frames.

Usage::

    python sales_plan_ppt.py \\
        --content-map  "/path/to/deals/company/content/content_map.json" \\
        --template     "/path/to/Lite DX Business Plan Template.pptx" \\
        --output       "/path/to/deals/company/artifacts"

Dependencies:
    python-pptx >= 0.6.23
    lxml >= 4.0

Slides populated (shape-name lookup unless noted):
    1   Company name title
    2   Executive summary table — business issue, big idea, objectives,
        challenges, Adobe differentiated solution
    3   Company overview — background, LOB, account intel, Adobe strengths,
        opportunities
    4   Performance history — upcoming renewals, renewal strategy
        (from CLI args or Clari/Panorama; [FILL] placeholder otherwise)
    5   Market landscape — trends, goals, digital priorities, opportunities,
        challenges, partner strategy, implementation partners
    6   Buying committee table — org chart shapes cleared first; fresh table
        inserted with Adobe red header row (up to 10 contacts)
    7   Value strategy — big idea, tagline, business issue, portfolio plays,
        path to value
    8   FY opportunities — pipeline summary
        (from CLI args or Clari; [FILL] placeholder otherwise)
    9   FY timeline — H1 (Dec–Jul) and H2 (Aug–Nov) touchpoints (table rows)
    11  Big Idea appendix — goals, challenges, initiatives, impact rows,
        full Big Idea paragraph, tagline

Notes:
    - The source template is never modified; the script always clones it.
    - If an output file already exists it is renamed with a ``_HHMM`` suffix
      before the new file is written, preserving prior versions.
    - Shape names are matched exactly (case-sensitive). Run the shape
      inspector snippet in README.md if slides appear blank after population.

Version: 1.1.0
Changelog: CHANGELOG.md
"""

import argparse
import copy
import json
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

# Windows console may be cp1252 — reconfigure to utf-8 so Unicode output works
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from lxml import etree


__version__ = "1.1.0"

TEMPLATE_DEFAULT = "__TEMPLATE_PATH__"


# ── Shape helpers ─────────────────────────────────────────────────────────────

def find_shape(container: Any, name: str) -> Optional[Any]:
    """Recursively search a slide or group for a shape with the given name.

    Args:
        container: A python-pptx Slide or GroupShapes object with a ``.shapes``
            collection. Any other type is treated as having no shapes.
        name: The exact shape name to match (case-sensitive).

    Returns:
        The first matching shape, or ``None`` if no match is found.
    """
    shapes = getattr(container, 'shapes', [])
    for shape in shapes:
        if shape.name == name:
            return shape
        if shape.shape_type == 6:  # MSO_SHAPE_TYPE.GROUP
            result = find_shape(shape, name)
            if result:
                return result
    return None


def set_text(shape: Any, text: str) -> None:
    """Populate a shape's text frame, replacing all existing content.

    Extracts the first run's ``<a:rPr>`` (font size, bold, color, typeface)
    before clearing the text frame, then rebuilds one ``<a:p>`` per
    newline-delimited line. This guarantees no residual template text or
    field elements survive in the XML.

    Newline splitting enables multi-line bullet content from
    ``content_map.json`` (e.g. ``"• Trend 1\\n• Trend 2"``) to render as
    separate paragraphs rather than a single run-together line.

    Calls :func:`normalize_body_props` after writing to fix vertical anchor
    and autofit settings on the shape.

    Args:
        shape: A python-pptx shape with a text frame. Silently no-ops if
            the shape is falsy or does not have a text frame.
        text: Content to write. ``\\n`` characters produce separate
            ``<a:p>`` elements; an empty string produces one blank paragraph.
    """
    if not shape or not shape.has_text_frame:
        return
    txBody = shape.text_frame._txBody
    paras = txBody.findall(qn('a:p'))
    if not paras:
        return

    # Extract rPr from first run before clearing — captures font, size, color,
    # theme references, and any custom run properties set by the template.
    rPr: Optional[Any] = None
    runs = paras[0].findall(qn('a:r'))
    if runs:
        rPr_elem = runs[0].find(qn('a:rPr'))
        if rPr_elem is not None:
            rPr = copy.deepcopy(rPr_elem)

    for para in paras:
        txBody.remove(para)

    for line in (text.split('\n') if text else ['']):
        para = etree.SubElement(txBody, qn('a:p'))
        pPr = etree.SubElement(para, qn('a:pPr'))
        pPr.set('algn', 'l')
        r = etree.SubElement(para, qn('a:r'))
        if rPr is not None:
            r.insert(0, copy.deepcopy(rPr))
        t = etree.SubElement(r, qn('a:t'))
        t.text = line

    normalize_body_props(shape)


def normalize_body_props(shape: Any) -> None:
    """Fix text frame body properties that cause alignment and clipping defects.

    Sets three properties on ``<a:bodyPr>`` that are frequently misconfigured
    in PowerPoint templates and produce three distinct visual failures:

    - ``anchor="t"`` — aligns text to the top of the shape. Without this,
      text can appear at the middle or bottom of a partially-filled box.
    - ``<a:normAutofit>`` — shrinks the font proportionally to fit the shape's
      fixed dimensions, replacing ``<a:noAutofit>`` (clips overflow) and
      ``<a:spAutoFit>`` (resizes the shape). Preserves template geometry.
    - Removes ``anchorCtr`` — eliminates secondary vertical centering that
      can re-introduce misalignment even after ``anchor`` is corrected.

    Internal margins (``lIns``, ``rIns``, ``tIns``, ``bIns``) are intentionally
    left untouched; they are set by the template designer per shape.

    Args:
        shape: A python-pptx shape with a text frame. Silently no-ops if the
            shape is falsy, lacks a text frame, or has no ``<a:bodyPr>``.
    """
    if not shape or not shape.has_text_frame:
        return
    txBody = shape.text_frame._txBody
    bodyPr = txBody.find(qn('a:bodyPr'))
    if bodyPr is None:
        return

    bodyPr.set('anchor', 't')
    bodyPr.attrib.pop('anchorCtr', None)

    for tag in (qn('a:noAutofit'), qn('a:spAutoFit')):
        for el in bodyPr.findall(tag):
            bodyPr.remove(el)
    if not bodyPr.findall(qn('a:normAutofit')):
        etree.SubElement(bodyPr, qn('a:normAutofit'))


def set_cell_text(cell: Any, text: str) -> None:
    """Write text into a table cell, clearing all existing paragraphs first.

    Removes extra paragraphs that carry residual template content before
    delegating to the python-pptx high-level API for the first paragraph.
    Silently suppresses exceptions so a single bad cell does not abort the
    enclosing slide population.

    Args:
        cell: A python-pptx table cell object.
        text: The text string to write into the cell.
    """
    try:
        tf = cell.text_frame
        txBody = tf._txBody
        paras = txBody.findall(qn('a:p'))
        for para in paras[1:]:
            txBody.remove(para)
        if tf.paragraphs:
            tf.paragraphs[0].text = text
    except Exception:
        pass


def get_tables(slide: Any) -> List[Any]:
    """Return all table shapes on a slide in document order.

    Args:
        slide: A python-pptx Slide object.

    Returns:
        List of shapes where ``shape_type == 19`` (``MSO_SHAPE_TYPE.TABLE``).
        Returns an empty list if the slide has no tables.
    """
    return [s for s in slide.shapes if s.shape_type == 19]


# ── Slide populators ──────────────────────────────────────────────────────────

def slide_1(slide: Any, cm: Dict[str, Any]) -> None:
    """Populate Slide 1 — company name in the title shape."""
    shape = find_shape(slide, 'Title 1')
    if shape:
        set_text(shape, cm.get('company_name', '[FILL: Company Name]'))


def slide_2(slide: Any, cm: Dict[str, Any]) -> None:
    """Populate Slide 2 — Executive Summary table and last-refreshed date.

    The executive summary table has 8 columns; content lands in row 1 at
    columns 0, 1, 3, 5, and 7. Columns 2, 4, and 6 are merged continuations
    of their preceding cell and must be skipped to avoid XML corruption.
    """
    s = cm.get('slide_2', {})
    tables = get_tables(slide)
    if tables:
        t = tables[0].table
        mapping: Dict[int, str] = {
            0: s.get('business_issue',    '[FILL: Business Issue]'),
            1: s.get('big_idea',           '[FILL: Big Idea]'),
            3: s.get('company_objectives', '[FILL: Company Objectives]'),
            5: s.get('challenges',         '[FILL: Key Challenges]'),
            7: s.get('adobe_solution',     '[FILL: Adobe Differentiated Solution]'),
        }
        for col, text in mapping.items():
            try:
                set_cell_text(t.cell(1, col), text)
            except Exception:
                pass

    date_shape = find_shape(slide, 'TextBox 15')
    if date_shape:
        set_text(date_shape, f"Last Refreshed on: {cm.get('date', datetime.today().strftime('%Y-%m-%d'))}")


def slide_3(slide: Any, cm: Dict[str, Any]) -> None:
    """Populate Slide 3 — Company Overview content boxes."""
    s = cm.get('slide_3', {})
    targets: Dict[str, str] = {
        'Rectangle 11': s.get('account_background',    '[FILL: Revenue · Employees · Business Focus · Industry · Position]'),
        'Rectangle 32': s.get('account_background_lob','[FILL: LOB/Department account overview]'),
        'Rectangle 34': s.get('account_intel',          '[FILL: Account intelligence — whitespace rationale, Tier 1 reasoning]'),
        'Rectangle 51': s.get('adobe_strengths',        '[FILL: Adobe advantages and capabilities for this account]'),
        'Rectangle 59': s.get('opportunities',          '[FILL: Specific opportunity areas to create value]'),
    }
    for name, text in targets.items():
        shape = find_shape(slide, name)
        if shape:
            set_text(shape, text)


def slide_4(slide: Any, cm: Dict[str, Any]) -> None:
    """Populate Slide 4 — Performance History (renewals and renewal strategy).

    Populated from CLI arguments (``--arr``, ``--renewal``) if provided by the
    ``/sales plan`` skill; otherwise a ``[FILL]`` placeholder is written.
    Complete population requires Clari/Panorama data.

    Template table layout: index 0 = instructions, 1 = Upcoming Renewals,
    2 = Renewal Strategy.
    """
    s = cm.get('slide_4', {})
    tables = get_tables(slide)
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


def slide_5(slide: Any, cm: Dict[str, Any]) -> None:
    """Populate Slide 5 — Market Landscape (trends, goals, priorities, partners)."""
    s = cm.get('slide_5', {})
    targets: Dict[str, str] = {
        'Rectangle 49':  s.get('market_trends',          '[FILL: Broad industry/market developments — 3-4 bullets]'),
        'Rectangle 9':   s.get('company_goals',           '[FILL: Company strategic goals — 3-4 bullets]'),
        'Rectangle 15':  s.get('digital_priorities',      '[FILL: Digital priorities and technology strategy]'),
        'Rectangle 14':  s.get('market_opportunities',    '[FILL: Specific areas to create value or solve problems]'),
        'Rectangle 4':   s.get('customer_challenges',     '[FILL: Key obstacles customer must overcome — in customer language]'),
        'Rectangle 183': s.get('partner_strategy',        '[FILL: Partner strategy and relationships]'),
        'Rectangle 187': s.get('implementation_partners', '[FILL: Historical implementation partners]'),
    }
    for name, text in targets.items():
        shape = find_shape(slide, name)
        if shape:
            set_text(shape, text)


def slide_6(slide: Any, cm: Dict[str, Any]) -> None:
    """Populate Slide 6 — Buying Committee table.

    Removes all non-placeholder rectangle and text-box shapes from the slide
    before inserting the new table. This clears the template's pre-built org
    chart diagram (individual ``AUTO_SHAPE`` rectangles) and associated legend
    text boxes, which cannot coexist cleanly with the programmatically inserted
    table.

    Placeholder shapes (e.g. the slide title) are preserved via the
    ``is_placeholder`` guard. If org chart shapes happen to be grouped
    (``shape_type == 6``), add ``6`` to the type filter.

    Table layout:
        - Columns: Name (2.2 in) | Title (3.0 in) | Role (2.8 in) | Attitude (4.5 in)
        - Header row: Adobe red background (``#FA0F00``), white bold text
        - Data rows: up to 10 contacts from ``content_map.json``
          ``slide_6.buying_committee``
    """
    # Remove org chart shapes and legend before adding the table.
    # Shape types: 1 = AUTO_SHAPE (rectangles), 17 = TEXT_BOX (legend labels)
    to_remove = [
        s for s in slide.shapes
        if not s.is_placeholder and s.shape_type in (1, 17)
    ]
    for s in to_remove:
        s._element.getparent().remove(s._element)

    contacts = cm.get('slide_6', {}).get('buying_committee', [])
    if not contacts:
        return

    rows = min(len(contacts), 10) + 1  # +1 for header row
    tbl_shape = slide.shapes.add_table(
        rows, 4, Inches(0.3), Inches(1.15), Inches(12.5), Pt(20) * rows
    )
    tbl = tbl_shape.table
    tbl.columns[0].width = Inches(2.2)
    tbl.columns[1].width = Inches(3.0)
    tbl.columns[2].width = Inches(2.8)
    tbl.columns[3].width = Inches(4.5)

    def styled_cell(cell: Any, text: str, bold: bool = False, header: bool = False) -> None:
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
            tcPr = cell._tc.get_or_add_tcPr()
            srgbClr = etree.SubElement(
                etree.SubElement(tcPr, qn('a:solidFill')), qn('a:srgbClr')
            )
            srgbClr.set('val', 'FA0F00')  # Adobe red

    for ci, header_text in enumerate(['Name', 'Title', 'Role', 'Attitude toward Adobe']):
        styled_cell(tbl.cell(0, ci), header_text, bold=True, header=True)

    for ri, contact in enumerate(contacts[:10]):
        styled_cell(tbl.cell(ri + 1, 0), str(contact.get('name',     '')))
        styled_cell(tbl.cell(ri + 1, 1), str(contact.get('title',    '')))
        styled_cell(tbl.cell(ri + 1, 2), str(contact.get('role',     '')))
        styled_cell(tbl.cell(ri + 1, 3), str(contact.get('attitude', '')))


def slide_7(slide: Any, cm: Dict[str, Any]) -> None:
    """Populate Slide 7 — Value Strategy (big idea, business issue, portfolio plays)."""
    s = cm.get('slide_7', {})
    targets: Dict[str, str] = {
        'Rectangle 25':  s.get('big_idea',       '[FILL: Big Idea in customer language — 2-3 sentences]'),
        'Rectangle 26':  s.get('tagline',          '[FILL: One-sentence tagline]'),
        'Rectangle 69':  s.get('business_issue',   '[FILL: Business issue impacting performance and goals]'),
        'Rectangle 84':  s.get('portfolio_plays',  '[FILL: Adobe Portfolio / Sales Plays to pitch]'),
        'Rectangle 217': s.get('path_to_value',    '[FILL: Path to Value and Budget alignment. Potential pipeline: $X]'),
    }
    for name, text in targets.items():
        shape = find_shape(slide, name)
        if shape:
            set_text(shape, text)


def slide_8(slide: Any, cm: Dict[str, Any]) -> None:
    """Populate Slide 8 — FY Opportunities summary.

    Populated from CLI arguments (``--arr``, ``--stage``, ``--close-date``) if
    provided; otherwise a ``[FILL]`` placeholder directs the user to Clari.
    """
    s = cm.get('slide_8', {})
    tables = get_tables(slide)
    if tables:
        try:
            set_cell_text(tables[0].table.cell(0, 0),
                          s.get('opportunities_summary',
                                '[FILL: Complete in Clari. Paste pipeline table: Opp Name | Solution | ARR | Stage | Close Date | Next Step.]'))
        except Exception:
            pass


def slide_9(slide: Any, cm: Dict[str, Any]) -> None:
    """Populate Slide 9 — FY Timeline touchpoints.

    Two tables cover the fiscal year: index 0 spans Dec–Jul (H1),
    index 1 spans Aug–Nov (H2). Content lands in row 1, column 1 of each.
    """
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


def slide_11(slide: Any, cm: Dict[str, Any]) -> None:
    """Populate Slide 11 — Big Idea appendix.

    Writes four labeled rows (Goals, Challenges, Initiatives, Impact) into the
    first table on the slide, then populates the full Big Idea paragraph text
    box and a standalone tagline text box.
    """
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

POPULATORS: Dict[int, Any] = {
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


# ── Entry point ───────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description='Populate Lite DX Business Plan PPT from content_map.json'
    )
    parser.add_argument('--content-map', required=True, help='Path to content_map.json')
    parser.add_argument('--template', default=TEMPLATE_DEFAULT, help='Path to .pptx template')
    parser.add_argument('--output', required=True, help='Output directory path')
    args = parser.parse_args()

    with open(args.content_map, 'r', encoding='utf-8') as f:
        cm = json.load(f)

    template_path = Path(args.template)
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
        print(f"  ↳ Prior version renamed: {backup.name}")

    shutil.copy2(str(template_path), str(out_path))
    prs = Presentation(str(out_path))

    errors: List[str] = []
    for i, slide in enumerate(prs.slides):
        n = i + 1
        if n in POPULATORS:
            try:
                POPULATORS[n](slide, cm)
            except Exception as e:
                errors.append(f"Slide {n}: {e}")

    prs.save(str(out_path))

    print(f"\n✓ PPT saved: {out_path}")
    if errors:
        print("\n⚠ Non-fatal errors:")
        for e in errors:
            print(f"  {e}")

    fill_needed: List[int] = []
    for slide_key, slide_num in [('slide_4', 4), ('slide_8', 8)]:
        vals = cm.get(slide_key, {}).values()
        if any(str(v).startswith('[FILL') for v in vals):
            fill_needed.append(slide_num)
    if fill_needed:
        print(f"\n⚑ Manual fill needed on slides: {', '.join(str(s) for s in fill_needed)}")
        print("  Source: Clari / Panorama")


if __name__ == '__main__':
    main()
