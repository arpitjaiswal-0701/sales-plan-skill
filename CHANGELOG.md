# Changelog

All notable changes to `tools/sales_plan_ppt.py` are documented here.

Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).
Versioning follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

---

## [1.2.0] — 2026-05-13

### Fixed

- **Slide 6 org chart still visible** (`slide_6`): The v1.1.0 pre-clear only
  removed `AUTO_SHAPE` (type 1) and `TEXT_BOX` (type 17) shapes, leaving all
  `GROUP` (type 6), `PICTURE` (type 13), and `LINE` (type 9) shapes — which
  comprise the bulk of the org chart diagram, connector lines, photo ovals,
  and legend elements. The pre-clear now removes **all** non-placeholder
  shapes, completely clearing the slide before table insertion.

- **Slide 2 table cells not populating** (`set_cell_text`): The prior
  high-level `tf.paragraphs[0].text = text` API silently fails on merge-anchor
  cells and cells with protected run structure. Rewrote using the same
  XML-level txBody rebuild as `set_text`: all `<a:p>` elements removed and
  rebuilt from scratch with preserved `<a:rPr>` formatting.

- **Table rows expanding and distorting slide** (`set_cell_text`): Table cells
  lacked `<a:normAutofit>` on their `<a:bodyPr>`, allowing rows to auto-expand
  when content exceeded the design height. `set_cell_text` now applies the
  same `normalize_body_props` logic (top anchor, normAutofit) to each table
  cell's text body.

- **Slide 7 content missing** (`slide_7`): `Rectangle 63` (Like
  Customers/Customer Reference, nested two levels deep in Group 218 → Group
  60) was not included in the target map. Added with the new `like_customers`
  field from `content_map.json`.

- **Slide 9 content overflowing single cell** (`slide_9`): All H1 and H2
  touchpoints were written to a single cell (December / August), causing
  overflow and leaving 7 of 8 month columns blank. Rewrote to distribute
  content across individual month columns using the new `h1`/`h2` nested
  object schema. Legacy flat `touchpoints_h1`/`touchpoints_h2` strings remain
  supported for backward compatibility.

### Added

- **Slide 1 DALP title format** (`slide_1`): Title placeholder now reads
  "{Company Name} — DALP Account Plan" instead of the bare company name.

- **Slide 1 company logo** (`slide_1`, `fetch_logo`): New `fetch_logo(domain,
  output_dir)` function downloads the company logo from the Clearbit Logo API
  (`logo.clearbit.com/{domain}`) and inserts it into the lower-right corner of
  Slide 1 (10.43", 5.60", 2.5" × 1.5"). Cached to `logo_cache.png` in the
  output directory on first run. Silently no-ops on network failure — the
  DALP title is still set.

- **Slide 6 Reason for Engagement column**: Buying committee table expanded
  from 4 columns to 5: Name (2.0"), Title (2.5"), Role (1.8"), **Reason for
  Engagement** (4.2"), Attitude toward Adobe (2.0"). Requires `reason` key in
  each `buying_committee` object in `content_map.json`.

- **`ctx` parameter**: All slide populator functions now accept an optional
  `ctx: Dict` argument containing `output_dir` and any future per-run context.
  The `main()` function constructs `ctx` and passes it to every populator,
  replacing the prior zero-argument dispatch.

- `import urllib.request` for logo download (stdlib, no new dependency).

### Changed

- `SKILL.md` schema updated:
  - Added `company_domain` top-level field (used by `fetch_logo`).
  - `slide_3.account_background` and `slide_3.account_background_lob` now
    specify the exact 10-line labeled format matching the template's structured
    placeholder text (Annual Revenue, Online Revenue, Size, etc.).
  - `slide_6.buying_committee` objects now include a `reason` field.
  - `slide_7` adds `like_customers` field (maps to `Rectangle 63`).
  - `slide_9` schema changed from flat `touchpoints_h1`/`touchpoints_h2`
    strings to nested `h1`/`h2` objects with monthly keys (`december` through
    `july` for H1; `august` through `november` for H2).
  - Source mapping table updated to reflect all new fields.

---

## [1.1.0] — 2026-05-08

### Fixed

- **Template text bleed-through** (`set_text`): The prior implementation
  surgically removed `<a:r>` run elements but left other XML children
  (`<a:fld>`, `<a:br>`, field elements) that carried the original placeholder
  text. The function now performs a full txBody rebuild — all `<a:p>` elements
  are removed before new content is written, guaranteeing a clean slate on
  every write.

- **Wrong vertical text position** (`normalize_body_props`): Template shapes
  frequently carry `anchor="ctr"` or `anchor="b"`, causing generated content
  to appear at the middle or bottom of a box rather than the top. A new
  `normalize_body_props` helper enforces `anchor="t"` on every shape write.

- **Text clipping** (`normalize_body_props`): Template shapes with
  `<a:noAutofit>` silently clip text that exceeds the shape height. Replaced
  with `<a:normAutofit>` (shrink-to-fit), which keeps shape geometry and
  template layout intact while ensuring all content remains visible.

- **Wrong horizontal alignment** (`set_text`): Each rebuilt `<a:pPr>` now
  explicitly sets `algn="l"`, overriding any conflicting alignment instruction
  inherited from the template shape.

- **Slide 6 org chart not removed** (`slide_6`): The buying committee table
  was added on top of the template's pre-built org chart diagram (individual
  rectangle shapes) and associated legend text boxes, which remained visible
  and overlaid the new table. The function now removes all non-placeholder
  `AUTO_SHAPE` (type 1) and `TEXT_BOX` (type 17) shapes before calling
  `add_table()`.

- **Residual table cell paragraphs** (`set_cell_text`): Setting
  `tf.paragraphs[0].text` replaced only the first paragraph; additional
  template paragraphs in a cell were left intact. Extra paragraphs are now
  removed via lxml before writing.

### Added

- `normalize_body_props(shape)` — standalone helper for body-property
  normalization; callable independently of `set_text` if needed.
- `__version__ = "1.1.0"` module-level constant.
- `import copy` for safe deep-copying of extracted `<a:rPr>` elements.
- Type hints (`Optional`, `List`, `Dict`, `Any` from `typing`) on all
  function signatures.
- Google-style docstrings on all public functions, including `Args`,
  `Returns`, and design rationale for non-obvious behaviour.
- Module-level docstring with usage, dependency table, per-slide content
  inventory, and notes on template-safety and backup behaviour.

### Changed

- `set_text`: Full txBody rebuild replaces the prior partial paragraph/run
  removal approach. Newline characters in the `text` argument now produce
  separate `<a:p>` elements, enabling bullet content from `content_map.json`
  (e.g. `"• Item 1\n• Item 2"`) to render as distinct lines.

- `slide_6`: Pre-clear step added before `add_table()`. The intermediate
  `cols` and `height` local variables are inlined into the `add_table()`
  call (no behaviour change).

- `POPULATORS` dict and `errors` / `fill_needed` variables annotated with
  explicit types.

---

## [1.0.0] — 2026-04-01

### Added

- Initial release: shape-name-based population of slides 1–9, 11 from
  `content_map.json`.
- Slide 6 buying committee table with Adobe red (`#FA0F00`) header row and
  four columns: Name, Title, Role, Attitude toward Adobe.
- Timestamp-based backup of prior output file (e.g.
  `Company-Business-Plan-2026-05-08_1430.pptx`) to prevent silent overwrite
  on refresh runs.
- UTF-8 stdout reconfiguration for Windows cp1252 console environments.
- CLI arguments `--arr`, `--renewal`, `--stage`, `--close-date`, `--products`
  forwarded through the `/sales plan` skill for pre-filling Slides 4 and 8.
- `find_shape` recursive group traversal so shapes nested inside
  `MSO_SHAPE_TYPE.GROUP` containers are reachable by name.

---

[Unreleased]: https://github.com/arpitjaiswal-0701/sales-plan-skill/compare/v1.2.0...HEAD
[1.2.0]: https://github.com/arpitjaiswal-0701/sales-plan-skill/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/arpitjaiswal-0701/sales-plan-skill/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/arpitjaiswal-0701/sales-plan-skill/releases/tag/v1.0.0
