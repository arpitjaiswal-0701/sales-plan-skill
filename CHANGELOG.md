# Changelog

All notable changes to `tools/sales_plan_ppt.py` are documented here.

Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).
Versioning follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

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

[Unreleased]: https://github.com/arpitjaiswal-0701/sales-plan-skill/compare/v1.1.0...HEAD
[1.1.0]: https://github.com/arpitjaiswal-0701/sales-plan-skill/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/arpitjaiswal-0701/sales-plan-skill/releases/tag/v1.0.0
