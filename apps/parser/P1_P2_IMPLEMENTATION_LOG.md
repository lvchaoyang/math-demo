# P1 + P2 Implementation Log

Date: 2026-03-23
Scope: Formula asset model (P1) + renderer skeleton (P2) + P3 initial OMML rendering

## Goals

- P1: Add a unified `formula_assets` data model to parser results.
- P2: Add a `FormulaRenderer` skeleton to produce an auditable render plan.
- P3 (initial): Attempt OMML/LaTeX pre-render to PNG when local toolchain is available.
- P4 (minimal): Attempt WMF preview pre-render to PNG via converter.
- Keep existing question splitting behavior unchanged.

## Files Changed

### 1) `app/core/parser.py`

- Added `FormulaRenderer` import.
- Added `hashlib` import.
- In `parse_docx(...)`, added:
  - `result["formula_assets"]`
  - `result["formula_render_plan"]`
  - metadata field `total_formula_assets`
- Added helper function `_build_formula_assets(...)`:
  - Scans each paragraph `content_items`
  - Extracts formula-like assets into a unified list
  - Supports source types:
    - `omml` (`latex`, `latex_block`)
    - `wmf_preview` (`image` with `.wmf/.emf`)
    - `mathtype_ole` (image metadata type)
  - Generates stable-ish asset IDs and writes paragraph-local refs:
    - `paragraph["formula_asset_refs"] = [...]`

### 2) `app/core/formula_renderer.py` (new)

- Added `FormulaRenderer` skeleton class:
  - `build_cache_key(asset)`
  - `build_render_plan(formula_assets, file_id=None)`
- Current behavior:
  - Produces actionable plan entries with:
    - `asset_id`, `source_type`, `action`, `cache_key`, `target_path`, `status`
  - P3 initial rendering path:
    - Detects local LaTeX engine (`xelatex/pdflatex/lualatex`)
    - Detects PDF-to-image tool (`pdftoppm`/`magick`)
    - For `source_type == omml`, attempts `LaTeX -> PDF -> PNG`
    - On failure, degrades to `source_only` with reason in `note`

### 4) `parse_docx` integration update

- `formula_render_plan` generation now passes:
  - `output_dir = image_output_dir/formula_assets` (when available)
  - `render_omml = True`
- This keeps rendering artifacts colocated with document image outputs.

### 3) `app/core/__init__.py`

- Exported `FormulaRenderer` in module exports.

## Output Contract Additions

`parse_docx(...)` now additionally returns:

- `formula_assets: List[Dict]`
- `formula_render_plan: List[Dict]`

And metadata adds:

- `total_formula_assets: int`

## Compatibility Notes

- Existing fields used by splitter/front-end are untouched.
- Existing API behavior remains compatible.
- New fields are additive only.
- If render toolchain is missing, no exception is raised; plan marks OMML assets as `source_only`.

## Verification

- Python compile check:
  - `app/core/parser.py`
  - `app/core/formula_renderer.py`
  - `app/core/__init__.py`
- No linter errors reported for changed files.

## Follow-up Integration (P3 wiring)

### 5) `app/core/splitter.py`

- `QuestionSplitter` now accepts `formula_render_plan`.
- Added rendered asset lookup map (`asset_id -> rendered_image`).
- In `_paragraph_to_html(...)`:
  - For `latex` / `latex_block` items, if a rendered asset exists, prefer `<img>` output.
  - Otherwise keep existing MathJax HTML output.
- `split_questions(...)` signature updated to accept optional `formula_render_plan`.
- Fixed block formula delimiter output to standard `$$...$$`.

### 6) `main.py`

- Updated all `split_questions(...)` call sites to pass:
  - `formula_render_plan=result.get("formula_render_plan")`
- This enables parser-rendered OMML images to flow into question HTML without changing API shape.

### 7) Render observability

- `parser.py` now returns:
  - `formula_render_summary` with keys:
    - `total`, `rendered`, `source_only`, `planned`, `skip`
    - `by_source_type` (e.g. `omml`, `wmf_preview`, `mathtype_ole`)
    - `by_action` (e.g. `render_from_wmf_preview`)
    - `by_note` (failure reason distribution)
- `/parse` and `/parse/v2` question responses now include:
  - `formula_render_summary`
- This makes it easier to validate how many formulas were actually rendered as images.

## P4 Minimal Implementation

### 8) `app/core/formula_renderer.py`

- Added WMF preview rendering path:
  - `source_type == wmf_preview` -> `render_from_wmf_preview`
  - Uses `WMFConverter` to produce PNG into formula asset output directory
  - On success marks plan status `rendered`; on failure marks `source_only`

### 9) `app/core/parser.py`

- `build_render_plan(...)` now passes `source_image_dir=image_output_dir`
- Enables renderer to locate extracted WMF source files for conversion.

## P4.1 WMF Squeeze Mitigation

### 10) `app/core/formula_renderer.py`

- Added post-conversion normalization for WMF preview render outputs:
  - After `WMFConverter.convert(...)` succeeds, renderer runs safe normalization
  - If content width utilization is low, attempts:
    1) fill-first (`resize WxH! + extent`)
    2) automatic fallback to contain (`resize WxH + extent`) when fill fails
- Key behavior:
  - Fill is preferred to reduce left-side squeeze/unused right area
  - Any normalization failure is silent and non-blocking (keeps pipeline stable)
  - Existing output file is preserved unless a temp normalized file is successfully generated

## P4.2 MathType OLE Priority

### 11) `app/core/parser.py`

- In formula asset extraction:
  - If a paragraph contains `mathtype_ole` image items, `wmf_preview` items in the same paragraph are suppressed.
  - Purpose: avoid using low-quality preview WMF when OLE semantic source exists.

### 12) `app/core/formula_renderer.py`

- Added `mathtype_ole` execution path:
  - action: `use_mathtype_png`
  - If extracted MathType PNG exists in source image dir, mark as `rendered` and reuse it directly.
  - Otherwise mark `source_only` with explicit note.

## P4.3 WMF De-stretch Safety

### 13) `app/core/formula_renderer.py`

- Changed WMF post-normalization policy from fill-first to contain-only:
  - Removed forced resize with `!` (non-proportional scaling).
  - Kept trim + proportional resize + centered extent only.
- Goal:
  - Prevent glyph deformation and overlap caused by aggressive stretch.
  - Prioritize formula shape fidelity over canvas fill ratio.
