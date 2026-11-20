---
name: pdf-to-ppt
description: >-
  Convert vector slide PDFs (exported from WPS Presentation, PowerPoint, Keynote,
  Beamer, Canva, etc.) into 100% native, object-level editable PowerPoint (.pptx)
  files with exact geometry, shadow-free native shapes, East Asian XML font mapping
  (<a:ea>), native hanging-indent bullets, programmatic text-parity checks, and
  PowerPoint COM visual diff QA. Use whenever the user asks to convert a PDF to PPT/PPTX,
  PDF转PPT, or reconstruct an editable slide deck from a PDF; route pure scanned or
  bitmap-only PDFs with no selectable text to nature-image2ppt.
---

# Vector PDF to Editable PPTX

Convert native vector slide PDFs into clean, 100% native editable PowerPoint presentations (no full-page raster screenshots, no broken line-by-line textboxes, no missing East Asian font bindings).

Retain the directory containing this file as `<skill_dir>` and resolve all relative paths under it. Keep all temporary scripts, exported verification PDFs, and comparison images inside `<appDataDir>/brain/<conversation-id>/scratch/` so the user's target directory remains clean with only the final `.pptx`.

## 5-Stage Workflow

### Stage 1 — Triage & Deep PDF Inspection

Inspect the source PDF structure, vector drawings, text blocks, embedded fonts, and locally installed Windows CJK fonts:

```bash
python "<skill_dir>/scripts/inspect_pdf.py" "<input.pdf>" --detail
```

- **Route decision**:
  - **Vector / hybrid slide PDF** (has selectable text and/or vector drawings): continue with Stage 2 of this skill.
  - **Pure scanned / bitmap-only PDF** (`text_len == 0` across pages with full-page raster images): switch to `nature-image2ppt` instead.
- Check the reported PDF fonts against installed Windows fonts and note any special geometry (rounded-rectangle corner radii, arrow shapes, overlapping/occluded shapes, inline color highlights).

### Stage 2 — Native Vector & Semantic Text Reconstruction

Run the generic automatic converter ([convert_vector_pdf.py](./scripts/convert_vector_pdf.py)), which builds on [pptx_helpers.py](./scripts/pptx_helpers.py):

```bash
python "<skill_dir>/scripts/convert_vector_pdf.py" "<input.pdf>" "<output.pptx>"
```

The converter handles:
1. **1:1 slide canvas**: Sets `prs.slide_width = Pt(w)` and `prs.slide_height = Pt(h)` directly from page 1 dimensions (zero coordinate scaling).
2. **Native slide background**: Converts full-page background rectangles (`0, 0, w, h`) into `slide.background.fill.solid()`.
3. **WPS outline deduplication & occlusion removal**: Filters out `0.5pt`-expanded outline-emulation polygons (`ops.get('l', 0) >= 50`) and completely occluded identical-bbox shapes.
4. **Shadow-free native shapes**: Reconstructs `ROUNDED_RECTANGLE` (with measured `adj = radius / min(w, h)`), `OVAL`, and `RIGHT_ARROW` (with exact shaft/head proportions), stripping PowerPoint's default theme drop shadow (`<a:effectRef idx="0"/>`, `shadow.inherit = False`, empty `<a:effectLst/>`).
5. **East Asian XML font binding (`<a:ea>`)**: Maps PDF font subset names to installed Windows Latin + CJK fonts, writes `<a:latin>`, `<a:ea>`, and `<a:cs>` explicitly on every run, and normalizes `±0.05pt` PDF export rounding (`round(sz * 2) / 2`).
6. **Semantic text grouping**:
   - Places short badge digits inside enclosing `OVAL` text frames (`PP_ALIGN.CENTER`, `MSO_ANCHOR.MIDDLE`).
   - Places single-line banners/pills inside enclosing empty `ROUNDED_RECTANGLE` shapes (`MSO_ANCHOR.MIDDLE`).
   - Groups consecutive hanging-indent bullets (`•`) in the same column into one native `<a:buChar char="•"/>` list (`marL`, `indent`, `<a:buClr>`).
   - Joins wrapped multi-line blocks into editable paragraphs using exact point line spacing `Pt(pitch_pt)` (never proportional float `1.2`) and font-specific `y_offset` calibration.

If a deck has unusual custom diagrams, multi-column merged blocks, or layout edge cases that need slide-specific overrides, write a dedicated conversion script in `scratch/` that imports [pptx_helpers.py](./scripts/pptx_helpers.py) rather than duplicating helper logic. Consult [vector-pitfalls-and-calibration.md](./references/vector-pitfalls-and-calibration.md) when tuning `y_offset`, `pitch_pt`, or font mappings.

### Stage 3 — Programmatic Fidelity Gate

Verify text completeness, slide dimensions, and `<a:ea>` font coverage on every run:

```bash
python "<skill_dir>/scripts/verify_pptx.py" "<input.pdf>" "<output.pptx>"
```

All slides must report `PASS` (`Counter` normalized character equality + every PDF text line contained verbatim) and `missing <a:ea>: 0`.

### Stage 4 — PowerPoint Visual Diff QA & Calibration

Export the generated `.pptx` via locally installed Microsoft PowerPoint (do not skip visual QA when PowerPoint COM is available) and compare every slide against the original PDF:

1. **Export verification PDF into `scratch/`** (try sandboxed first; if COM activation fails with `80070005 Access is denied`, rerun only this command with `BypassSandbox: true` and the identical `toolSummary`):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "<skill_dir>/scripts/export_pdf_powerpoint.ps1" -PptxPath "<output.pptx>" -OutPdfPath "<scratch_dir>/verify_export.pdf"
```

2. **Compute per-slide pixel diff & render side-by-side contact sheets**:

```bash
python "<skill_dir>/scripts/render_compare.py" "<input.pdf>" "<scratch_dir>/verify_export.pdf" "<scratch_dir>"
```

3. **Inspect all generated `compare_*.png` sheets** with `read_file`. Check for unintended line wraps, vertical drift, clipped cards, missing font fallbacks, or lingering shadows. Consult [vector-pitfalls-and-calibration.md](./references/vector-pitfalls-and-calibration.md) to adjust `y_offset`, `w_override`, or `pitch_pt`, then re-run Stages 2–4 until every slide aligns cleanly.

### Stage 5 — Workspace Hygiene & Delivery

Verify the user's target folder contains only the original file(s) and the final `.pptx` (no temporary PDFs, scripts, or PNGs). Deliver a clickable `file:///` link to the `.pptx` along with a concise summary of the verified slide count, text parity, and visual QA results.
