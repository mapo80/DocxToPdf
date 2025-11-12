# Phase 7 – Tables

Spec ready for implementation and CI.

---

## 1. Parsing & normalization (Docx layer)

- [ ] Extend `DocxDocument` to enumerate blocks in-order (`Paragraph`, `Table`).
- [ ] Build `DocxTable` model that captures:
  - `tblW` (value + type `auto|dxa|pct`)
  - `tblLayout` (`fixed|autofit`, default `autofit`)
  - `tblBorders`, `tblCellMar`
  - For each `w:tr` → rows; for each `w:tc` → cells with:
    - `tcW` (value + type), `gridSpan`, `vMerge`, `vAlign`, `tcBorders`, `tcMar`
    - Paragraph content (reuse existing paragraph parser)
- [ ] Normalize spans:
  - Horizontal: `gridSpan` → single owner cell with `columnSpan`
  - Vertical: handle `vMerge restart/continue`; attach “owner” reference so renderer knows rowSpan
- [ ] Validation: warn (diagnostics logger) if gridSpan/vMerge references exceed grid width or missing owner.

## 2. Width resolution

- [ ] Implement width solver (`TableWidthResolver`):
  - Inputs: `tblLayout`, `tblW`, column definitions, preferred cell widths.
  - `fixed`: lock to grid widths; if `tblW` present (dxa/pct) scale columns proportionally.
  - `autofit`: use `tblW` if present; otherwise autosize to content (measure paragraphs) but clamp deterministically.
  - Support `pct`: Word stores 1/50 percent; convert to pt vs. table text area.
- [ ] Produce final column widths in pt, stored alongside the table model.

## 3. Rendering (DocxToPdfConverter)

- [ ] Iterate blocks; when encountering tables:
  - Layout cell rectangles using resolved column widths and row heights (sum of tallest cells per row, including rowSpan).
  - Apply padding: start from `tblCellMar`, override with `tcMar`.
  - Support `vAlign`: top/center/bottom inside the content box.
- [ ] Spans:
  - Horizontal: merged cell draws once, width = sum columns; skip placeholders.
  - Vertical: owner spans multiple rows; height = sum row heights; skip placeholders.
- [ ] Borders:
  - Compute effective border per edge (table default → cell override).
  - Draw contiguous lines once per grid edge to avoid double stroke; consistent thickness (pt).
- [ ] Pagination: if row does not fit remaining height, push to next page before drawing.

## 4. Tests

### Unit tests (DocxToPdf.Tests)

- [ ] `DocxTableParserTests`
  - `tblW` auto/dxa/pct
  - `tblLayout` fixed/autofit
  - `tcW` overrides, `gridSpan`, `vMerge`, `tcMar`, `vAlign`
- [ ] `TableWidthResolverTests`
  - fixed grid scaling
  - pct widths vs. text area
  - autosize fallback determinism
- [ ] `TableBorderLogicTests`
  - precedence table vs. cell, shared edges
- [ ] `TablePaddingTests`
  - default + override combination
- [ ] `TableRowSpanTests`
  - vertical merge continuity and height accumulation

Target coverage ≥ 90 % (line + branch) for new table-related classes.

### Integration tests

- [ ] `LayoutStressTests.ShouldMatchBaseline`
  - Render DOCX → PDF, rasterize both via PDFium at fixed DPI, diff pixel-perfect (threshold 0).
  - Use PdfPig to confirm text counts (no missing words).
- [ ] `TabsAlignmentTests.ShouldRemainPixelPerfect`
  - Regression guard to ensure table code doesn’t regress other samples.

### Tooling

- [ ] Add PDFium raster diff helper (CLI or test fixture) + instructions.
- [ ] Update CI to fail if coverage < 90 %.

## 5. Acceptance checklist

- [ ] Tables respect Word width semantics (tblLayout/tblW/tcW).
- [ ] gridSpan/vMerge produce single visual cells with correct borders.
- [ ] Borders render cleanly (no gaps/double strokes).
- [ ] Padding + vAlign behave per tblCellMar/tcMar/vAlign.
- [ ] Integration tests green; coverage ≥ 90 %.
