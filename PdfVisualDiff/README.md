# PdfVisualDiff

`PdfVisualDiff` is a .NET 9 CLI that rasterises two PDFs with Poppler `pdftoppm`, compares the resulting PNGs pixel-by-pixel with ImageMagick, and emits both numeric metrics and visual artefacts suitable for CI/CD sign-offs.

## Requirements

- .NET SDK 9.0+
- [Poppler `pdftoppm`](https://manpages.debian.org/pdftoppm) (available via `brew install poppler`, `apt install poppler-utils`, etc.)
- [ImageMagick](https://imagemagick.org) with the `compare`, `identify`, and `magick` (or `convert`) executables on `PATH`

If any tool resides in a custom location you can point the CLI to it via `--pdftoppm-path`, `--compare-path`, `--identify-path`, or `--magick-path`.

## Build

```bash
dotnet build PdfVisualDiff.sln
```

## CLI usage

```
pdf-visual-diff <baseline.pdf> <candidate.pdf> --out <output-dir> [options]
```

Key options:

| Option | Description |
| --- | --- |
| `--dpi <int>` | Rasterisation DPI passed to `pdftoppm` (`-r`). Default: `300`. |
| `--page-range 1-3,5` | Only compare the listed pages. Pages outside either PDF are reported as skipped. |
| `--cropbox` | Adds `-cropbox` so both PDFs use the CropBox when rasterised (helps align margins). |
| `--fuzz 0.1%` | Forwards the fuzz tolerance to `compare` when computing metrics/diff images. |
| `--baseline-user-password`, `--baseline-owner-password` | Credentials forwarded to `pdftoppm -upw/-opw` for the baseline PDF. Candidate equivalents exist. |
| `--thumbnail-size 320` | Maximum edge for thumbnail PNGs embedded in the HTML report. |
| `--geometry-check --geometry-tolerance 0.25` | (Optional) run the PdfPig-based geometry analyser after the pixel diff. |
| `--ghostscript-path` | Explicit Ghostscript binary path (required for the Ghostscript/ImageMagick renderer). |
| `--pdfbox-dir` | Directory containing `pdfbox-app-*.jar` and the `geometry-extractor.jar` helper (default: `tools/pdfbox`). |
| `--warning-ae-percent 0.005`, `--warning-ssim 0.995` | Tune the warning band for AE%/SSIM before promoting differences to FAIL. |
| `--ssim-radius`, `--ssim-sigma`, `--ssim-k1`, `--ssim-k2` | Forwarded as `-define compare:ssim-*` to ImageMagick when computing SSIM/DSSIM. |
| `--pdftoppm-path`, `--compare-path`, `--identify-path`, `--magick-path` | Explicit tool locations. |

Per i sample inclusi nel repository, i PDF esportati da Word e forniti come riferimento
sono sotto `samples/golden-word/*.pdf` (baseline/A); i PDF prodotti dal nostro
convertitore (es. `out/<nome>.pdf`) sono i candidate/B.
Quando le rasterizzazioni hanno bounding box che differiscono di pochi pixel,
`pdf-visual-diff` estende automaticamente la pagina più piccola con margini bianchi
per allineare l’area d’immagine (niente resampling o scaling).

Example:

```bash
dotnet run --project PdfVisualDiff/src/PdfVisualDiff \
  samples/baseline.pdf samples/candidate.pdf \
  --out artifacts/invoice-001 \
  --dpi 600 \
  --page-range 1-5 \
  --cropbox \
  --fuzz 0.02% \
  --geometry-check
```

### Workflow example (samples/lorem.docx)

1. Genera il PDF candidato dal DOCX: `dotnet run --project DocxToPdf.Demo samples/lorem.docx out/lorem-generated.pdf`.
2. Confronta il PDF generato con il riferimento Word incluso nei sample:\
   `dotnet run --project PdfVisualDiff/src/PdfVisualDiff samples/golden-word/lorem.pdf out/lorem-generated.pdf --out out/diff-lorem --dpi 300`.
3. I risultati sono disponibili nella directory indicata con `--out` (in questo esempio `out/diff-lorem`). All'interno trovi:
   - `poppler/` e `ghostscript/` con sottocartelle `A/`, `B/`, `diff/`, `overlay/`, `thumbs/` per i due motori di rendering.
   - `report.csv` con le metriche per pagina (apribile con Excel/LibreOffice o analizzabile in CI).
   - `index.html` con il riepilogo grafico dual-render, badge Consistent/Engine-only, filtri e pannello PDFBox con overlay interattive su miniature.
   - `report.json` con summary + pagine + versioni tool + risultato PDFBox (comodo per CI/dashboard).

## Output structure

The output directory contains a deterministic, inspectable dataset:

- `poppler/…` e `ghostscript/…` – PNG rasterisations (`A/`, `B/`), diff (`diff/`), overlay (`overlay/`) e thumbnails per ciascun motore.
- `report.csv` – una riga per renderer con colonne `page,renderer,classification,width,height,dpi,ae_pixels,ae_percent,mae,rmse,psnr,ssim,dssim,status,path_img_*`.
- `report.json` – payload strutturato (summary, pagine, classificazioni, versioni tool/soglie e analisi PDFBox) per automazioni o dashboard.
- `index.html` – dashboard dual-render con badge “Consistent / Engine-only”, filtri, miniature Poppler/Ghostscript e (facoltativo) pannello geometry.

Metrics collected for every page:

- **Dual render**: per ogni pagina ottieni due set completi di metriche (Poppler/pdftoppm e Ghostscript/ImageMagick) per distinguere differenze reali da artefatti engine-specifici.
- **AE** (Absolute Error) → pixel diversi (AE%) = AE / (w × h) × 100.
- **MAE**, **RMSE** → errori medi (assoluto/quadratico) normalizzati 0–1.
- **PSNR** → peak signal-to-noise ratio in dB (`inf` quando i raster sono identici).
- **SSIM/DSSIM** → similarità strutturale (SSIM ↑ meglio, DSSIM ↓ meglio). Puoi passare parametri dedicated (`--ssim-*`) per allineare il comportamento agli script ImageMagick più esigenti.

Status rules:

- **PASS**: `AE = 0`, `PSNR = inf`, `SSIM ≈ 1` (⇒ `DSSIM ≈ 0`) su entrambi i renderer.
- **WARNING**: `AE% ≤ soglia` **oppure** `SSIM ≥ soglia`.
- **FAIL**: oltre le soglie oppure geometry check KO (soglie configurabili via CLI).
- Il report etichetta ogni pagina come **Consistent** (differenze visibili con entrambi i motori), **Engine-only** (solo uno dei due) o **Clean**; usa i filtri dell’HTML per concentrarti sulle differenze “reali”.

Exit codes: `0` for pass/warn, `1` for fail, `2` for unexpected errors/tool failures.

## Testing

```bash
dotnet test PdfVisualDiff.sln
```

Unit tests cover page-range parsing and geometry helpers; the CLI tests exercise the overall pipeline (skip automatically if Poppler/ImageMagick are missing).

## CI tips

1. Install Poppler and ImageMagick in the agent image (e.g., `apt install poppler-utils imagemagick`).
2. Cache `~/.nuget/packages` between runs for faster restores.
3. Publish `report.csv`, `report.json`, `index.html`, and le directory `poppler/` + `ghostscript/` come artefatti per analisi forensi.
4. Run the CLI with `--geometry-check` in pipelines where text layout regressions matter.
