using System.Globalization;
using System.Text;
using System.Text.Json;

namespace PdfVisualDiff.Core;

internal static class ReportBuilder
{
    private static readonly JsonSerializerOptions OverlayJsonOptions = new(JsonSerializerDefaults.Web)
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = false
    };

    private static readonly IReadOnlyDictionary<string, string> WordDescriptions = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        ["Missing"] = "Presente solo nel baseline (sparito nel candidato).",
        ["Extra"] = "Presente solo nel candidato (baseline invariato).",
        ["Moved"] = "Stessa parola ma spostata oltre la soglia di tolleranza."
    };

    private static readonly IReadOnlyDictionary<string, string> GraphicDescriptions = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        ["Added"] = "Elemento vettoriale aggiunto nel candidato.",
        ["Removed"] = "Elemento presente solo nel baseline.",
        ["Changed"] = "Elemento presente in entrambi ma spostato/ridimensionato."
    };

    public static void WriteCsv(string destination, IReadOnlyList<PageComparisonResult> pages, string root)
    {
        var sb = new StringBuilder();
        sb.AppendLine("page,renderer,classification,width,height,dpi,ae_pixels,ae_percent,mae,rmse,psnr,ssim,dssim,status,path_img_A,path_img_B,path_img_diff,path_img_overlay");

        foreach (var page in pages.OrderBy(p => p.PageNumber))
        {
            foreach (var renderer in page.Renderers)
            {
                sb.Append(page.PageNumber).Append(',')
                    .Append(renderer.Engine).Append(',')
                    .Append(page.Classification).Append(',')
                    .Append(renderer.Width).Append(',')
                    .Append(renderer.Height).Append(',')
                    .Append(renderer.Dpi).Append(',')
                    .Append(renderer.AePixels).Append(',')
                    .Append(renderer.AePercent.ToString("0.##########", CultureInfo.InvariantCulture)).Append(',')
                    .Append(FormatDouble(renderer.Mae)).Append(',')
                    .Append(FormatDouble(renderer.Rmse)).Append(',')
                    .Append(FormatDouble(renderer.Psnr)).Append(',')
                    .Append(FormatDouble(renderer.Ssim)).Append(',')
                    .Append(FormatDouble(renderer.Dssim)).Append(',')
                    .Append(renderer.Status).Append(',')
                    .Append(RelativePath(root, renderer.BaselineImagePath)).Append(',')
                    .Append(RelativePath(root, renderer.CandidateImagePath)).Append(',')
                    .Append(RelativePath(root, renderer.DiffImagePath)).Append(',')
                    .Append(RelativePath(root, renderer.OverlayImagePath)).AppendLine();
            }
        }

        File.WriteAllText(destination, sb.ToString());
    }

    public static void WriteHtml(
        string destination,
        VisualDiffSummary summary,
        IReadOnlyList<PageComparisonResult> pages,
        GeometryAnalysisResult? geometry,
        string outputRoot)
    {
        var sb = new StringBuilder();
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang=\"en\"><head><meta charset=\"utf-8\"/>");
        sb.AppendLine("<title>PDF Visual Diff</title>");
        sb.AppendLine("""
<style>
body{font-family:system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;margin:2rem;background:#f7f7f9;color:#111}
h1{margin-top:0}
table{border-collapse:collapse;width:100%;margin-top:1.25rem;font-size:0.95rem}
th,td{border:1px solid #ddd;padding:0.35rem 0.5rem;text-align:left;vertical-align:top}
th{background:#fafafa}
.status-Pass{color:#0a6a0a;font-weight:600}
.status-Warning{color:#b58900;font-weight:600}
.status-Fail{color:#b00020;font-weight:600}
.thumbs{display:flex;gap:0.5rem}
.thumbs img{border:1px solid #ccc;background:#fff;max-width:160px;height:auto}
.summary-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:0.5rem}
.summary-card{background:#fff;border:1px solid #e0e0e0;border-radius:6px;padding:0.75rem;box-shadow:0 1px 2px rgba(0,0,0,.05)}
.meta{margin:0;list-style:none;padding:0}
.meta li{margin:0.2rem 0}
.badge{display:inline-block;padding:0.1rem 0.4rem;border-radius:12px;font-size:0.8rem;background:#ececec}
.badge.status-Pass{background:#e6f4ea;color:#1e4620}
.badge.status-Warning{background:#fff4db;color:#8a6d00}
.badge.status-Fail{background:#fde7e9;color:#8a051f}
details summary{cursor:pointer;font-weight:600}
.legend{margin-top:2rem;font-size:0.9rem;color:#333}
.classification-badge{display:inline-block;padding:0.2rem 0.6rem;border-radius:999px;font-size:0.8rem;font-weight:600}
.classification-Clean{background:#e6f4ea;color:#1e4620}
.classification-Consistent{background:#fde7e9;color:#8a051f}
.classification-EngineOnly{background:#fff4db;color:#8a6d00}
.filters{margin:1rem 0;display:flex;gap:1rem;flex-wrap:wrap;font-size:0.9rem}
.geometry-section{margin-top:2rem}
.geometry-panel{border:1px solid #e0e0e0;border-radius:6px;padding:1rem;margin-top:1rem;background:#fff;box-shadow:0 1px 2px rgba(0,0,0,.04)}
.geom-controls{display:flex;gap:0.75rem;flex-wrap:wrap;font-size:0.85rem;margin-bottom:0.5rem}
.geom-controls label{display:flex;align-items:center;gap:0.3rem}
.geom-preview{position:relative;display:inline-block}
.geom-preview img{display:block;max-width:320px;height:auto;border:1px solid #ddd;border-radius:4px}
.geom-preview canvas{position:absolute;left:0;top:0;pointer-events:none}
.geom-tables{display:flex;flex-wrap:wrap;gap:1rem;margin-top:1rem}
.geom-table{flex:1 1 300px;background:#fafafa;border:1px solid #e0e0e0;border-radius:4px;padding:0.5rem}
.geom-table h4{margin:0 0 0.5rem 0;font-size:0.95rem}
.geom-table table{width:100%;border-collapse:collapse;font-size:0.85rem}
.geom-table th,.geom-table td{border:1px solid #ddd;padding:0.25rem 0.35rem;text-align:left}
.geom-table tbody tr:nth-child(odd){background:#fff}
.geom-tag{padding:0.1rem 0.35rem;border-radius:999px;font-weight:600;font-size:0.8rem;display:inline-block}
.geom-desc{font-size:0.8rem;color:#444}
.geom-row-Missing{background:#fde0e0}
.geom-row-Extra{background:#e1f5e1}
.geom-row-Moved{background:#fff2da}
.geom-row-Added{background:#ede7f6}
.geom-row-Removed{background:#e3f2fd}
.geom-row-Changed{background:#fbe9e7}
.geom-tag-Missing{background:#d32f2f;color:#fff}
.geom-tag-Extra{background:#2e7d32;color:#fff}
.geom-tag-Moved{background:#f57c00;color:#fff}
.geom-tag-Added{background:#7b1fa2;color:#fff}
.geom-tag-Removed{background:#1565c0;color:#fff}
.geom-tag-Changed{background:#6d4c41;color:#fff}
.geom-explain{list-style:none;padding:0;margin:0.75rem 0;display:flex;gap:1rem;flex-wrap:wrap;font-size:0.85rem;color:#333}
.geom-explain li{display:flex;align-items:center;gap:0.35rem}
.geom-warning{margin-top:0.5rem;padding:0.5rem 0.75rem;border-radius:4px;border:1px solid #ffc107;background:#fff8e1;color:#7a5a00;font-size:0.85rem}
</style>
""");
        sb.AppendLine("</head><body>");
        sb.AppendLine("<h1>PDF Visual Diff</h1>");

        sb.AppendLine("<section class=\"summary-grid\">");
        sb.AppendLine(SummaryCard("Baseline", HtmlEncode(summary.BaselinePath)));
        sb.AppendLine(SummaryCard("Candidate", HtmlEncode(summary.CandidatePath)));
        sb.AppendLine(SummaryCard("DPI", summary.Dpi.ToString(CultureInfo.InvariantCulture)));
        sb.AppendLine(SummaryCard("Pages Compared", $"{summary.ComparedPages.Count}"));
        sb.AppendLine(SummaryCard("Total AE Pixels", summary.TotalAePixels.ToString(CultureInfo.InvariantCulture)));
        sb.AppendLine(SummaryCard("Weighted AE %", $"{summary.WeightedAePercent:F6}"));
        sb.AppendLine(SummaryCard("Consistent pages", $"{summary.ConsistentPages}"));
        sb.AppendLine(SummaryCard("Engine-only pages", $"{summary.EngineOnlyPages}"));
        sb.AppendLine(SummaryCard("Clean pages", $"{summary.CleanPages}"));
        sb.AppendLine(SummaryCard("Overall Status", $"<span class=\"badge status-{summary.OverallStatus}\">{summary.OverallStatus}</span>"));
        sb.AppendLine("</section>");

        sb.AppendLine("<ul class=\"meta\">");
        sb.AppendLine($"<li>Baseline pages: {summary.BaselinePageCount}, Candidate pages: {summary.CandidatePageCount}</li>");
        sb.AppendLine($"<li>Page selection: {HtmlEncode(summary.PageRangeExpression ?? "all")}</li>");
        sb.AppendLine($"<li>Fuzz: {HtmlEncode(summary.Fuzz ?? "none")} | CropBox: {(summary.UseCropBox ? "enabled" : "disabled")}</li>");
        sb.AppendLine($"<li>Thresholds → warning AE% ≤ {summary.Thresholds.WarningAePercent.ToString("G6", CultureInfo.InvariantCulture)}, warning SSIM ≥ {summary.Thresholds.WarningSsim.ToString("G6", CultureInfo.InvariantCulture)}</li>");
        if (summary.SsimOptions.HasValues)
        {
            sb.AppendLine($"<li>SSIM defines: {HtmlEncode(DescribeSsimOptions(summary.SsimOptions))}</li>");
        }
        if (!string.IsNullOrWhiteSpace(summary.Tools.ImageMagickVersion))
        {
            sb.AppendLine($"<li>ImageMagick: {HtmlEncode(summary.Tools.ImageMagickVersion!)}</li>");
        }
        sb.AppendLine($"<li>Generated at: {summary.GeneratedAtUtc.ToLocalTime():u}</li>");
        if (summary.GeometryRequested)
        {
            if (geometry is null)
            {
                sb.AppendLine("<li>Geometry check requested but unavailable.</li>");
            }
            else
            {
                var label = geometry.Passed ? "PASS" : "FAIL";
                var css = geometry.Passed ? "status-Pass" : "status-Fail";
                sb.AppendLine($"<li>Geometry check (<span class=\"{css}\">{label}</span>) tolerance ±{geometry.TolerancePt:F3} pt</li>");
            }
        }
        if (summary.MissingFromBaseline.Count > 0)
        {
            sb.AppendLine($"<li>Missing in baseline: {string.Join(", ", summary.MissingFromBaseline)}</li>");
        }
        if (summary.MissingFromCandidate.Count > 0)
        {
            sb.AppendLine($"<li>Missing in candidate: {string.Join(", ", summary.MissingFromCandidate)}</li>");
        }
        sb.AppendLine("</ul>");

        sb.AppendLine("<div class=\"filters\">");
        sb.AppendLine("<label><input type=\"checkbox\" data-filter value=\"Consistent\" checked> Consistent</label>");
        sb.AppendLine("<label><input type=\"checkbox\" data-filter value=\"EngineOnly\" checked> Engine-only</label>");
        sb.AppendLine("<label><input type=\"checkbox\" data-filter value=\"Clean\" checked> Clean</label>");
        sb.AppendLine("</div>");

        sb.AppendLine("<table>");
        sb.AppendLine("<thead><tr><th>Page</th><th>Diff type</th><th>Renderer</th><th>Status</th><th>Pixels (WxH)</th><th>AE</th><th>AE %</th><th>MAE</th><th>RMSE</th><th>PSNR</th><th>SSIM</th><th>DSSIM</th><th>Images</th></tr></thead>");

        foreach (var page in pages.OrderBy(p => p.PageNumber))
        {
            var renderers = page.Renderers.ToArray();
            var rowSpan = renderers.Length;
            sb.AppendLine($"<tbody class=\"page-block\" data-page=\"{page.PageNumber}\" data-classification=\"{page.Classification}\">");

            for (var index = 0; index < renderers.Length; index++)
            {
                var renderer = renderers[index];
                sb.Append("<tr>");

                if (index == 0)
                {
                    sb.Append($"<td rowspan=\"{rowSpan}\">{page.PageNumber}</td>");
                    sb.Append($"<td rowspan=\"{rowSpan}\"><span class=\"classification-badge classification-{page.Classification}\">{ClassificationLabel(page.Classification)}</span></td>");
                }

                sb.Append($"<td>{RendererLabel(renderer.Engine)}</td>");
                sb.Append($"<td class=\"status-{renderer.Status}\">{renderer.Status}</td>");
                sb.Append($"<td>{renderer.Width}×{renderer.Height}</td>");
                sb.Append($"<td>{renderer.AePixels}</td>");
                sb.Append($"<td>{renderer.AePercent:F6}</td>");
                sb.Append($"<td>{renderer.Mae:E6}</td>");
                sb.Append($"<td>{renderer.Rmse:E6}</td>");
                sb.Append($"<td>{FormatDouble(renderer.Psnr)}</td>");
                sb.Append($"<td>{FormatDouble(renderer.Ssim)}</td>");
                sb.Append($"<td>{FormatDouble(renderer.Dssim)}</td>");

                var aFull = RelativePath(outputRoot, renderer.BaselineImagePath);
                var bFull = RelativePath(outputRoot, renderer.CandidateImagePath);
                var diffFull = RelativePath(outputRoot, renderer.DiffImagePath);
                var aThumb = RelativePath(outputRoot, renderer.BaselineThumbPath);
                var bThumb = RelativePath(outputRoot, renderer.CandidateThumbPath);
                var diffThumb = RelativePath(outputRoot, renderer.DiffThumbPath);
                var overlayFull = RelativePath(outputRoot, renderer.OverlayImagePath);
                var overlayThumb = RelativePath(outputRoot, renderer.OverlayThumbPath);

                sb.Append("<td>");
                sb.Append("<div class=\"thumbs\">");
                sb.Append(ThumbLink("Baseline", aFull, aThumb));
                sb.Append(ThumbLink("Candidate", bFull, bThumb));
                sb.Append(ThumbLink("Diff", diffFull, diffThumb));
                sb.Append(ThumbLink("Overlay", overlayFull, overlayThumb));
                sb.Append("</div>");
                sb.Append("</td>");
                sb.AppendLine("</tr>");
            }

            sb.AppendLine("</tbody>");
        }

        sb.AppendLine("</table>");

        if (geometry is { Pages.Count: > 0 })
        {
            RenderGeometrySection(sb, geometry, pages, outputRoot);
        }

        sb.AppendLine("<section class=\"legend\">");
        sb.AppendLine("<h2>Metric legend</h2>");
        sb.AppendLine("<p>AE counts raw differing pixels (AE% normalises by area). MAE/RMSE track average intensity error, PSNR grows with similarity. <strong>SSIM ↑ meglio</strong>, <strong>DSSIM ↓ meglio</strong>; analizzale insieme ad AE%/PSNR per stimare l'impatto visivo.</p>");
        sb.AppendLine("</section>");

        sb.AppendLine("""
<script>
(() => {
  const toggles = document.querySelectorAll('[data-filter]');
  function updateTable() {
    const active = Array.from(toggles).filter(t => t.checked).map(t => t.value);
    document.querySelectorAll('tbody.page-block').forEach(block => {
      block.style.display = active.includes(block.dataset.classification) ? '' : 'none';
    });
  }
  toggles.forEach(t => t.addEventListener('change', updateTable));
  updateTable();
})();
(() => {
  const colors = {
    missingWords: '#d32f2f',
    extraWords: '#2e7d32',
    movedWords: '#f57c00',
    addedGraphics: '#7b1fa2',
    removedGraphics: '#1565c0',
    changedGraphics: '#6d4c41'
  };
  document.querySelectorAll('.geometry-panel').forEach(panel => {
    const dataEl = panel.querySelector('.geometry-data');
    const canvas = panel.querySelector('canvas');
    const img = panel.querySelector('img');
    if (!dataEl || !canvas || !img) return;
    let data = null;
    try {
      data = JSON.parse(dataEl.textContent || '{}');
    } catch (_e) {
      return;
    }
    const ctx = canvas.getContext('2d');
    function resize() {
      canvas.width = img.clientWidth;
      canvas.height = img.clientHeight;
      draw();
    }
    function draw() {
      if (!data) return;
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      const scaleX = canvas.width / (data.width || 1);
      const scaleY = canvas.height / (data.height || 1);
      panel.querySelectorAll('[data-geom-layer]').forEach(layer => {
        if (!layer.checked) return;
        const key = layer.dataset.geomLayer || layer.value;
        const entries = data[key];
        if (!entries) return;
        ctx.strokeStyle = colors[layer.value] || '#d32f2f';
        ctx.lineWidth = 2;
        entries.forEach(entry => {
          const box = entry.bbox;
          if (!box) return;
          const x = box.x * scaleX;
          const y = canvas.height - (box.y + box.height) * scaleY;
          const w = box.width * scaleX;
          const h = box.height * scaleY;
          ctx.strokeRect(x, y, w, h);
        });
      });
    }
    panel.querySelectorAll('[data-geom-layer]').forEach(layer => {
      layer.addEventListener('change', draw);
    });
    if (img.complete) {
      resize();
    } else {
      img.addEventListener('load', resize);
    }
    window.addEventListener('resize', resize);
  });
})();
</script>
</body></html>
""");
        File.WriteAllText(destination, sb.ToString());
    }

    public static void WriteJson(
        string destination,
        VisualDiffSummary summary,
        IReadOnlyList<PageComparisonResult> pages,
        GeometryAnalysisResult? geometry)
    {
        var payload = new
        {
            generatedAt = summary.GeneratedAtUtc,
            summary,
            pages,
            geometry
        };

        File.WriteAllText(destination, JsonSerializer.Serialize(payload, new JsonSerializerOptions
        {
            WriteIndented = true
        }));
    }

    private static void RenderGeometrySection(
        StringBuilder sb,
        GeometryAnalysisResult analysis,
        IReadOnlyList<PageComparisonResult> pageResults,
        string outputRoot)
    {
        sb.AppendLine("<section class=\"geometry-section\">");
        sb.AppendLine("<h2>Geometry (PDFBox)</h2>");
        sb.AppendLine($"<p>Advanced analysis via Apache PDFBox. Tolerance: ±{analysis.TolerancePt:F3} pt. Toggle layers to highlight problematic regions.</p>");

        foreach (var report in analysis.Pages)
        {
            var visual = pageResults.FirstOrDefault(p => p.PageNumber == report.PageNumber);
            if (visual == null)
                continue;

            var renderer = visual.Poppler;
            var preview = RelativePath(outputRoot, renderer.BaselineThumbPath);
            var overlay = BuildOverlayPayload(report, renderer);

            sb.AppendLine($"<article class=\"geometry-panel\" data-page=\"{report.PageNumber}\">");
            sb.AppendLine($"<header><strong>Page {report.PageNumber}</strong> · missing: {report.MissingWords.Count}, moved: {report.MovedWords.Count}, extra: {report.ExtraWords.Count}, graphics Δ: {report.AddedGraphics.Count + report.RemovedGraphics.Count + report.ChangedGraphics.Count}</header>");
            sb.AppendLine("<div class=\"geom-controls\">");
            sb.AppendLine("<label><input type=\"checkbox\" data-geom-layer=\"missingWords\" value=\"missingWords\" checked> Missing words</label>");
            sb.AppendLine("<label><input type=\"checkbox\" data-geom-layer=\"extraWords\" value=\"extraWords\" checked> Extra words</label>");
            sb.AppendLine("<label><input type=\"checkbox\" data-geom-layer=\"movedWords\" value=\"movedWords\" checked> Moved words</label>");
            sb.AppendLine("<label><input type=\"checkbox\" data-geom-layer=\"addedGraphics\" value=\"addedGraphics\" checked> Added graphics</label>");
            sb.AppendLine("<label><input type=\"checkbox\" data-geom-layer=\"removedGraphics\" value=\"removedGraphics\" checked> Removed graphics</label>");
            sb.AppendLine("<label><input type=\"checkbox\" data-geom-layer=\"changedGraphics\" value=\"changedGraphics\" checked> Changed graphics</label>");
            sb.AppendLine("</div>");
            if (report.WordStatus != GeometryWordStatus.Ok && !string.IsNullOrWhiteSpace(report.WordStatusNote))
            {
                sb.AppendLine($"<div class=\"geom-warning\">{HtmlEncode(report.WordStatusNote)}</div>");
            }
            sb.AppendLine("<ul class=\"geom-explain\">");
            sb.AppendLine("<li><span class=\"geom-tag geom-tag-Missing\">Missing</span> Baseline contiene il testo/la forma, il candidato no.</li>");
            sb.AppendLine("<li><span class=\"geom-tag geom-tag-Extra\">Extra</span> Nuovo elemento nel candidato.</li>");
            sb.AppendLine("<li><span class=\"geom-tag geom-tag-Moved\">Moved</span> Presente in entrambi ma spostato oltre la soglia.</li>");
            sb.AppendLine("</ul>");
        sb.AppendLine("<div class=\"geom-preview\">");
            sb.AppendLine($"<img src=\"{HtmlAttributeEncode(preview)}\" alt=\"Geometry preview page {report.PageNumber}\"/>");
            sb.AppendLine("<canvas></canvas>");
            sb.AppendLine("</div>");
            sb.AppendLine("<div class=\"geom-tables\">");
            WriteWordTable(sb, report);
            WriteGraphicTable(sb, report);
            sb.AppendLine("</div>");
            sb.AppendLine($"<script type=\"application/json\" class=\"geometry-data\">{HtmlEncode(overlay)}</script>");
            sb.AppendLine("</article>");
        }

        sb.AppendLine("</section>");
    }

    private static void WriteWordTable(StringBuilder sb, GeometryPageReport report)
    {
        sb.AppendLine("<div class=\"geom-table\">");
        sb.AppendLine("<h4>Words</h4>");
        sb.AppendLine("<table>");
        sb.AppendLine("<thead><tr><th>Tipo</th><th>Significato</th><th>Testo</th><th>BBox</th><th>Δx</th><th>Δy</th><th>Δ</th></tr></thead><tbody>");
        AppendWordRows(sb, "Missing", report.MissingWords, preferCandidate: false);
        AppendWordRows(sb, "Extra", report.ExtraWords, preferCandidate: true);
        AppendWordRows(sb, "Moved", report.MovedWords, preferCandidate: true);
        sb.AppendLine("</tbody></table>");
        sb.AppendLine("</div>");
    }

    private static void WriteGraphicTable(StringBuilder sb, GeometryPageReport report)
    {
        sb.AppendLine("<div class=\"geom-table\">");
        sb.AppendLine("<h4>Graphics</h4>");
        sb.AppendLine("<table>");
        sb.AppendLine("<thead><tr><th>Tipo</th><th>Significato</th><th>Baseline bbox</th><th>Candidate bbox</th><th>Stroke</th><th>Δ</th></tr></thead><tbody>");
        AppendGraphicRows(sb, "Added", report.AddedGraphics);
        AppendGraphicRows(sb, "Removed", report.RemovedGraphics);
        AppendGraphicRows(sb, "Changed", report.ChangedGraphics);
        sb.AppendLine("</tbody></table>");
        sb.AppendLine("</div>");
    }

    private static void AppendWordRows(StringBuilder sb, string label, IEnumerable<WordDifference> words, bool preferCandidate)
    {
        var description = GetWordDescription(label);
        var tagClass = $"geom-tag geom-tag-{label}";
        var rowClass = $"geom-row-{label}";

        foreach (var word in words)
        {
            var bbox = preferCandidate ? word.Candidate ?? word.Baseline : word.Baseline ?? word.Candidate;
            sb.Append($"<tr class=\"{rowClass}\">");
            sb.Append($"<td><span class=\"{tagClass}\">{label}</span></td>");
            sb.Append($"<td class=\"geom-desc\">{HtmlEncode(description)}</td>");
            sb.Append($"<td>{HtmlEncode(word.Text)}</td>");
            sb.Append($"<td>{FormatBoundingBox(bbox)}</td>");
            sb.Append($"<td>{word.Delta.DeltaX:F2}</td>");
            sb.Append($"<td>{word.Delta.DeltaY:F2}</td>");
            sb.Append($"<td>{FormatDistance(word.Delta.MaxDeviation)}</td>");
            sb.AppendLine("</tr>");
        }
    }

    private static void AppendGraphicRows(StringBuilder sb, string label, IEnumerable<GraphicDifference> graphics)
    {
        var description = GetGraphicDescription(label);
        var tagClass = $"geom-tag geom-tag-{label}";
        var rowClass = $"geom-row-{label}";

        foreach (var graphic in graphics)
        {
            sb.Append($"<tr class=\"{rowClass}\">");
            sb.Append($"<td><span class=\"{tagClass}\">{label}</span></td>");
            sb.Append($"<td class=\"geom-desc\">{HtmlEncode(description)}</td>");
            sb.Append($"<td>{FormatBoundingBox(graphic.Baseline)}</td>");
            sb.Append($"<td>{FormatBoundingBox(graphic.Candidate)}</td>");
            sb.Append($"<td>{graphic.StrokeWidth:F2}</td>");
            sb.Append($"<td>{FormatDistance(graphic.Delta.MaxDeviation)}</td>");
            sb.AppendLine("</tr>");
        }
    }

    private static string BuildOverlayPayload(GeometryPageReport report, RendererComparisonResult renderer)
    {
        var payload = new
        {
            width = renderer.Width,
            height = renderer.Height,
            missingWords = BuildWordOverlay(report.MissingWords, preferCandidate: false),
            extraWords = BuildWordOverlay(report.ExtraWords, preferCandidate: true),
            movedWords = BuildWordOverlay(report.MovedWords, preferCandidate: true),
            addedGraphics = BuildGraphicOverlay(report.AddedGraphics, preferCandidate: true),
            removedGraphics = BuildGraphicOverlay(report.RemovedGraphics, preferCandidate: false),
            changedGraphics = BuildGraphicOverlay(report.ChangedGraphics, preferCandidate: true)
        };

        return JsonSerializer.Serialize(payload, OverlayJsonOptions);
    }

    private static IEnumerable<object> BuildWordOverlay(IEnumerable<WordDifference> words, bool preferCandidate)
    {
        foreach (var word in words)
        {
            var bbox = preferCandidate ? word.Candidate ?? word.Baseline : word.Baseline ?? word.Candidate;
            var entry = ToOverlayEntry(bbox, word.Text);
            if (entry != null)
                yield return entry;
        }
    }

    private static IEnumerable<object> BuildGraphicOverlay(IEnumerable<GraphicDifference> graphics, bool preferCandidate)
    {
        foreach (var graphic in graphics)
        {
            var bbox = preferCandidate ? graphic.Candidate ?? graphic.Baseline : graphic.Baseline ?? graphic.Candidate;
            var entry = ToOverlayEntry(bbox, graphic.Type);
            if (entry != null)
                yield return entry;
        }
    }

    private static object? ToOverlayEntry(BoundingBox? box, string? label)
    {
        if (box is null)
            return null;
        return new
        {
            bbox = new { box.X, box.Y, box.Width, box.Height },
            label = label ?? string.Empty
        };
    }

    private static string FormatBoundingBox(BoundingBox? box) =>
        box is null
            ? "—"
            : $"[{box.X:F1}, {box.Y:F1}, {box.Width:F1}, {box.Height:F1}]";

    private static string FormatDistance(double value) =>
        Math.Abs(value) < 0.0001 ? "0.00" : value.ToString("F2", CultureInfo.InvariantCulture);

    private static string GetWordDescription(string label) =>
        WordDescriptions.TryGetValue(label, out var description)
            ? description
            : "Differenza sul testo.";

    private static string GetGraphicDescription(string label) =>
        GraphicDescriptions.TryGetValue(label, out var description)
            ? description
            : "Differenza vettoriale.";

    private static string SummaryCard(string title, string content) =>
        $"<article class=\"summary-card\"><div>{title}</div><div style=\"font-size:1.1rem;font-weight:600\">{content}</div></article>";

    private static string ThumbLink(string label, string full, string thumb)
    {
        var encodedLabel = HtmlEncode(label);
        var href = HtmlAttributeEncode(full);
        var src = HtmlAttributeEncode(thumb);
        return $"<figure><a href=\"{href}\" target=\"_blank\" rel=\"noreferrer\"><img src=\"{src}\" alt=\"{encodedLabel}\"/></a><figcaption>{encodedLabel}</figcaption></figure>";
    }

    private static string RelativePath(string root, string absolute)
    {
        var relative = Path.GetRelativePath(root, absolute);
        return relative.Replace('\\', '/');
    }

    private static string HtmlEncode(string input) =>
        input
            .Replace("&", "&amp;", StringComparison.Ordinal)
            .Replace("<", "&lt;", StringComparison.Ordinal)
            .Replace(">", "&gt;", StringComparison.Ordinal)
            .Replace("\"", "&quot;", StringComparison.Ordinal);

    private static string HtmlAttributeEncode(string input) =>
        HtmlEncode(input).Replace("'", "&#39;", StringComparison.Ordinal);

    private static string FormatDouble(double value) =>
        double.IsPositiveInfinity(value) ? "inf" :
        double.IsNegativeInfinity(value) ? "-inf" :
        double.IsNaN(value) ? "NaN" :
        value.ToString("G6", CultureInfo.InvariantCulture);

    private static string DescribeSsimOptions(SsimOptions options)
    {
        var parts = new List<string>();
        Append("radius", options.Radius);
        Append("sigma", options.Sigma);
        Append("k1", options.K1);
        Append("k2", options.K2);
        return parts.Count == 0 ? "default" : string.Join(", ", parts);

        void Append(string name, double? value)
        {
            if (!value.HasValue)
                return;
            parts.Add($"{name}={value.Value.ToString("G", CultureInfo.InvariantCulture)}");
        }
    }

    private static string ClassificationLabel(PageDiffClassification classification) =>
        classification switch
        {
            PageDiffClassification.Consistent => "Consistent",
            PageDiffClassification.EngineOnly => "Engine-only",
            _ => "Clean"
        };

    private static string RendererLabel(RendererEngine engine) =>
        engine switch
        {
            RendererEngine.Poppler => "Poppler (pdftoppm)",
            RendererEngine.Ghostscript => "Ghostscript (magick)",
            _ => engine.ToString()
        };
}
