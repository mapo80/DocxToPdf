using DocumentFormat.OpenXml.Wordprocessing;
using DocxToPdf.Sdk.Docx;
using DocxToPdf.Sdk.Docx.Formatting;
using DocxToPdf.Sdk.Docx.Numbering;
using DocxToPdf.Sdk.Layout;
using DocxToPdf.Sdk.Pdf;
using DocxToPdf.Sdk.Text;
using DocxToPdf.Sdk.Units;
using SkiaSharp;
using System;
using System.Linq;

namespace DocxToPdf.Sdk;

/// <summary>
/// Convertitore DOCX → PDF.
/// </summary>
public sealed class DocxToPdfConverter
{
    private readonly TextRenderer _textRenderer;
    private readonly TextLayoutEngine _layoutEngine;
    private readonly FontManager _fontManager;

    public DocxToPdfConverter()
    {
        _textRenderer = new TextRenderer();
        _layoutEngine = new TextLayoutEngine();
        _fontManager = FontManager.Instance;
    }

    public Action<string>? DiagnosticsLogger { get; set; }

    /// <summary>
    /// Converte un documento DOCX in PDF.
    /// </summary>
    /// <param name="docxPath">Percorso del file DOCX di input</param>
    /// <param name="pdfPath">Percorso del file PDF di output</param>
    public void Convert(string docxPath, string pdfPath)
    {
        using var docx = DocxDocument.Open(docxPath);

        var previousNumberingSink = NumberingDiagnostics.Sink;
        var previousTabSink = TabDiagnostics.Sink;
        NumberingDiagnostics.Sink = DiagnosticsLogger;
        TabDiagnostics.Sink = DiagnosticsLogger;

        // Leggi sezione (pagina e margini)
        var section = docx.GetSection();

        // Metadati PDF
        var metadata = new PdfMetadata
        {
            Creator = "DocxToPdf.Sdk",
            CreationDate = DateTime.Now
        };

        // Crea documento PDF
        using var pdfBuilder = PdfDocumentBuilder.Create(pdfPath, metadata);

        // Setup pagina
        var pageSize = section.PageSize;
        var margins = section.Margins;
        var contentWidth = margins.GetContentWidth(pageSize.WidthPt);
        var contentHeight = margins.GetContentHeight(pageSize.HeightPt);

        // Inizia prima pagina
        var page = pdfBuilder.BeginPage(pageSize);
        float currentY = margins.Top;

        // Rendering paragrafi
        int paragraphIndex = 0;
        foreach (var paragraph in docx.GetParagraphs())
        {
            paragraphIndex++;
            var lineSpacingDescriptor = paragraph.ParagraphFormatting.LineSpacing;
            string lineSpacingValue;
            if (lineSpacingDescriptor == null)
                lineSpacingValue = "n/a";
            else if (lineSpacingDescriptor.Value.Rule == ParagraphLineSpacingRule.Auto)
                lineSpacingValue = $"{lineSpacingDescriptor.Value.Value:F2}x";
            else
                lineSpacingValue = $"{lineSpacingDescriptor.Value.Value:F2}pt";

            DiagnosticsLogger?.Invoke(
                $"Paragraph {paragraphIndex}: spacingBefore={paragraph.ParagraphFormatting.SpacingBeforePt:F2} " +
                $"spacingAfter={paragraph.ParagraphFormatting.SpacingAfterPt:F2} " +
                $"lineSpacingRule={lineSpacingDescriptor?.Rule} " +
                $"lineSpacingValue={lineSpacingValue}");

            // Layout del paragrafo
            var lines = _layoutEngine.LayoutParagraph(paragraph, contentWidth);

            // Spaziatura prima del paragrafo
            currentY += paragraph.ParagraphFormatting.SpacingBeforePt;

            for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++)
            {
                var line = lines[lineIndex];
                // Calcola line spacing corretto usando le impostazioni del paragrafo
                var defaultSpacing = line.GetLineSpacing();
                var resolvedLineSpacing = paragraph.ParagraphFormatting.ResolveLineSpacing(defaultSpacing);

                // Verifica se serve una nuova pagina
                if (currentY + resolvedLineSpacing > pageSize.HeightPt - margins.Bottom)
                {
                    // Nuova pagina
                    pdfBuilder.EndPage();
                    page = pdfBuilder.BeginPage(pageSize);
                    currentY = margins.Top;
                }

                // Rendering della riga: il baseline è currentY - MaxAscent
                // (MaxAscent è negativo, quindi sottraendolo saliamo verso l'alto)
                float baseline = currentY - line.MaxAscent;
                float indent = line.IsFirstLine
                    ? paragraph.ParagraphFormatting.GetFirstLineOffsetPt()
                    : paragraph.ParagraphFormatting.GetSubsequentLineOffsetPt();
                float currentX = margins.Left + indent;

                // Allineamento paragrafo
                var availableWidth = line.AvailableWidthPt;
                var extraSpace = Math.Max(0f, availableWidth - line.WidthPt);
                var alignment = paragraph.ParagraphFormatting.Alignment;
                currentX += alignment switch
                {
                    ParagraphAlignment.Center => extraSpace / 2f,
                    ParagraphAlignment.Right => extraSpace,
                    _ => 0f
                };
                var shouldJustify = alignment == ParagraphAlignment.Justified && lineIndex < lines.Count - 1;
                var shouldDistribute = alignment == ParagraphAlignment.Distributed;
                var stretchableSpaces = (shouldJustify || shouldDistribute)
                    ? line.Runs.Count(IsStretchableSpace)
                    : 0;
                var perSpaceAdvance = stretchableSpaces > 0
                    ? extraSpace / stretchableSpaces
                    : 0f;

                if (paragraph.ListMarker != null && line.IsFirstLine)
                {
                    DrawListMarker(
                        page.Canvas,
                        margins,
                        paragraph,
                        paragraph.ListMarker,
                        baseline);
                    currentX = margins.Left + paragraph.ParagraphFormatting.GetSubsequentLineOffsetPt();
                }

                DrawBarTabs(page.Canvas, margins, paragraph, line, baseline);

                DiagnosticsLogger?.Invoke(
                    $"Paragraph {paragraphIndex} line {lineIndex + 1}: baseline={baseline:F2} currentY={currentY:F2} " +
                    $"lineSpacing={resolvedLineSpacing:F2} width={line.WidthPt:F2}/{line.AvailableWidthPt:F2}");

                foreach (var run in line.Runs)
                {
                    if (!run.IsDrawable)
                    {
                        currentX += run.AdvanceWidthOverride;
                        continue;
                    }

                    var color = new SKColor(run.Formatting.Color.R, run.Formatting.Color.G, run.Formatting.Color.B);
                    float width;
                    float startXForText = currentX;
                    string? literalSegment;
                    if (DiagnosticsLogger != null && (run.Text.Contains("©") || run.Text.Contains("™")))
                    {
                        DiagnosticsLogger.Invoke($"Literal candidate text='{run.Text}' chars={run.Text.Length}");
                    }
                    if (run.Shaped is { } shaped && run.ShapedLength > 0)
                    {
                        width = shaped.DrawRange(
                            page.Canvas,
                            currentX,
                            baseline,
                            color,
                            run.ShapedStart,
                            run.ShapedLength,
                            run.Formatting.CharacterSpacingPt);
                        literalSegment = run.Text;
                    }
                    else
                    {
                        width = _textRenderer.DrawShapedTextWithFallback(
                            page.Canvas,
                            run.Text,
                            currentX,
                            baseline,
                            run.Typeface,
                            run.FontSizePt,
                            color,
                            run.Formatting.CharacterSpacingPt,
                            run.Formatting.KerningEnabled
                        );
                        literalSegment = run.Text;
                    }
                    currentX += width;

                    if (!string.IsNullOrEmpty(literalSegment) && ContainsNonAscii(literalSegment))
                    {
                        _textRenderer.DrawInvisibleText(
                            page.Canvas,
                            literalSegment,
                            startXForText,
                            baseline,
                            run.Typeface,
                            run.FontSizePt);
                    }

                    if (run.Text.Contains('\u2122'))
                    {
                        DiagnosticsLogger?.Invoke($"Trademark run text='{run.Text}' width={width:F2}");
                    }

                    if ((shouldJustify || shouldDistribute) && perSpaceAdvance > 0f && IsStretchableSpace(run))
                    {
                        currentX += perSpaceAdvance;
                    }
                }

                // Avanza alla prossima riga usando line spacing
                currentY += resolvedLineSpacing;
            }

            // Spazio dopo il paragrafo
            currentY += paragraph.ParagraphFormatting.SpacingAfterPt;
        }

        pdfBuilder.EndPage();
        pdfBuilder.Close();

        NumberingDiagnostics.Sink = previousNumberingSink;
        TabDiagnostics.Sink = previousTabSink;
    }

    private void DrawListMarker(SKCanvas canvas, Margins margins, DocxParagraph paragraph, DocxListMarker marker, float baseline)
    {
        var typeface = _fontManager.GetTypeface(marker.Formatting.FontFamily, marker.Formatting.Bold, marker.Formatting.Italic);
        var color = new SKColor(marker.Formatting.Color.R, marker.Formatting.Color.G, marker.Formatting.Color.B);

        var markerAreaStart = margins.Left + paragraph.ParagraphFormatting.GetFirstLineOffsetPt();
        var contentStart = margins.Left + paragraph.ParagraphFormatting.GetSubsequentLineOffsetPt();
        var areaStart = Math.Min(markerAreaStart, contentStart);
        var rawAreaWidth = contentStart - areaStart;
        var markerWidth = _textRenderer.MeasureTextWithFallback(marker.Text, typeface, marker.Formatting.FontSizePt);
        string suffixText;
        if (marker.Suffix == LevelSuffixValues.Space)
            suffixText = " ";
        else if (marker.Suffix == LevelSuffixValues.Tab)
            suffixText = "\t";
        else
            suffixText = string.Empty;
        var suffixWidth = string.IsNullOrEmpty(suffixText)
            ? 0f
            : _textRenderer.MeasureTextWithFallback(suffixText, typeface, marker.Formatting.FontSizePt);
        var totalWidth = markerWidth + suffixWidth;
        var areaWidth = Math.Max(rawAreaWidth, totalWidth);

        float markerX = marker.Alignment switch
        {
            ParagraphAlignment.Right => areaStart,
            ParagraphAlignment.Center => areaStart + (areaWidth - totalWidth) / 2f,
            _ => contentStart - totalWidth
        };

        _textRenderer.DrawShapedTextWithFallback(
            canvas,
            marker.Text,
            markerX,
            baseline,
            typeface,
            marker.Formatting.FontSizePt,
            color);

        if (!string.IsNullOrEmpty(suffixText))
        {
            var suffixX = markerX + markerWidth;
            _textRenderer.DrawShapedTextWithFallback(
                canvas,
                suffixText,
                suffixX,
                baseline,
                typeface,
                marker.Formatting.FontSizePt,
                color);

            _textRenderer.DrawInvisibleText(
                canvas,
                suffixText,
                suffixX,
                baseline,
                typeface,
                marker.Formatting.FontSizePt);
        }
    }

    private void DrawBarTabs(SKCanvas canvas, Margins margins, DocxParagraph paragraph, LayoutLine line, float baseline)
    {
        if (line.BarTabs.Count == 0)
            return;

        foreach (var bar in line.BarTabs)
        {
            var indent = line.IsFirstLine
                ? paragraph.ParagraphFormatting.GetFirstLineOffsetPt()
                : paragraph.ParagraphFormatting.GetSubsequentLineOffsetPt();
            var x = margins.Left + indent + bar.RelativePositionPt;
            using var paint = new SKPaint
            {
                Color = new SKColor(bar.Formatting.Color.R, bar.Formatting.Color.G, bar.Formatting.Color.B),
                IsAntialias = true,
                StrokeWidth = Math.Max(0.5f, bar.Formatting.FontSizePt / 24f)
            };

            var top = baseline + line.MaxAscent;
            var bottom = baseline + line.MaxDescent;
            canvas.DrawLine(x, top, x, bottom, paint);
        }
    }

    private static bool IsStretchableSpace(LayoutRun run) =>
        run.IsDrawable && !string.IsNullOrEmpty(run.Text) && run.Text.All(static c => c == ' ');

    private static bool ContainsNonAscii(string text)
    {
        foreach (var ch in text)
        {
            if (ch > 127)
                return true;
        }
        return false;
    }
}
