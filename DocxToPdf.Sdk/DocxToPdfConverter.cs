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

        var previousSink = NumberingDiagnostics.Sink;
        NumberingDiagnostics.Sink = DiagnosticsLogger;

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
        foreach (var paragraph in docx.GetParagraphs())
        {
            // Layout del paragrafo
            var lines = _layoutEngine.LayoutParagraph(paragraph, contentWidth);

            // Spaziatura prima del paragrafo
            currentY += paragraph.ParagraphFormatting.SpacingBeforePt;

            foreach (var line in lines)
            {
                // Calcola line spacing corretto usando font metrics
                var lineSpacing = line.GetLineSpacing();

                // Verifica se serve una nuova pagina
                if (currentY + lineSpacing > pageSize.HeightPt - margins.Bottom)
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
                currentX += paragraph.ParagraphFormatting.Alignment switch
                {
                    ParagraphAlignment.Center => extraSpace / 2f,
                    ParagraphAlignment.Right => extraSpace,
                    _ => 0f
                };

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

                foreach (var run in line.Runs)
                {
                    var color = new SKColor(run.Formatting.Color.R, run.Formatting.Color.G, run.Formatting.Color.B);
                    var width = _textRenderer.DrawShapedTextWithFallback(
                        page.Canvas,
                        run.Text,
                        currentX,
                        baseline,
                        run.Typeface,
                        run.FontSizePt,
                        color
                    );
                    currentX += width;
                }

                // Avanza alla prossima riga usando line spacing
                currentY += lineSpacing;
            }

            // Spazio dopo il paragrafo
            currentY += paragraph.ParagraphFormatting.SpacingAfterPt;
        }

        pdfBuilder.EndPage();
        pdfBuilder.Close();

        NumberingDiagnostics.Sink = previousSink;
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
        var suffixText = marker.Suffix == LevelSuffixValues.Space ? " " : string.Empty;
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
            _textRenderer.DrawShapedTextWithFallback(
                canvas,
                suffixText,
                markerX + markerWidth,
                baseline,
                typeface,
                marker.Formatting.FontSizePt,
                color);
        }
    }
}
