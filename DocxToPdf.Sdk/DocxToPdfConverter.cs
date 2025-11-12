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
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocxToPdf.Sdk;

/// <summary>
/// Convertitore DOCX → PDF.
/// </summary>
public sealed class DocxToPdfConverter
{
    private const bool EnableSpacingCalibration = true;
    private readonly TextRenderer _textRenderer;
    private readonly TextLayoutEngine _layoutEngine;
    private readonly FontManager _fontManager;
    private readonly SpacingCompensator _spacingCompensator;

    public DocxToPdfConverter()
    {
        _textRenderer = new TextRenderer();
        _layoutEngine = new TextLayoutEngine();
        _fontManager = FontManager.Instance;
        _spacingCompensator = new SpacingCompensator();
        Metadata = new PdfMetadata
        {
            Creator = "DocxToPdf.Sdk"
        };
    }

    public Action<string>? DiagnosticsLogger { get; set; }
    public PdfMetadata Metadata { get; set; }

    /// <summary>
    /// Converte un documento DOCX in PDF.
    /// </summary>
    /// <param name="docxPath">Percorso del file DOCX di input</param>
    /// <param name="pdfPath">Percorso del file PDF di output</param>
    public void Convert(string docxPath, string pdfPath)
    {
        using var docx = DocxDocument.Open(docxPath);
        var layoutStressAudit = LayoutStressFontAuditor.Create(docxPath, DiagnosticsLogger);

        var previousNumberingSink = NumberingDiagnostics.Sink;
        var previousTabSink = TabDiagnostics.Sink;
        NumberingDiagnostics.Sink = DiagnosticsLogger;
        TabDiagnostics.Sink = DiagnosticsLogger;
        _textRenderer.DiagnosticsLogger = DiagnosticsLogger;

        // Leggi sezione (pagina e margini)
        var section = docx.GetSection();

        // Metadati PDF
        var metadata = Metadata with { CreationDate = Metadata.CreationDate ?? DateTime.Now };

        // Crea documento PDF
        using var pdfBuilder = PdfDocumentBuilder.Create(pdfPath, metadata);

        // Setup pagina
        var pageSize = section.PageSize;
        var margins = section.Margins;
        DiagnosticsLogger?.Invoke(
            $"Section: page={pageSize.WidthPt:F2}x{pageSize.HeightPt:F2} margins L={margins.Left:F2} T={margins.Top:F2} R={margins.Right:F2} B={margins.Bottom:F2}");
        var contentWidth = margins.GetContentWidth(pageSize.WidthPt);
        var contentHeight = margins.GetContentHeight(pageSize.HeightPt);

        // Inizia prima pagina
        var page = pdfBuilder.BeginPage(pageSize);
        float currentY = margins.Top;

        // Rendering paragrafi
        int paragraphIndex = 0;
        DocxParagraph? previousParagraph = null;
        float previousSpacingAfter = 0f;
        foreach (var block in docx.GetBlocks())
        {
            if (block is DocxTableBlock tableBlock)
            {
                // Mantieni lo spacing dopo il paragrafo precedente anche prima di una tabella:
                // Word non lo azzera implicitamente. Evitiamo quindi di sottrarre previousSpacingAfter.
                previousParagraph = null;
                previousSpacingAfter = 0f;
                RenderTable(tableBlock.Table, pdfBuilder, ref page, pageSize, margins, contentWidth, ref currentY, layoutStressAudit);
                continue;
            }

            if (block is not DocxParagraphBlock paragraphBlock)
                continue;

            var paragraph = paragraphBlock.Paragraph;
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
                $"lineSpacingValue={lineSpacingValue} " +
                $"leftIndent={paragraph.ParagraphFormatting.LeftIndentPt:F2} " +
                $"firstLine={paragraph.ParagraphFormatting.GetFirstLineOffsetPt():F2} " +
                $"subsequent={paragraph.ParagraphFormatting.GetSubsequentLineOffsetPt():F2} " +
                $"hanging={paragraph.ParagraphFormatting.HangingIndentPt:F2}");

            var suppressSpacingBetween = ShouldSuppressSpacingBetween(previousParagraph, paragraph);
            if (suppressSpacingBetween && previousSpacingAfter > 0f)
            {
                currentY -= previousSpacingAfter;
            }

            // Layout del paragrafo
            var lines = _layoutEngine.LayoutParagraph(paragraph, contentWidth);

            // Spaziatura prima del paragrafo
            var spacingBefore = suppressSpacingBetween ? 0f : paragraph.ParagraphFormatting.SpacingBeforePt;
            currentY += spacingBefore;

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
                    ? ResolveFirstLineIndent(paragraph)
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
                var spacingMode = shouldJustify
                    ? SpacingMode.Justified
                    : shouldDistribute ? SpacingMode.Distributed : SpacingMode.None;
                var stretchableSpaces = spacingMode != SpacingMode.None
                    ? line.Runs.Count(IsStretchableSpace)
                    : 0;
                var calibrationPerSpace = GetSpaceCalibration(spacingMode, line);
                var calibratedExtraSpace = extraSpace + calibrationPerSpace * stretchableSpaces;
                var perSpaceAdvance = stretchableSpaces > 0
                    ? calibratedExtraSpace / stretchableSpaces
                    : 0f;

                if (DiagnosticsLogger != null && spacingMode != SpacingMode.None)
                {
                    DiagnosticsLogger.Invoke(
                        $"[spacing] Paragraph {paragraphIndex} line {lineIndex + 1} mode={spacingMode} " +
                        $"spaces={stretchableSpaces} extra={extraSpace:F2} basePerSpace={perSpaceAdvance:F3}");
                }

                Dictionary<int, float>? perSpaceBonuses = null;
                float bonusMean = 0f;
                float[]? runWidths = null;
                if (EnableSpacingCalibration && spacingMode != SpacingMode.None && stretchableSpaces > 0 && line.Runs.Count > 0)
                {
                    runWidths = new float[line.Runs.Count];
                    for (int i = 0; i < line.Runs.Count; i++)
                    {
                        runWidths[i] = MeasureRunWidth(line.Runs[i]);
                    }

                    perSpaceBonuses = _spacingCompensator.ComputeBonuses(
                        line,
                        spacingMode,
                        runWidths,
                        stretchableSpaces,
                        out bonusMean);
                }

                if (paragraph.ListMarker != null && line.IsFirstLine)
                {
                    DrawListMarker(
                        page.Canvas,
                        margins,
                        paragraph,
                        paragraph.ListMarker,
                        baseline);
                }

                DrawBarTabs(page.Canvas, margins, paragraph, line, baseline);

                DiagnosticsLogger?.Invoke(
                    $"Paragraph {paragraphIndex} line {lineIndex + 1}: baseline={baseline:F2} currentY={currentY:F2} " +
                    $"lineSpacing={resolvedLineSpacing:F2} width={line.WidthPt:F2}/{line.AvailableWidthPt:F2}");

                for (int runIndex = 0; runIndex < line.Runs.Count; runIndex++)
                {
                    var run = line.Runs[runIndex];
                    layoutStressAudit.Observe(run);
                    if (!run.IsDrawable)
                    {
                        currentX += run.AdvanceWidthOverride;
                        continue;
                    }

                    var color = new SKColor(run.Formatting.Color.R, run.Formatting.Color.G, run.Formatting.Color.B);
                    float width;
                    if (DiagnosticsLogger != null && (run.Text.Contains("©") || run.Text.Contains("™")))
                    {
                        DiagnosticsLogger.Invoke($"Literal candidate text='{run.Text}' chars={run.Text.Length}");
                    }
                    DiagnosticsLogger?.Invoke($"Run '{run.Text}' font={run.Typeface.FamilyName} size={run.FontSizePt:F2} bold={run.Formatting.Bold} italic={run.Formatting.Italic} kerning={run.Formatting.KerningEnabled}");
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
                            run.Formatting.KerningEnabled,
                            run.Formatting.KerningEnabled
                        );
                    }
                    currentX += width;

                    if (run.Text.Contains('\u2122'))
                    {
                        DiagnosticsLogger?.Invoke($"Trademark run text='{run.Text}' width={width:F2}");
                    }

                    if (spacingMode != SpacingMode.None && stretchableSpaces > 0 && IsStretchableSpace(run))
                    {
                        var bonus = 0f;
                        if (perSpaceBonuses != null && perSpaceBonuses.TryGetValue(runIndex, out var rawBonus))
                            bonus = rawBonus - bonusMean;
                        else
                            bonus = -bonusMean;

                        currentX += perSpaceAdvance + bonus;
                    }
                }

                if (DiagnosticsLogger != null && spacingMode != SpacingMode.None)
                {
                    var targetX = margins.Left + indent + line.AvailableWidthPt;
                    DiagnosticsLogger.Invoke($"[spacing] line end={currentX - margins.Left:F2} target={targetX - margins.Left:F2}");
                }

                // Avanza alla prossima riga usando line spacing
                currentY += resolvedLineSpacing;
            }

            // Spazio dopo il paragrafo
            previousSpacingAfter = paragraph.ParagraphFormatting.SpacingAfterPt;
            currentY += previousSpacingAfter;
            previousParagraph = paragraph;
        }

        pdfBuilder.EndPage();
        pdfBuilder.Close();
        layoutStressAudit.Complete();

        NumberingDiagnostics.Sink = previousNumberingSink;
        TabDiagnostics.Sink = previousTabSink;
    }

    private static float GetSpaceCalibration(SpacingMode mode, LayoutLine line) =>
        mode switch
        {
            // leggero abbassamento per ridurre l'overshoot osservato (~+0.15pt complessivo)
            SpacingMode.Justified => 0.295f,
            SpacingMode.Distributed => 0.05f,
            _ => 0f
        };

    private static float GetRepresentativeFontSize(LayoutLine line)
    {
        foreach (var run in line.Runs)
        {
            if (run.IsDrawable && !string.IsNullOrWhiteSpace(run.Text))
                return run.FontSizePt;
        }

        return 12f;
    }

    private float MeasureRunWidth(LayoutRun run)
    {
        if (!run.IsDrawable)
            return run.AdvanceWidthOverride;

        if (run.Shaped is { } shaped && run.ShapedLength > 0)
        {
            return shaped.MeasureRange(run.ShapedStart, run.ShapedLength, run.Formatting.CharacterSpacingPt);
        }

        var width = _textRenderer.MeasureTextWithFallback(
            run.Text,
            run.Typeface,
            run.FontSizePt,
            run.Formatting.KerningEnabled,
            run.Formatting.KerningEnabled);
        if (run.Text.Length > 1 && Math.Abs(run.Formatting.CharacterSpacingPt) > 0.0001f)
        {
            width += run.Formatting.CharacterSpacingPt * (run.Text.Length - 1);
        }

        return width;
    }

    private static float ResolveFirstLineIndent(DocxParagraph paragraph)
    {
        if (paragraph.ListMarker != null)
            return paragraph.ParagraphFormatting.GetSubsequentLineOffsetPt();
        return paragraph.ParagraphFormatting.GetFirstLineOffsetPt();
    }

    private void DrawListMarker(SKCanvas canvas, Margins margins, DocxParagraph paragraph, DocxListMarker marker, float baseline)
    {
        var markerFamily = string.Equals(marker.Formatting.FontFamily, "Symbol", StringComparison.OrdinalIgnoreCase)
            ? "Times New Roman"
            : marker.Formatting.FontFamily;
        var typeface = _fontManager.GetTypeface(markerFamily, marker.Formatting.Bold, marker.Formatting.Italic);
        var color = new SKColor(marker.Formatting.Color.R, marker.Formatting.Color.G, marker.Formatting.Color.B);

        var markerAreaStart = margins.Left + paragraph.ParagraphFormatting.GetFirstLineOffsetPt();
        var contentStart = margins.Left + paragraph.ParagraphFormatting.GetSubsequentLineOffsetPt();
        var areaStart = Math.Min(markerAreaStart, contentStart);
        var areaWidth = Math.Max(0f, contentStart - areaStart);

        var textWidth = _textRenderer.MeasureTextWithFallback(
            marker.Text,
            typeface,
            marker.Formatting.FontSizePt,
            marker.Formatting.KerningEnabled,
            marker.Formatting.KerningEnabled);

        float markerX = marker.Alignment switch
        {
            ParagraphAlignment.Right => areaStart + Math.Max(areaWidth - textWidth, 0f),
            ParagraphAlignment.Center => areaStart + Math.Max((areaWidth - textWidth) / 2f, 0f),
            _ => areaStart
        };

        _textRenderer.DrawShapedTextWithFallback(
            canvas,
            marker.Text,
            markerX,
            baseline,
            typeface,
            marker.Formatting.FontSizePt,
            color,
            0f,
            marker.Formatting.KerningEnabled,
            marker.Formatting.KerningEnabled);

        if (marker.Suffix == LevelSuffixValues.Space)
        {
            var suffixX = markerX + textWidth;
            _textRenderer.DrawShapedTextWithFallback(
                canvas,
                " ",
                suffixX,
                baseline,
                typeface,
                marker.Formatting.FontSizePt,
                color,
                0f,
                marker.Formatting.KerningEnabled,
                marker.Formatting.KerningEnabled);
        }
        // LevelSuffix=Tab è già rispettato dal fatto che il contenuto parte dal subsequent indent.
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

    private void RenderTable(
        DocxTable table,
        PdfDocumentBuilder pdfBuilder,
        ref PdfPage page,
        PaperSize pageSize,
        Margins margins,
        float contentWidth,
        ref float currentY,
        LayoutStressFontAuditor fontAudit)
    {
        if (table.Rows.Count == 0)
            return;

        var columnWidths = ResolveColumnWidths(table, contentWidth);
        var columnOffsets = BuildColumnOffsets(columnWidths);
        var rowPlans = BuildRowPlans(table, columnOffsets);
        if (rowPlans.Count == 0)
            return;

        DiagnosticsLogger?.Invoke($"[table] rows={rowPlans.Count} columns={columnWidths.Length}");
        DiagnosticsLogger?.Invoke($"[table] margins.L={margins.Left:F2} startX(left)={margins.Left:F2}");

        for (int rowIndex = 0; rowIndex < rowPlans.Count; rowIndex++)
        {
            var rowPlan = rowPlans[rowIndex];
            if (currentY + rowPlan.Height > pageSize.HeightPt - margins.Bottom)
            {
                pdfBuilder.EndPage();
                page = pdfBuilder.BeginPage(pageSize);
                currentY = margins.Top;
            }

            DiagnosticsLogger?.Invoke($"[table] draw row={rowIndex} height={rowPlan.Height:F2} top={currentY:F2}");
            DrawTableRow(rowPlan, page.Canvas, margins.Left, currentY, table.Borders, columnOffsets, rowIndex, rowPlans.Count, fontAudit);
            currentY += rowPlan.Height;
        }
    }

    private List<TableRowRenderPlan> BuildRowPlans(DocxTable table, float[] columnOffsets)
    {
        var rowCount = table.Rows.Count;
        var rowPlans = new List<TableRowRenderPlan>(rowCount);
        var rowHeights = new float[rowCount];
        var perRowCells = new List<List<TableCellRenderPlan>>(rowCount);

        for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
        {
            var row = table.Rows[rowIndex];
            var cellPlans = new List<TableCellRenderPlan>();
            foreach (var cell in row.Cells)
            {
                if (cell.VerticalMerge == TableVerticalMerge.Continue)
                    continue;

                var width = GetCellWidth(columnOffsets, cell.ColumnIndex, cell.ColumnSpan);
                if (width <= 0f)
                    continue;

                var padding = table.DefaultCellPadding.ApplyOverride(cell.PaddingOverride);
                var innerWidth = Math.Max(1f, width - padding.Left - padding.Right);
                DiagnosticsLogger?.Invoke($"[table-cell] row={rowIndex} col={cell.ColumnIndex} span={cell.ColumnSpan} width={width:F2} padL={padding.Left:F2} padR={padding.Right:F2} inner={innerWidth:F2}");
                var paragraphLayouts = new List<ParagraphLayoutPlan>();
                float contentHeight = 0f;

                foreach (var paragraph in cell.Paragraphs)
                {
                    var lines = _layoutEngine.LayoutParagraph(paragraph, innerWidth);
                    var plan = new ParagraphLayoutPlan(paragraph, lines, suppressOuterSpacing: true);
                    paragraphLayouts.Add(plan);
                    contentHeight += plan.TotalHeight;
                }

                var height = padding.Top + padding.Bottom + contentHeight;
                var rowSpan = cell.VerticalMerge == TableVerticalMerge.Restart
                    ? ComputeRowSpan(table, rowIndex, cell)
                    : 1;

                var cellPlan = new TableCellRenderPlan(
                    cell.ColumnIndex,
                    cell.ColumnSpan,
                    rowSpan,
                    width,
                    innerWidth,
                    height,
                    contentHeight,
                    paragraphLayouts,
                    padding,
                    cell.VerticalAlignment,
                    cell.Borders ?? table.Borders);

                DiagnosticsLogger?.Invoke(
                    $"[table] row={rowIndex} col={cell.ColumnIndex} span={cell.ColumnSpan} rowSpan={rowSpan} width={width:F2} height={height:F2} content={contentHeight:F2}");

                cellPlans.Add(cellPlan);
                var perRowContribution = rowSpan <= 1 ? height : height / rowSpan;
                rowHeights[rowIndex] = Math.Max(rowHeights[rowIndex], perRowContribution);
            }

            if (rowHeights[rowIndex] <= 0f)
                rowHeights[rowIndex] = 18f;

            perRowCells.Add(cellPlans);
        }

        // Ensure vertically merged cells have enough combined height
        for (int rowIndex = 0; rowIndex < perRowCells.Count; rowIndex++)
        {
            foreach (var cell in perRowCells[rowIndex])
            {
                if (cell.RowSpan <= 1)
                    continue;

                var spanHeight = SumRowHeights(rowHeights, rowIndex, cell.RowSpan);
                if (spanHeight + 0.01f < cell.Height)
                {
                    var deficit = cell.Height - spanHeight;
                    var perRow = deficit / cell.RowSpan;
                    for (int r = 0; r < cell.RowSpan && rowIndex + r < rowHeights.Length; r++)
                        rowHeights[rowIndex + r] += perRow;
                }
            }
        }

        // Assign final span heights and build plans
        for (int rowIndex = 0; rowIndex < perRowCells.Count; rowIndex++)
        {
            foreach (var cell in perRowCells[rowIndex])
            {
                cell.SpannedHeight = SumRowHeights(rowHeights, rowIndex, cell.RowSpan);
            }

            rowPlans.Add(new TableRowRenderPlan(perRowCells[rowIndex], rowHeights[rowIndex]));
        }

        return rowPlans;
    }

    private void DrawTableRow(
        TableRowRenderPlan rowPlan,
        SKCanvas canvas,
        float startX,
        float rowTop,
        TableBorderSet tableBorders,
        float[] columnOffsets,
        int rowIndex,
        int totalRows,
        LayoutStressFontAuditor fontAudit)
    {
        for (int cellIndex = 0; cellIndex < rowPlan.Cells.Count; cellIndex++)
        {
            var cell = rowPlan.Cells[cellIndex];
            var left = startX + columnOffsets[Math.Min(cell.ColumnIndex, columnOffsets.Length - 1)];
            var right = startX + columnOffsets[Math.Min(cell.ColumnIndex + cell.ColumnSpan, columnOffsets.Length - 1)];
            var rect = new SKRect(left, rowTop, right, rowTop + cell.SpannedHeight);
            var cellBorders = cell.Borders ?? tableBorders;
            var isFirstColumn = cell.ColumnIndex == 0;
            var isLastColumn = cell.ColumnIndex + cell.ColumnSpan >= columnOffsets.Length - 1;
            var isFirstRow = rowIndex == 0;
            var isLastRow = rowIndex + cell.RowSpan >= totalRows;
            DrawCellBorders(canvas, rect, cellBorders, tableBorders, isFirstColumn, isLastColumn, isFirstRow, isLastRow);

            var padding = cell.Padding;
            var availableHeight = cell.SpannedHeight - padding.Top - padding.Bottom;
            var extra = Math.Max(0f, availableHeight - cell.ContentHeight);
            var offset = cell.Alignment switch
            {
                VerticalAlignment.Center => extra / 2f,
                VerticalAlignment.Bottom => extra,
                _ => 0f
            };

            var cursorY = rowTop + padding.Top + offset;
            var contentLeft = left + padding.Left;

            foreach (var paragraph in cell.Paragraphs)
            {
                var context = $"table r{rowIndex} c{cellIndex}";
                RenderParagraphLayout(paragraph, canvas, contentLeft, cell.InnerWidth, ref cursorY, fontAudit, context);
            }
        }
    }

    private void RenderParagraphLayout(
        ParagraphLayoutPlan layout,
        SKCanvas canvas,
        float startX,
        float innerWidth,
        ref float cursorY,
        LayoutStressFontAuditor fontAudit,
        string? context = null)
    {
        cursorY += layout.SpacingBefore;
        for (int i = 0; i < layout.Lines.Count; i++)
        {
            var line = layout.Lines[i];
            var spacing = layout.LineSpacings[i];
            var paragraph = layout.Paragraph;

            var indent = line.IsFirstLine
                ? paragraph.ParagraphFormatting.GetFirstLineOffsetPt()
                : paragraph.ParagraphFormatting.GetSubsequentLineOffsetPt();

            // Allineamento all'interno della cella: usa la stessa logica del renderer principale
            // basata su AvailableWidthPt (limite effettivo usato per il wrapping), senza bias.
            float currentX = startX + indent;
            var availableWidth = line.AvailableWidthPt; // già calcolato con indent/subsequent e rightIndent
            var extraSpace = Math.Max(0f, availableWidth - line.WidthPt);
            switch (paragraph.ParagraphFormatting.Alignment)
            {
                case ParagraphAlignment.Center:
                    currentX += extraSpace / 2f;
                    break;
                case ParagraphAlignment.Right:
                    currentX += extraSpace;
                    break;
            }

            var baseline = cursorY - line.MaxAscent;
            if (context != null && DiagnosticsLogger != null)
            {
                var centerX = startX + innerWidth / 2f;
                DiagnosticsLogger.Invoke($"[table-line-x] {context} i={i} startX={startX:F2} indent={indent:F2} avail={availableWidth:F2} lineW={line.WidthPt:F2} inner={innerWidth:F2} currentX={currentX:F2} centerX={centerX:F2} dx={currentX + line.WidthPt/2f - centerX:F2}");
            }
            foreach (var run in line.Runs)
            {
                fontAudit.Observe(run);
                if (!run.IsDrawable)
                {
                    currentX += run.AdvanceWidthOverride;
                    continue;
                }

                var color = new SKColor(run.Formatting.Color.R, run.Formatting.Color.G, run.Formatting.Color.B);
                float width;
                if (run.Shaped is { } shaped && run.ShapedLength > 0)
                {
                    width = shaped.DrawRange(
                        canvas,
                        currentX,
                        baseline,
                        color,
                        run.ShapedStart,
                        run.ShapedLength,
                        run.Formatting.CharacterSpacingPt);
                }
                else
                {
                    width = _textRenderer.DrawShapedTextWithFallback(
                        canvas,
                        run.Text,
                        currentX,
                        baseline,
                        run.Typeface,
                        run.FontSizePt,
                        color,
                        run.Formatting.CharacterSpacingPt,
                        run.Formatting.KerningEnabled,
                        run.Formatting.KerningEnabled);
                }

                currentX += width;
            }

            if (context != null && DiagnosticsLogger != null)
            {
                DiagnosticsLogger.Invoke(
                    $"[table-line] {context} baseline={baseline:F2} cursorY={cursorY:F2} spacing={spacing:F2} ascent={line.MaxAscent:F2}");
            }

            cursorY += spacing;
        }

        cursorY += layout.SpacingAfter;
    }

    private void DrawCellBorders(
        SKCanvas canvas,
        SKRect rect,
        TableBorderSet cellBorders,
        TableBorderSet tableBorders,
        bool isFirstColumn,
        bool isLastColumn,
        bool isFirstRow,
        bool isLastRow)
    {
        var top = isFirstRow ? SelectBorder(cellBorders.Top, tableBorders.Top) : null;
        var bottomFallback = isLastRow ? tableBorders.Bottom : tableBorders.InsideHorizontal;
        var bottom = SelectBorder(cellBorders.Bottom, bottomFallback);
        var left = isFirstColumn ? SelectBorder(cellBorders.Left, tableBorders.Left) : null;
        var rightFallback = isLastColumn ? tableBorders.Right : tableBorders.InsideVertical;
        var right = SelectBorder(cellBorders.Right, rightFallback);

        if (top != null)
            DrawBorderLine(canvas, rect.Left, rect.Top, rect.Right, rect.Top, top);
        if (bottom != null)
            DrawBorderLine(canvas, rect.Left, rect.Bottom, rect.Right, rect.Bottom, bottom);
        if (left != null)
            DrawBorderLine(canvas, rect.Left, rect.Top, rect.Left, rect.Bottom, left);
        if (right != null)
            DrawBorderLine(canvas, rect.Right, rect.Top, rect.Right, rect.Bottom, right);
    }

    private static TableBorderStyle? SelectBorder(TableBorderStyle? cellStyle, TableBorderStyle? tableStyle) =>
        cellStyle ?? tableStyle;

    private static void DrawBorderLine(SKCanvas canvas, float x1, float y1, float x2, float y2, TableBorderStyle? style)
    {
        if (style == null || style.WidthPt <= 0f)
            return;

        using var paint = new SKPaint
        {
            IsAntialias = true,
            Style = SKPaintStyle.Stroke,
            StrokeWidth = style.WidthPt,
            Color = new SKColor(style.Color.R, style.Color.G, style.Color.B)
        };
        canvas.DrawLine(x1, y1, x2, y2, paint);
    }

    private static bool ShouldSuppressSpacingBetween(DocxParagraph? previous, DocxParagraph current)
    {
        if (previous == null)
            return false;

        if (!previous.ParagraphFormatting.SuppressSpacingBetweenSameStyle ||
            !current.ParagraphFormatting.SuppressSpacingBetweenSameStyle)
            return false;

        return string.Equals(previous.StyleId, current.StyleId, StringComparison.OrdinalIgnoreCase);
    }

    private float[] ResolveColumnWidths(DocxTable table, float availableWidth)
    {
        // Se è presente una grid (tblGrid), Word usa tali larghezze come base senza
        // ridistribuzioni uniformi. Evitiamo scaling globale che introduce shift.
        var hasGrid = table.ColumnGridPt.Count > 0;
        var columnCount = hasGrid
            ? table.ColumnGridPt.Count
            : table.Rows.Select(r => r.Cells.Sum(c => c.ColumnSpan)).DefaultIfEmpty(1).Max();

        columnCount = Math.Max(1, columnCount);

        var widths = hasGrid
            ? table.ColumnGridPt.Select(w => Math.Max(1f, w)).ToArray()
            : Enumerable.Repeat(Math.Max(1f, availableWidth / columnCount), columnCount).ToArray();

        // In auto-layout, applica preferenze per-colonna come massimo, senza scaling globale.
        if (table.LayoutType == TableLayoutType.Auto)
        {
            var preferredPerColumn = new float[columnCount];
            foreach (var row in table.Rows)
            {
                foreach (var cell in row.Cells)
                {
                    if (cell.PreferredWidth is not TablePreferredWidth cellPreferred)
                        continue;

                    var preferred = ResolvePreferredWidth(cellPreferred, availableWidth);
                    if (preferred <= 0f)
                        continue;

                    var span = Math.Max(1, cell.ColumnSpan);
                    var perColumn = preferred / span;
                    for (int i = 0; i < span && cell.ColumnIndex + i < preferredPerColumn.Length; i++)
                    {
                        preferredPerColumn[cell.ColumnIndex + i] = Math.Max(preferredPerColumn[cell.ColumnIndex + i], perColumn);
                    }
                }
            }

            for (int i = 0; i < widths.Length; i++)
            {
                if (preferredPerColumn[i] > 0f)
                    widths[i] = Math.Max(widths[i], preferredPerColumn[i]);
            }
        }

        // Niente scaling globale: se non c'è grid e la somma è zero, riparti equamente
        var sum = widths.Sum();
        if (!hasGrid && sum <= 0f)
        {
            var even = Math.Max(1f, availableWidth / columnCount);
            for (int i = 0; i < widths.Length; i++)
                widths[i] = even;
        }

        return widths;
    }

    private static float[] BuildColumnOffsets(IReadOnlyList<float> columnWidths)
    {
        var offsets = new float[columnWidths.Count + 1];
        for (int i = 0; i < columnWidths.Count; i++)
            offsets[i + 1] = offsets[i] + columnWidths[i];
        return offsets;
    }

    private static float GetCellWidth(float[] columnOffsets, int columnIndex, int columnSpan)
    {
        if (columnOffsets.Length == 0)
            return 0f;

        var start = Math.Clamp(columnIndex, 0, columnOffsets.Length - 1);
        var end = Math.Clamp(columnIndex + columnSpan, 0, columnOffsets.Length - 1);
        if (end < start)
            end = start;
        return columnOffsets[end] - columnOffsets[start];
    }

    private static float SumRowHeights(IReadOnlyList<float> heights, int startIndex, int span)
    {
        float sum = 0f;
        for (int i = 0; i < span && startIndex + i < heights.Count; i++)
            sum += heights[startIndex + i];
        return sum;
    }

    private static int ComputeRowSpan(DocxTable table, int rowIndex, DocxTableCell cell)
    {
        if (cell.VerticalMerge != TableVerticalMerge.Restart)
            return 1;

        var span = 1;
        for (int nextRow = rowIndex + 1; nextRow < table.Rows.Count; nextRow++)
        {
            var continuation = table.Rows[nextRow].Cells.FirstOrDefault(c =>
                c.VerticalMerge == TableVerticalMerge.Continue &&
                c.ColumnIndex == cell.ColumnIndex &&
                c.ColumnSpan == cell.ColumnSpan);

            if (continuation == null)
                break;

            span++;
        }

        return span;
    }

    private static float ResolvePreferredWidth(TablePreferredWidth preferredWidth, float availableWidth) =>
        preferredWidth.Unit switch
        {
            TableWidthUnit.Dxa => UnitConverter.DxaToPoints(preferredWidth.Value),
            TableWidthUnit.Percent => availableWidth * preferredWidth.Value / 5000f,
            _ => 0f
        };

    private sealed class LayoutStressFontAuditor
    {
        private static readonly HashSet<string> TargetFamilies = new(StringComparer.OrdinalIgnoreCase)
        {
            "Aptos",
            "Cambria",
            "Calibri"
        };

        private static readonly HashSet<string> ForbiddenFallbacks = new(StringComparer.OrdinalIgnoreCase)
        {
            "Caladea",
            "Carlito"
        };

        private readonly bool _enabled;
        private readonly Action<string> _logger;
        private readonly string _prefix;
        private readonly HashSet<string> _hits = new(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _fallbacks = new(StringComparer.OrdinalIgnoreCase);

        private LayoutStressFontAuditor(bool enabled, Action<string>? logger, string docName)
        {
            _enabled = enabled;
            _prefix = "[layout-stress-fonts]";
            _logger = enabled
                ? logger ?? (message => Console.Error.WriteLine(message))
                : _ => { };

            if (_enabled)
                _logger($"{_prefix} Tracking Microsoft font usage for '{docName}'");
        }

        public static LayoutStressFontAuditor Create(string docxPath, Action<string>? logger)
        {
            var filename = Path.GetFileName(docxPath);
            var enabled = string.Equals(filename, "layout-stress.docx", StringComparison.OrdinalIgnoreCase);
            return new LayoutStressFontAuditor(enabled, logger, filename ?? docxPath);
        }

        public void Observe(LayoutRun run)
        {
            if (!_enabled || !run.IsDrawable)
                return;

            var actual = run.Typeface?.FamilyName ?? "(unknown)";
            var requested = run.Formatting.FontFamily;
            var preview = BuildPreview(run.Text);
            _logger($"{_prefix} text='{preview}' requested={requested} resolved={actual}");

            if (TargetFamilies.Contains(actual))
                _hits.Add(actual);

            if (ForbiddenFallbacks.Contains(actual))
                _fallbacks.Add(actual);
        }

        public void Complete()
        {
            if (!_enabled)
                return;

            var missing = TargetFamilies.Where(f => !_hits.Contains(f)).ToList();
            if (_fallbacks.Count == 0 && missing.Count == 0)
            {
                _logger($"{_prefix} ✓ All runs used Aptos/Cambria/Calibri without Caladea/Carlito fallback.");
                return;
            }

            if (_fallbacks.Count > 0)
                _logger($"{_prefix} ⚠ Unexpected fallback fonts detected: {string.Join(", ", _fallbacks)}");
            if (missing.Count > 0)
                _logger($"{_prefix} ⚠ Missing required families: {string.Join(", ", missing)}");
        }

        private static string BuildPreview(string text)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;

            var flattened = text.Replace('\r', ' ').Replace('\n', ' ');
            return flattened.Length <= 32 ? flattened : $"{flattened[..32]}…";
        }
    }

    private sealed class SpacingCompensator
    {
        private readonly SpacingCoefficients _justified = new(-0.111f, 0.00333f, -0.00052f, -0.8f, 0.8f);
        private readonly SpacingCoefficients _distributed = new(0.0112f, -0.00245f, 0.00254f, -0.5f, 0.5f);

        public Dictionary<int, float> ComputeBonuses(
            LayoutLine line,
            SpacingMode mode,
            float[] runWidths,
            int stretchableSpaces,
            out float mean)
        {
            mean = 0f;
            var bonuses = new Dictionary<int, float>();

            if (mode == SpacingMode.None || stretchableSpaces == 0)
                return bonuses;

            var coefficients = mode switch
            {
                SpacingMode.Justified => _justified,
                SpacingMode.Distributed => _distributed,
                _ => _justified
            };

            float previousWidth = 0f;
            bool hasPrev = false;
            float sum = 0f;

            for (int runIndex = 0; runIndex < line.Runs.Count; runIndex++)
            {
                var run = line.Runs[runIndex];
                if (!run.IsDrawable)
                    continue;

                if (!string.IsNullOrWhiteSpace(run.Text))
                {
                    previousWidth = runWidths[runIndex];
                    hasPrev = true;
                    continue;
                }

                if (!IsStretchableSpace(run))
                    continue;

                var nextWidth = FindNextDrawableWidth(line, runWidths, runIndex + 1);
                var prev = hasPrev ? previousWidth : run.FontSizePt;
                if (nextWidth <= 0f)
                    nextWidth = prev;

                var bonus = coefficients.Intercept +
                            coefficients.PrevSlope * prev +
                            coefficients.NextSlope * nextWidth;
                bonus = Math.Clamp(bonus, coefficients.MinClamp, coefficients.MaxClamp);
                if (Math.Abs(bonus) > 0.0001f)
                {
                    bonuses[runIndex] = bonus;
                    sum += bonus;
                }
            }

            mean = stretchableSpaces > 0 ? sum / stretchableSpaces : 0f;
            return bonuses;
        }

        private static float FindNextDrawableWidth(LayoutLine line, float[] runWidths, int startIndex)
        {
            for (int i = startIndex; i < line.Runs.Count; i++)
            {
                var candidate = line.Runs[i];
                if (candidate.IsDrawable && !string.IsNullOrWhiteSpace(candidate.Text))
                    return runWidths[i];
            }

            return 0f;
        }
    }

    private readonly record struct SpacingCoefficients(
        float Intercept,
        float PrevSlope,
        float NextSlope,
        float MinClamp,
        float MaxClamp);

    private sealed class ParagraphLayoutPlan
    {
        public ParagraphLayoutPlan(
            DocxParagraph paragraph,
            IReadOnlyList<LayoutLine> lines,
            bool suppressOuterSpacing = false)
        {
            Paragraph = paragraph;
            Lines = lines;
            SpacingBefore = suppressOuterSpacing ? 0f : paragraph.ParagraphFormatting.SpacingBeforePt;
            SpacingAfter = suppressOuterSpacing ? 0f : paragraph.ParagraphFormatting.SpacingAfterPt;
            LineSpacings = new List<float>(lines.Count);
            float total = SpacingBefore + SpacingAfter;
            foreach (var line in lines)
            {
                var spacing = paragraph.ParagraphFormatting.ResolveLineSpacing(line.GetLineSpacing());
                LineSpacings.Add(spacing);
                total += spacing;
            }

            TotalHeight = total;
        }

        public DocxParagraph Paragraph { get; }
        public IReadOnlyList<LayoutLine> Lines { get; }
        public List<float> LineSpacings { get; }
        public float SpacingBefore { get; }
        public float SpacingAfter { get; }
        public float TotalHeight { get; }
    }

    private sealed class TableCellRenderPlan
    {
        public TableCellRenderPlan(
            int columnIndex,
            int columnSpan,
            int rowSpan,
            float width,
            float innerWidth,
            float height,
            float contentHeight,
            IReadOnlyList<ParagraphLayoutPlan> paragraphs,
            TablePadding padding,
            VerticalAlignment alignment,
            TableBorderSet borders)
        {
            ColumnIndex = Math.Max(0, columnIndex);
            ColumnSpan = Math.Max(1, columnSpan);
            RowSpan = Math.Max(1, rowSpan);
            Width = width;
            InnerWidth = innerWidth;
            Height = height;
            ContentHeight = contentHeight;
            Paragraphs = paragraphs;
            Padding = padding;
            Alignment = alignment;
            Borders = borders;
            SpannedHeight = height;
        }

        public int ColumnIndex { get; }
        public int ColumnSpan { get; }
        public int RowSpan { get; }
        public float Width { get; }
        public float InnerWidth { get; }
        public float Height { get; }
        public float ContentHeight { get; }
        public IReadOnlyList<ParagraphLayoutPlan> Paragraphs { get; }
        public TablePadding Padding { get; }
        public VerticalAlignment Alignment { get; }
        public TableBorderSet Borders { get; }
        public float SpannedHeight { get; set; }
    }

    private sealed record TableRowRenderPlan(
        IReadOnlyList<TableCellRenderPlan> Cells,
        float Height);

    private enum SpacingMode
    {
        None,
        Justified,
        Distributed
    }
}
