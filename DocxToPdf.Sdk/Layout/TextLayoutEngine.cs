using DocxToPdf.Sdk.Docx;
using DocxToPdf.Sdk.Docx.Formatting;
using DocxToPdf.Sdk.Text;
using DocxToPdf.Sdk.Units;
using SkiaSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocxToPdf.Sdk.Layout;

/// <summary>
/// Engine di layout per testo con wrapping automatico.
/// Usa algoritmo greedy per spezzare il testo in righe.
/// </summary>
public sealed class TextLayoutEngine
{
    private readonly TextRenderer _textRenderer;
    private readonly FontManager _fontManager;

    public TextLayoutEngine()
    {
        _textRenderer = new TextRenderer();
        _fontManager = FontManager.Instance;
    }

    /// <summary>
    /// Layout di un paragrafo con wrapping automatico.
    /// Restituisce le righe pronte per il rendering.
    /// </summary>
    public List<LayoutLine> LayoutParagraph(
        DocxParagraph paragraph,
        float maxWidthPt)
    {
        var lines = new List<LayoutLine>();
        var currentLine = new List<LayoutRun>();
        var currentBarTabs = new List<BarTabInstruction>();
        float currentLineWidth = 0f;
        float currentMaxAscent = 0f;
        float currentMaxDescent = 0f;
        float currentMaxLeading = 0f;

        var formatting = paragraph.ParagraphFormatting;
        var firstLineOffset = formatting.GetFirstLineOffsetPt();
        var subsequentOffset = formatting.GetSubsequentLineOffsetPt();
        var rightIndent = formatting.RightIndentPt;
        var hasMarker = paragraph.ListMarker != null;
        var textIndentFirstLine = hasMarker ? subsequentOffset : firstLineOffset;

        float SafeWidth(float width) => Math.Max(1f, width);

        var firstLineWidth = SafeWidth(maxWidthPt - textIndentFirstLine - rightIndent);
        var otherLinesWidth = SafeWidth(maxWidthPt - subsequentOffset - rightIndent);
        var currentLineLimit = firstLineWidth;
        var isFirstLine = true;
        var currentLineIndent = textIndentFirstLine;
        var elements = paragraph.InlineElements.Count > 0
            ? paragraph.InlineElements
            : paragraph.Runs.Select(r => (DocxInlineElement)new DocxTextInline(r.Text, r.Formatting)).ToArray();

        void CommitLine()
        {
            if (currentLine.Count == 0 && currentLineWidth <= 0f && currentBarTabs.Count == 0)
                return;

            lines.Add(new LayoutLine(
                new List<LayoutRun>(currentLine),
                currentLineWidth,
                currentMaxAscent,
                currentMaxDescent,
                currentMaxLeading,
                isFirstLine,
                currentLineLimit,
                new List<BarTabInstruction>(currentBarTabs)
            ));

            currentLine.Clear();
            currentBarTabs.Clear();
            currentLineWidth = 0f;
            currentMaxAscent = 0f;
            currentMaxDescent = 0f;
            currentMaxLeading = 0f;
            isFirstLine = false;
            currentLineLimit = otherLinesWidth;
            currentLineIndent = subsequentOffset;
        }

        TabResolution CreateResolutionForPositionalTab(DocxPositionalTabInline positionalTab)
        {
            var absolute = positionalTab.Reference switch
            {
                PositionalTabReference.Indent => currentLineIndent + positionalTab.PositionPt,
                _ => positionalTab.PositionPt
            };
            return new TabResolution(absolute, positionalTab.Leader, positionalTab.Alignment, false);
        }

        bool NeedsLookAhead(TabAlignment alignment) =>
            alignment == TabAlignment.Right || alignment == TabAlignment.Center || alignment == TabAlignment.Decimal;

        float ComputeDesiredStart(float relativeTarget, TabAlignment alignment, SegmentMeasurement measurement) =>
            alignment switch
            {
                TabAlignment.Right => relativeTarget - measurement.TotalWidth,
                TabAlignment.Center => relativeTarget - (measurement.TotalWidth / 2f),
                TabAlignment.Decimal when measurement.HasDecimal => relativeTarget - measurement.WidthBeforeDecimal,
                TabAlignment.Decimal => relativeTarget - measurement.TotalWidth,
                _ => relativeTarget
            };

        SegmentMeasurement MeasureSegment(IReadOnlyList<DocxInlineElement> source, int startIndex)
        {
            float total = 0f;
            float widthBeforeDecimal = 0f;
            bool hasDecimal = false;

            for (var i = startIndex; i < source.Count; i++)
            {
                switch (source[i])
                {
                    case DocxTextInline textInline:
                        var typeface = GetTypefaceForFormatting(textInline.Formatting);
                        var width = _textRenderer.MeasureTextWithFallback(textInline.Text, typeface, textInline.Formatting.FontSizePt);

                        if (!hasDecimal)
                        {
                            var decimalIndex = FindDecimalIndex(textInline.Text);
                            if (decimalIndex >= 0)
                            {
                                var prefix = textInline.Text[..decimalIndex];
                                var prefixWidth = _textRenderer.MeasureTextWithFallback(prefix, typeface, textInline.Formatting.FontSizePt);
                                widthBeforeDecimal = total + prefixWidth;
                                hasDecimal = true;
                            }
                        }

                        total += width;
                        break;
                    case DocxTabInline:
                    case DocxPositionalTabInline:
                        return new SegmentMeasurement(total, hasDecimal ? widthBeforeDecimal : total, hasDecimal);
                }
            }

            return new SegmentMeasurement(total, hasDecimal ? widthBeforeDecimal : total, hasDecimal);
        }

        void AppendLeaderOrPlaceholder(float spanWidth, TabLeader leader, RunFormatting formatting)
        {
            if (leader == TabLeader.None)
            {
                if (spanWidth <= 0f)
                    return;

                currentLine.Add(LayoutRun.Placeholder(spanWidth));
                currentLineWidth += spanWidth;
                return;
            }

            if (spanWidth <= 0f)
                return;

            var typeface = GetTypefaceForFormatting(formatting);
            var leaderText = BuildLeaderString(spanWidth, typeface, formatting.FontSizePt, leader);
            if (string.IsNullOrEmpty(leaderText))
            {
                currentLine.Add(LayoutRun.Placeholder(spanWidth));
                currentLineWidth += spanWidth;
                return;
            }

            var width = _textRenderer.MeasureTextWithFallback(leaderText, typeface, formatting.FontSizePt);
            currentLine.Add(new LayoutRun(leaderText, typeface, formatting.FontSizePt, formatting));
            currentLineWidth += width;
            var metrics = _textRenderer.GetFontMetrics(typeface, formatting.FontSizePt);
            currentMaxAscent = Math.Min(currentMaxAscent, metrics.Ascent);
            currentMaxDescent = Math.Max(currentMaxDescent, metrics.Descent);
            currentMaxLeading = Math.Max(currentMaxLeading, metrics.Leading);
        }

        for (var index = 0; index < elements.Count; index++)
        {
            switch (elements[index])
            {
                case DocxTextInline textInline:
                    ProcessTextInline(textInline);
                    break;
                case DocxTabInline tabInline:
                    ProcessTabInline(tabInline.Formatting, index, elements, null);
                    break;
                case DocxPositionalTabInline positionalTab:
                    ProcessTabInline(positionalTab.Formatting, index, elements, CreateResolutionForPositionalTab(positionalTab));
                    break;
            }
        }

        if (currentLine.Count > 0 || currentLineWidth > 0f || currentBarTabs.Count > 0)
        {
            CommitLine();
        }

        if (lines.Count == 0)
        {
            var defaultTypeface = _fontManager.GetDefaultTypeface();
            var defaultFontSize = paragraph.Runs.FirstOrDefault()?.Formatting.FontSizePt ?? 11f;
            var defaultMetrics = _textRenderer.GetFontMetrics(defaultTypeface, defaultFontSize);
            lines.Add(new LayoutLine(
                new List<LayoutRun>(),
                0f,
                defaultMetrics.Ascent,
                defaultMetrics.Descent,
                defaultMetrics.Leading,
                true,
                firstLineWidth,
                Array.Empty<BarTabInstruction>()
            ));
        }

        return lines;

        void ProcessTextInline(DocxTextInline textInline)
        {
            if (string.IsNullOrEmpty(textInline.Text))
                return;

            var typeface = GetTypefaceForFormatting(textInline.Formatting);
            var fontSize = textInline.Formatting.FontSizePt;
            var metrics = _textRenderer.GetFontMetrics(typeface, fontSize);
            var words = SplitIntoWords(textInline.Text);

            foreach (var word in words)
            {
                var wordWidth = _textRenderer.MeasureTextWithFallback(word, typeface, fontSize);

                if (currentLine.Count > 0 && currentLineWidth + wordWidth > currentLineLimit)
                {
                    CommitLine();
                    typeface = GetTypefaceForFormatting(textInline.Formatting);
                    metrics = _textRenderer.GetFontMetrics(typeface, fontSize);
                }

                currentLine.Add(new LayoutRun(word, typeface, fontSize, textInline.Formatting));
                currentLineWidth += wordWidth;
                currentMaxAscent = Math.Min(currentMaxAscent, metrics.Ascent);
                currentMaxDescent = Math.Max(currentMaxDescent, metrics.Descent);
                currentMaxLeading = Math.Max(currentMaxLeading, metrics.Leading);
            }
        }

        void ProcessTabInline(RunFormatting formatting, int inlineIndex, IReadOnlyList<DocxInlineElement> sourceElements, TabResolution? forcedResolution)
        {
            var resolution = forcedResolution ?? ResolveTabStop(paragraph, formatting, currentLineIndent, currentLineWidth, currentBarTabs);
            if (resolution == null)
                return;

            HandleTabAdvance(resolution);

            void HandleTabAdvance(TabResolution currentResolution)
            {
                var relativeTarget = currentResolution.TargetAbsolutePositionPt - currentLineIndent;

                if (currentResolution.Alignment == TabAlignment.Bar)
                {
                    currentBarTabs.Add(new BarTabInstruction(Math.Max(0f, relativeTarget), formatting));
                    return;
                }

                if (relativeTarget <= currentLineWidth + 0.1f)
                {
                    if (!currentResolution.FromDefault)
                    {
                        var fallback = ResolveDefaultTab(paragraph, currentLineIndent + currentLineWidth);
                        HandleTabAdvance(fallback);
                        return;
                    }

                    if (currentLine.Count > 0)
                    {
                        CommitLine();
                        HandleTabAdvance(currentResolution);
                        return;
                    }
                }

                var measurement = NeedsLookAhead(currentResolution.Alignment)
                    ? MeasureSegment(sourceElements, inlineIndex + 1)
                    : SegmentMeasurement.Empty;

                var desiredStart = ComputeDesiredStart(relativeTarget, currentResolution.Alignment, measurement);
                if (desiredStart < currentLineWidth - 0.1f && currentLine.Count > 0)
                {
                    CommitLine();
                    HandleTabAdvance(currentResolution);
                    return;
                }

                var projectedEnd = currentResolution.Alignment switch
                {
                    TabAlignment.Right or TabAlignment.Center or TabAlignment.Decimal => desiredStart + measurement.TotalWidth,
                    _ => desiredStart
                };

                if (projectedEnd > currentLineLimit && currentLine.Count > 0)
                {
                    CommitLine();
                    HandleTabAdvance(currentResolution);
                    return;
                }

                var spanWidth = Math.Max(0f, desiredStart - currentLineWidth);
                if (spanWidth > 0f || currentResolution.Leader != TabLeader.None)
                {
                    AppendLeaderOrPlaceholder(spanWidth, currentResolution.Leader, formatting);
                }

                TabDiagnostics.Write($"Tab align={currentResolution.Alignment} target={currentResolution.TargetAbsolutePositionPt:F1}pt span={spanWidth:F1} leader={currentResolution.Leader}");
            }
        }
    }

    private SKTypeface GetTypefaceForFormatting(RunFormatting formatting) =>
        _fontManager.GetTypeface(formatting.FontFamily, formatting.Bold, formatting.Italic);

    private sealed record TabResolution(float TargetAbsolutePositionPt, TabLeader Leader, TabAlignment Alignment, bool FromDefault);

    private readonly record struct SegmentMeasurement(float TotalWidth, float WidthBeforeDecimal, bool HasDecimal)
    {
        public static readonly SegmentMeasurement Empty = new(0f, 0f, false);
    }

    private TabResolution? ResolveTabStop(
        DocxParagraph paragraph,
        RunFormatting formatting,
        float currentLineIndent,
        float currentLineWidth,
        List<BarTabInstruction> barTabs)
    {
        var currentAbsolute = currentLineIndent + currentLineWidth;
        foreach (var stop in paragraph.ParagraphFormatting.TabStops)
        {
            if (stop.PositionPt <= currentAbsolute + 0.01f)
                continue;

            if (stop.Alignment == TabAlignment.Bar)
            {
                barTabs.Add(new BarTabInstruction(Math.Max(0f, stop.PositionPt - currentLineIndent), formatting));
                continue;
            }

            return new TabResolution(stop.PositionPt, stop.Leader, stop.Alignment, false);
        }

        return ResolveDefaultTab(paragraph, currentAbsolute);
    }

    private TabResolution ResolveDefaultTab(DocxParagraph paragraph, float currentAbsolute)
    {
        var interval = paragraph.DefaultTabStopPt <= 0f
            ? UnitConverter.DxaToPoints(720)
            : paragraph.DefaultTabStopPt;
        var multiplier = (float)Math.Floor(currentAbsolute / interval) + 1f;
        var target = multiplier * interval;
        return new TabResolution(target, TabLeader.None, TabAlignment.Left, true);
    }

    private string BuildLeaderString(float spanWidth, SKTypeface typeface, float fontSize, TabLeader leader)
    {
        var glyph = leader switch
        {
            TabLeader.Dots => ".",
            TabLeader.Dashes => "-",
            TabLeader.Underscore or TabLeader.Line or TabLeader.ThickLine or TabLeader.Heavy => "_",
            _ => string.Empty
        };

        if (string.IsNullOrEmpty(glyph))
            return string.Empty;

        var glyphWidth = _textRenderer.MeasureTextWithFallback(glyph, typeface, fontSize);
        if (glyphWidth <= 0f)
            return string.Empty;

        var repetitions = Math.Max(1, (int)Math.Ceiling(spanWidth / glyphWidth));
        return string.Concat(Enumerable.Repeat(glyph, repetitions));
    }

    private static int FindDecimalIndex(string text)
    {
        for (var i = 0; i < text.Length; i++)
        {
            if (text[i] == '.' || text[i] == ',')
                return i;
        }

        return -1;
    }

    private static List<string> SplitIntoWords(string text)
    {
        var words = new List<string>();
        var currentWord = new StringBuilder();

        foreach (var ch in text)
        {
            if (char.IsWhiteSpace(ch))
            {
                // Salva la parola corrente
                if (currentWord.Length > 0)
                {
                    words.Add(currentWord.ToString());
                    currentWord.Clear();
                }

                // Aggiungi lo spazio come token separato
                words.Add(ch.ToString());
            }
            else
            {
                currentWord.Append(ch);
            }
        }

        // Aggiungi l'ultima parola
        if (currentWord.Length > 0)
        {
            words.Add(currentWord.ToString());
        }

        return words;
    }
}

/// <summary>
/// Rappresenta una riga di testo dopo il layout.
/// </summary>
public sealed record LayoutLine(
    IReadOnlyList<LayoutRun> Runs,
    float WidthPt,
    float MaxAscent,
    float MaxDescent,
    float MaxLeading,
    bool IsFirstLine,
    float AvailableWidthPt,
    IReadOnlyList<BarTabInstruction> BarTabs
)
{
    /// <summary>
    /// Calcola l'altezza totale della riga (line spacing).
    /// Line spacing = descent - ascent + leading
    /// </summary>
    public float GetLineSpacing() => MaxDescent - MaxAscent + MaxLeading;
};

/// <summary>
/// Rappresenta un run di testo con font e dimensione per il rendering.
/// </summary>
public sealed record LayoutRun(
    string Text,
    SKTypeface Typeface,
    float FontSizePt,
    RunFormatting Formatting,
    bool IsDrawable = true,
    float AdvanceWidthOverride = 0f)
{
    public static LayoutRun Placeholder(float widthPt) =>
        new(string.Empty, FontManager.Instance.GetDefaultTypeface(), 1f, RunFormatting.Default, false, widthPt);
}

public sealed record BarTabInstruction(float RelativePositionPt, RunFormatting Formatting);
