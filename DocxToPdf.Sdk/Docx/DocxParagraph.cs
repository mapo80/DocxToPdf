using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxToPdf.Sdk.Docx.Formatting;
using DocxToPdf.Sdk.Docx.Numbering;
using DocxToPdf.Sdk.Docx.Styles;
using DocxToPdf.Sdk.Units;
using System.Collections.Generic;
using System.Linq;

namespace DocxToPdf.Sdk.Docx;

/// <summary>
/// Paragrafo DOCX con testo e formattazioni risolte.
/// </summary>
public sealed record DocxParagraph
{
    public ParagraphFormatting ParagraphFormatting { get; init; } = ParagraphFormatting.Default;
    public IReadOnlyList<DocxRun> Runs { get; init; } = [];
    public IReadOnlyList<DocxInlineElement> InlineElements { get; init; } = Array.Empty<DocxInlineElement>();
    public DocxListMarker? ListMarker { get; init; }
    public float DefaultTabStopPt { get; init; }
    public char DecimalSymbol { get; init; } = '.';
    public string StyleId { get; init; } = string.Empty;

    internal static DocxParagraph FromParagraph(
        Paragraph paragraph,
        DocxStyleResolver styleResolver,
        NumberingResolver numberingResolver,
        float defaultTabStopPt,
        char decimalSymbol,
        bool isInTable = false)
    {
        var context = styleResolver.CreateParagraphContext(paragraph);
        var numberingResult = numberingResolver.Resolve(paragraph, context.ParagraphProperties, styleResolver);
        var runs = new List<DocxRun>();
        var inlineElements = new List<DocxInlineElement>();

        foreach (var run in paragraph.Elements<Run>())
        {
            var runFormatting = styleResolver.ResolveRunFormatting(context, run);
            if (DocxRun.FromRun(run, runFormatting) is { } docxRun)
            {
                runs.Add(docxRun);
            }
            inlineElements.AddRange(ExtractInlineElements(run, runFormatting));
        }

        // Marker di lista gestito a parte nel renderer; non va inserito nei run di testo.

        var paragraphFormatting = context.ParagraphFormatting;
        if (isInTable)
        {
            var lineSpacing = ParagraphLineSpacing.Auto(1f);
            paragraphFormatting = paragraphFormatting with
            {
                SpacingBeforePt = 0f,
                SpacingAfterPt = 0f,
                LineSpacing = lineSpacing
            };
        }

        return new DocxParagraph
        {
            ParagraphFormatting = paragraphFormatting,
            Runs = runs,
            InlineElements = inlineElements,
            ListMarker = numberingResult?.Marker,
            DefaultTabStopPt = defaultTabStopPt,
            DecimalSymbol = decimalSymbol,
            StyleId = context.ParagraphStyleId ?? string.Empty
        };
    }

    public string GetFullText() =>
        string.Concat(InlineElements.OfType<DocxTextInline>().Select(r => r.Text));

    private static IEnumerable<DocxInlineElement> ExtractInlineElements(Run run, RunFormatting formatting)
    {
        foreach (var child in run.ChildElements)
        {
            switch (child)
            {
                case DocumentFormat.OpenXml.Wordprocessing.Text text:
                    var textValue = text.Text ?? string.Empty;
                    if (!string.IsNullOrEmpty(textValue))
                        yield return new DocxTextInline(textValue, formatting);
                    break;
                case TabChar:
                    yield return new DocxTabInline(formatting);
                    break;
                case PositionalTab positionalTab:
                    if (TryGetPositionInPoints(positionalTab, out var positionPt))
                    {
                        yield return new DocxPositionalTabInline(
                            formatting,
                            positionPt,
                            MapAlignment(positionalTab.Alignment),
                            MapLeader(positionalTab.Leader),
                            MapReference(positionalTab.RelativeTo));
                    }
                    break;
            }
        }
    }

    private static bool TryGetPositionInPoints(PositionalTab positionalTab, out float positionPt)
    {
        const float emuToPoint = 1f / 12700f; // 1 EMU = 1/12700 pt
        foreach (var attribute in positionalTab.GetAttributes())
        {
            if (attribute.LocalName == "pos" && long.TryParse(attribute.Value, out var emu))
            {
                positionPt = emu * emuToPoint;
                return true;
            }
        }

        positionPt = 0f;
        return false;
    }

    private static TabAlignment MapAlignment(EnumValue<AbsolutePositionTabAlignmentValues>? alignment)
    {
        var value = alignment?.Value;
        if (value == null)
            return TabAlignment.Left;

        if (value == AbsolutePositionTabAlignmentValues.Center)
            return TabAlignment.Center;
        if (value == AbsolutePositionTabAlignmentValues.Right)
            return TabAlignment.Right;
        return TabAlignment.Left;
    }

    private static TabLeader MapLeader(EnumValue<AbsolutePositionTabLeaderCharValues>? leader)
    {
        var value = leader?.Value;
        if (value == null)
            return TabLeader.None;

        if (value == AbsolutePositionTabLeaderCharValues.Dot || value == AbsolutePositionTabLeaderCharValues.MiddleDot)
            return TabLeader.Dots;
        if (value == AbsolutePositionTabLeaderCharValues.Hyphen)
            return TabLeader.Dashes;
        if (value == AbsolutePositionTabLeaderCharValues.Underscore)
            return TabLeader.Underscore;
        return TabLeader.None;
    }

    private static PositionalTabReference MapReference(EnumValue<AbsolutePositionTabPositioningBaseValues>? reference)
    {
        var value = reference?.Value;
        if (value == null)
            return PositionalTabReference.Margin;

        if (value == AbsolutePositionTabPositioningBaseValues.Indent)
            return PositionalTabReference.Indent;

        return PositionalTabReference.Margin;
    }
}
