using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxToPdf.Sdk.Docx.Formatting;
using DocxToPdf.Sdk.Units;
using System;
using System.Collections.Generic;

namespace DocxToPdf.Sdk.Docx.Styles;

/// <summary>
/// Proprietà di paragrafo ereditabili (inclusi rPr condivisi).
/// </summary>
internal sealed class ParagraphPropertySet
{
    internal float? SpacingBeforePt { get; set; }
    internal float? SpacingAfterPt { get; set; }
    internal ParagraphLineSpacing? LineSpacing { get; set; }
    internal ParagraphAlignment? Alignment { get; set; }
    internal float? LeftIndentPt { get; set; }
    internal float? RightIndentPt { get; set; }
    internal float? FirstLineIndentPt { get; set; }
    internal float? HangingIndentPt { get; set; }
    internal RunPropertySet? RunProperties { get; set; }
    internal List<TabStopDefinition>? TabStops { get; set; }
    internal bool? SuppressSpacingBetweenSameStyle { get; set; }

    public ParagraphPropertySet Clone() =>
        new()
        {
            SpacingBeforePt = SpacingBeforePt,
            SpacingAfterPt = SpacingAfterPt,
            LineSpacing = LineSpacing,
            Alignment = Alignment,
            LeftIndentPt = LeftIndentPt,
            RightIndentPt = RightIndentPt,
            FirstLineIndentPt = FirstLineIndentPt,
            HangingIndentPt = HangingIndentPt,
            RunProperties = RunProperties?.Clone(),
            TabStops = TabStops != null ? new List<TabStopDefinition>(TabStops) : null,
            SuppressSpacingBetweenSameStyle = SuppressSpacingBetweenSameStyle
        };

    public void Apply(ParagraphPropertySet? overlay)
    {
        if (overlay == null)
            return;

        if (overlay.SpacingBeforePt.HasValue)
            SpacingBeforePt = overlay.SpacingBeforePt;
        if (overlay.SpacingAfterPt.HasValue)
            SpacingAfterPt = overlay.SpacingAfterPt;
        if (overlay.LineSpacing.HasValue)
            LineSpacing = overlay.LineSpacing;
        if (overlay.Alignment.HasValue)
            Alignment = overlay.Alignment;
        if (overlay.LeftIndentPt.HasValue)
            LeftIndentPt = overlay.LeftIndentPt;
        if (overlay.RightIndentPt.HasValue)
            RightIndentPt = overlay.RightIndentPt;
        if (overlay.FirstLineIndentPt.HasValue)
            FirstLineIndentPt = overlay.FirstLineIndentPt;
        if (overlay.HangingIndentPt.HasValue)
            HangingIndentPt = overlay.HangingIndentPt;

        if (overlay.RunProperties != null)
        {
            RunProperties ??= new RunPropertySet();
            RunProperties.Apply(overlay.RunProperties);
        }

        if (overlay.TabStops != null)
        {
            TabStops = new List<TabStopDefinition>(overlay.TabStops);
        }

        if (overlay.SuppressSpacingBetweenSameStyle.HasValue)
            SuppressSpacingBetweenSameStyle = overlay.SuppressSpacingBetweenSameStyle;
    }

    public ParagraphFormatting ToParagraphFormatting()
    {
        return new ParagraphFormatting
        {
            SpacingBeforePt = SpacingBeforePt ?? 0f,
            SpacingAfterPt = SpacingAfterPt ?? 0f,
            LineSpacing = LineSpacing,
            Alignment = Alignment ?? ParagraphAlignment.Left,
            LeftIndentPt = LeftIndentPt ?? 0f,
            RightIndentPt = RightIndentPt ?? 0f,
            FirstLineIndentPt = FirstLineIndentPt ?? 0f,
            HangingIndentPt = HangingIndentPt ?? 0f,
            TabStops = TabStops?.ToArray() ?? Array.Empty<TabStopDefinition>(),
            SuppressSpacingBetweenSameStyle = SuppressSpacingBetweenSameStyle ?? false
        };
    }

    public static ParagraphPropertySet Empty => new();

    public static ParagraphPropertySet CreateWordDefaults() =>
        new()
        {
            SpacingBeforePt = 0f,
            SpacingAfterPt = UnitConverter.DxaToPoints(160), // Word default: 8 pt after paragraph
            LineSpacing = ParagraphLineSpacing.Auto(1.15f)
        };

    public static ParagraphPropertySet FromOpenXml(OpenXmlElement? paragraphPropertiesElement)
    {
        var set = new ParagraphPropertySet();
        if (paragraphPropertiesElement == null)
            return set;

        if (paragraphPropertiesElement.GetFirstChild<SpacingBetweenLines>() is { } spacing)
        {
            if (spacing.Before?.Value is string before && int.TryParse(before, out var beforeTwips))
                set.SpacingBeforePt = UnitConverter.DxaToPoints(beforeTwips);
            if (spacing.After?.Value is string after && int.TryParse(after, out var afterTwips))
                set.SpacingAfterPt = UnitConverter.DxaToPoints(afterTwips);
            if (spacing.Line?.Value is string line && int.TryParse(line, out var lineValue))
            {
                var rule = spacing.LineRule?.Value;
                if (rule == LineSpacingRuleValues.Exact)
                {
                    set.LineSpacing = ParagraphLineSpacing.Exact(UnitConverter.DxaToPoints(lineValue));
                }
                else if (rule == LineSpacingRuleValues.AtLeast)
                {
                    set.LineSpacing = ParagraphLineSpacing.AtLeast(UnitConverter.DxaToPoints(lineValue));
                }
                else
                {
                    var multiple = Math.Max(0.1f, lineValue / 240f);
                    set.LineSpacing = ParagraphLineSpacing.Auto(multiple);
                }
            }
        }

        if (paragraphPropertiesElement.GetFirstChild<Indentation>() is { } indent)
        {
            if (indent.Left?.Value is string left && int.TryParse(left, out var leftDxa))
                set.LeftIndentPt = UnitConverter.DxaToPoints(leftDxa);
            if (indent.Right?.Value is string right && int.TryParse(right, out var rightDxa))
                set.RightIndentPt = UnitConverter.DxaToPoints(rightDxa);
            if (indent.FirstLine?.Value is string first && int.TryParse(first, out var firstDxa))
                set.FirstLineIndentPt = UnitConverter.DxaToPoints(firstDxa);
            if (indent.Hanging?.Value is string hanging && int.TryParse(hanging, out var hangingDxa))
                set.HangingIndentPt = UnitConverter.DxaToPoints(hangingDxa);
        }

        if (paragraphPropertiesElement.GetFirstChild<ContextualSpacing>() != null)
        {
            set.SuppressSpacingBetweenSameStyle = true;
        }

        if (paragraphPropertiesElement.GetFirstChild<Justification>() is { } jc)
        {
            if (jc.Val?.Value == JustificationValues.Center)
                set.Alignment = ParagraphAlignment.Center;
            else if (jc.Val?.Value == JustificationValues.Right)
                set.Alignment = ParagraphAlignment.Right;
            else if (jc.Val?.Value == JustificationValues.Both)
                set.Alignment = ParagraphAlignment.Justified;
            else if (jc.Val?.Value == JustificationValues.Distribute)
                set.Alignment = ParagraphAlignment.Distributed;
            else
                set.Alignment = ParagraphAlignment.Left;
        }

        var paraRunProps = RunPropertyHelpers.CloneRunProperties(
            paragraphPropertiesElement.GetFirstChild<ParagraphMarkRunProperties>());
        if (paraRunProps != null)
        {
            set.RunProperties = RunPropertySet.FromOpenXml(paraRunProps);
        }
        else if (paragraphPropertiesElement.GetFirstChild<RunProperties>() is { } directRunProps)
        {
            set.RunProperties = RunPropertySet.FromOpenXml((RunProperties)directRunProps.CloneNode(true));
        }

        if (paragraphPropertiesElement.GetFirstChild<Tabs>() is { } tabs)
        {
            set.TabStops = ParseTabStops(tabs);
        }

        return set;
    }

    private static List<TabStopDefinition> ParseTabStops(Tabs tabs)
    {
        var list = new List<TabStopDefinition>();
        foreach (var tab in tabs.Elements<TabStop>())
        {
            var posValue = tab.Position?.Value;
            if (!posValue.HasValue)
                continue;

            var positionPt = UnitConverter.DxaToPoints(posValue.Value);

            var alignment = TabAlignment.Left;
            var alignmentValue = tab.Val?.Value;
            if (alignmentValue.HasValue)
            {
                if (alignmentValue.Value == TabStopValues.Center)
                    alignment = TabAlignment.Center;
                else if (alignmentValue.Value == TabStopValues.Right)
                    alignment = TabAlignment.Right;
                else if (alignmentValue.Value == TabStopValues.Decimal)
                    alignment = TabAlignment.Decimal;
                else if (alignmentValue.Value == TabStopValues.Bar)
                    alignment = TabAlignment.Bar;
            }

            var leader = TabLeader.None;
            var leaderValue = tab.Leader?.Value;
            if (leaderValue.HasValue)
            {
                var val = leaderValue.Value;
                if (val == TabStopLeaderCharValues.Dot || val == TabStopLeaderCharValues.MiddleDot)
                    leader = TabLeader.Dots;
                else if (val == TabStopLeaderCharValues.Hyphen)
                    leader = TabLeader.Dashes;
                else if (val == TabStopLeaderCharValues.Heavy)
                    leader = TabLeader.Heavy;
                else if (val == TabStopLeaderCharValues.Underscore)
                    leader = TabLeader.Underscore;
            }

            list.Add(new TabStopDefinition(positionPt, alignment, leader));
        }

        list.Sort((a, b) => a.PositionPt.CompareTo(b.PositionPt));
        return list;
    }
}
