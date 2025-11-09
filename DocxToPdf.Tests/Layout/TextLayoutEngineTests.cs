using DocxToPdf.Sdk.Docx;
using DocxToPdf.Sdk.Docx.Formatting;
using DocxToPdf.Sdk.Layout;
using DocxToPdf.Sdk.Units;
using DocxToPdf.Sdk.Text;
using FluentAssertions;
using System;
using System.Linq;
using Xunit;

namespace DocxToPdf.Tests.Layout;

public sealed class TextLayoutEngineTests
{
    private static readonly RunFormatting DefaultFormatting = new()
    {
        FontFamily = "Arial",
        FontSizePt = 11f,
        Color = RgbColor.Black
    };

    [Fact]
    public void DefaultTabStopAdvancesCaretByHalfInch()
    {
        var paragraph = CreateParagraph([
            new DocxTabInline(DefaultFormatting),
            new DocxTextInline("Value", DefaultFormatting)
        ]);

        var engine = new TextLayoutEngine();
        var lines = engine.LayoutParagraph(paragraph, 400f);

        var line = Assert.Single(lines);
        var placeholder = Assert.Single(line.Runs, r => !r.IsDrawable);
        placeholder.AdvanceWidthOverride.Should().BeApproximately(UnitConverter.DxaToPoints(720), 0.1f);
        line.Runs.Should().Contain(r => r.Text == "Value");
    }

    [Fact]
    public void CustomTabStopOverridesDefault()
    {
        var customPosition = 200f;
        var paragraph = CreateParagraph([
            new DocxTabInline(DefaultFormatting),
            new DocxTextInline("Next", DefaultFormatting)
        ], new ParagraphFormatting
        {
            TabStops = new[] { new TabStopDefinition(customPosition, TabAlignment.Left, TabLeader.None) }
        });

        var engine = new TextLayoutEngine();
        var lines = engine.LayoutParagraph(paragraph, 400f);

        var line = Assert.Single(lines);
        var placeholder = Assert.Single(line.Runs, r => !r.IsDrawable);
        placeholder.AdvanceWidthOverride.Should().BeApproximately(customPosition, 0.1f);
    }

    [Fact]
    public void RightAlignedTabPositionsTextBeforeStop()
    {
        var target = UnitConverter.DxaToPoints(2880);
        var paragraph = CreateParagraph([
            new DocxTextInline("Label", DefaultFormatting),
            new DocxTabInline(DefaultFormatting),
            new DocxTextInline("Value", DefaultFormatting)
        ],
        new ParagraphFormatting
        {
            TabStops = new[] { new TabStopDefinition(target, TabAlignment.Right, TabLeader.None) }
        });

        var engine = new TextLayoutEngine();
        var lines = engine.LayoutParagraph(paragraph, 600f);

        var line = Assert.Single(lines);
        line.WidthPt.Should().BeApproximately(target, 0.5f);
    }

    [Fact]
    public void DecimalTabAlignsOnSeparator()
    {
        var target = UnitConverter.DxaToPoints(3600);
        var paragraph = CreateParagraph([
            new DocxTextInline("Cost", DefaultFormatting),
            new DocxTabInline(DefaultFormatting),
            new DocxTextInline("123.45", DefaultFormatting)
        ],
        new ParagraphFormatting
        {
            TabStops = new[] { new TabStopDefinition(target, TabAlignment.Decimal, TabLeader.None) }
        });

        var engine = new TextLayoutEngine();
        var lines = engine.LayoutParagraph(paragraph, 600f);
        var line = Assert.Single(lines);

        var renderer = new TextRenderer();
        float position = 0f;
        foreach (var run in line.Runs)
        {
            if (!run.IsDrawable)
            {
                position += run.AdvanceWidthOverride;
                continue;
            }

            var decimalIndex = run.Text.IndexOf('.');
            if (decimalIndex >= 0)
            {
                var prefix = run.Text[..decimalIndex];
                position += renderer.MeasureTextWithFallback(prefix, run.Typeface, run.FontSizePt);
                break;
            }

            position += renderer.MeasureTextWithFallback(run.Text, run.Typeface, run.FontSizePt);
        }

        position.Should().BeApproximately(target, 0.5f);
    }

    [Fact]
    public void PositionalTabUsesMarginReference()
    {
        var target = UnitConverter.DxaToPoints(4320);
        var paragraph = CreateParagraph([
            new DocxTextInline("Start", DefaultFormatting),
            new DocxPositionalTabInline(DefaultFormatting, target, TabAlignment.Left, TabLeader.None, PositionalTabReference.Margin),
            new DocxTextInline("Aligned", DefaultFormatting)
        ]);

        var engine = new TextLayoutEngine();
        var lines = engine.LayoutParagraph(paragraph, 600f);
        var line = Assert.Single(lines);
        var placeholder = line.Runs.FirstOrDefault(r => !r.IsDrawable);
        placeholder.Should().NotBeNull();

        var renderer = new TextRenderer();
        var startRun = line.Runs.First(r => r.Text == "Start");
        var startWidth = renderer.MeasureTextWithFallback(startRun.Text, startRun.Typeface, startRun.FontSizePt);
        placeholder!.AdvanceWidthOverride.Should().BeApproximately(target - startWidth, 1f);
    }

    private static DocxParagraph CreateParagraph(DocxInlineElement[] inlineElements, ParagraphFormatting? formattingOverride = null)
    {
        var formatting = formattingOverride ?? new ParagraphFormatting();
        return new DocxParagraph
        {
            ParagraphFormatting = formatting,
            Runs = new[]
            {
                new DocxRun
                {
                    Text = "placeholder",
                    Formatting = DefaultFormatting
                }
            },
            InlineElements = inlineElements,
            DefaultTabStopPt = UnitConverter.DxaToPoints(720)
        };
    }
}
