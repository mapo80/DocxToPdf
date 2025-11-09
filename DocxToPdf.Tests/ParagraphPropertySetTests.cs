using DocumentFormat.OpenXml.Wordprocessing;
using DocxToPdf.Sdk.Docx.Formatting;
using DocxToPdf.Sdk.Docx.Styles;
using FluentAssertions;
using Xunit;

namespace DocxToPdf.Tests;

public sealed class ParagraphPropertySetTests
{
    [Fact]
    public void FromOpenXml_ConvertsSpacingIndentAndRunProperties()
    {
        var props = new ParagraphProperties(
            new SpacingBetweenLines { Before = "240", After = "360", Line = "480" },
            new Indentation { Left = "720", Right = "360", FirstLine = "180", Hanging = "0" },
            new Justification { Val = JustificationValues.Right },
            new ParagraphMarkRunProperties(
                new RunFonts { Ascii = "Calibri" },
                new Bold()
            )
        );

        var set = ParagraphPropertySet.FromOpenXml(props);
        set.SpacingBeforePt.Should().Be(12);
        set.SpacingAfterPt.Should().Be(18);
        set.LineSpacingPt.Should().Be(24);
        set.LeftIndentPt.Should().Be(36);
        set.RightIndentPt.Should().Be(18);
        set.FirstLineIndentPt.Should().Be(9);
        set.Alignment.Should().Be(ParagraphAlignment.Right);
        set.RunProperties.Should().NotBeNull();
        set.RunProperties!.AsciiFont.Should().Be("Calibri");
        set.RunProperties.Bold.Should().BeTrue();
    }

    [Fact]
    public void Apply_MergesRunProperties()
    {
        var baseSet = new ParagraphPropertySet
        {
            SpacingBeforePt = 6,
            RunProperties = RunPropertySet.FromOpenXml(new RunProperties(new RunFonts { Ascii = "Base" }))
        };

        var overlay = new ParagraphPropertySet
        {
            SpacingAfterPt = 8,
            RunProperties = RunPropertySet.FromOpenXml(new RunProperties(new RunFonts { Ascii = "Overlay" }, new Italic()))
        };

        baseSet.Apply(overlay);

        baseSet.SpacingBeforePt.Should().Be(6);
        baseSet.SpacingAfterPt.Should().Be(8);
        baseSet.RunProperties.Should().NotBeNull();
        baseSet.RunProperties!.AsciiFont.Should().Be("Overlay");
        baseSet.RunProperties.Italic.Should().BeTrue();
    }

    [Fact]
    public void ParagraphFormattingComputesOffsetsForHangingIndent()
    {
        var formatting = new ParagraphFormatting
        {
            LeftIndentPt = 24,
            HangingIndentPt = 12,
            FirstLineIndentPt = 6
        };

        formatting.GetFirstLineOffsetPt().Should().Be(12);
        formatting.GetSubsequentLineOffsetPt().Should().Be(24);
    }
}
