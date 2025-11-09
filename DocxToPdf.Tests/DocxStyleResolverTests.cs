using DocxToPdf.Sdk.Docx;
using DocxToPdf.Sdk.Docx.Formatting;
using DocxToPdf.Tests.Helpers;
using FluentAssertions;
using System.IO;
using System.Linq;
using Xunit;

namespace DocxToPdf.Tests;

public sealed class DocxStyleResolverTests
{
    [Fact]
    public void ParagraphAndRunFormattingAreResolvedThroughCascade()
    {
        using var stream = OpenXmlTestHelper.CreateStyledDocumentStream();
        using var docx = DocxDocument.Open(stream);

        var paragraph = docx.GetParagraphs().Single();

        paragraph.ParagraphFormatting.Alignment.Should().Be(ParagraphAlignment.Center);
        paragraph.ParagraphFormatting.SpacingBeforePt.Should().BeApproximately(18f, 0.01f);
        paragraph.ParagraphFormatting.SpacingAfterPt.Should().BeApproximately(6f, 0.01f);
        paragraph.ParagraphFormatting.GetFirstLineOffsetPt().Should().BeApproximately(54f, 0.01f);
        paragraph.ParagraphFormatting.GetSubsequentLineOffsetPt().Should().BeApproximately(36f, 0.01f);

        paragraph.Runs.Should().HaveCount(4);
        paragraph.GetFullText().Should().Be("Plain linked Tint Hex Styled");

        var plain = paragraph.Runs[0].Formatting;
        plain.FontFamily.Should().Be("Contoso Headings");
        plain.Bold.Should().BeTrue();
        plain.Italic.Should().BeTrue(); // From linked character style
        plain.Color.ToHex().Should().Be("1F497D"); // Theme dark2 via clrSchemeMapping

        var tinted = paragraph.Runs[1].Formatting;
        tinted.Color.ToHex().Should().Be(ExpectedTintedHex("C0504D", 0x99));

        var hex = paragraph.Runs[2].Formatting;
        hex.Color.ToHex().Should().Be("3366FF");

        var styled = paragraph.Runs[3].Formatting;
        styled.Color.ToHex().Should().Be("4F81BD");
    }

    private static string ExpectedTintedHex(string baseHex, byte tint)
    {
        static byte Apply(byte channel, byte tintValue)
        {
            var factor = tintValue / 255f;
            var value = channel + (255 - channel) * factor;
            return (byte)Math.Clamp(value, 0f, 255f);
        }

        var rgb = System.Drawing.ColorTranslator.FromHtml("#" + baseHex);

        var r = Apply(rgb.R, tint);
        var g = Apply(rgb.G, tint);
        var b = Apply(rgb.B, tint);

        return $"{r:X2}{g:X2}{b:X2}";
    }
}
