using DocumentFormat.OpenXml.Wordprocessing;
using DocxToPdf.Sdk.Docx;
using DocxToPdf.Tests.Helpers;
using FluentAssertions;
using System.Linq;
using Xunit;

namespace DocxToPdf.Tests;

public sealed class NumberingResolverTests
{
    [Fact]
    public void ResolvesDecimalAndNestedNumbering()
    {
        using var stream = OpenXmlTestHelper.CreateNumberedDocumentStream();
        using var docx = DocxDocument.Open(stream);

        var paragraphs = docx.GetParagraphs().ToList();

        paragraphs[0].ListMarker.Should().NotBeNull();
        paragraphs[0].ListMarker!.Text.Should().Be("1.");
        paragraphs[1].ListMarker!.Text.Should().Be("1.a");
        paragraphs[1].ListMarker!.Suffix.Should().Be(DocumentFormat.OpenXml.Wordprocessing.LevelSuffixValues.Space);
        paragraphs[3].ListMarker!.Text.Should().Be("•");
        paragraphs[4].ListMarker!.Text.Should().Be("2.");

        var lvl1Formatting = paragraphs[1].ParagraphFormatting;
        lvl1Formatting.LeftIndentPt.Should().BeApproximately(DocxToPdf.Sdk.Units.UnitConverter.DxaToPoints(1080), 0.01f);
        lvl1Formatting.HangingIndentPt.Should().BeApproximately(DocxToPdf.Sdk.Units.UnitConverter.DxaToPoints(360), 0.01f);

        var lvl2Formatting = paragraphs[3].ParagraphFormatting;
        lvl2Formatting.LeftIndentPt.Should().BeApproximately(DocxToPdf.Sdk.Units.UnitConverter.DxaToPoints(1440), 0.01f);
        lvl2Formatting.HangingIndentPt.Should().BeApproximately(DocxToPdf.Sdk.Units.UnitConverter.DxaToPoints(360), 0.01f);
    }

    [Fact]
    public void RespectsStartOverrideAndContinuation()
    {
        using var stream = OpenXmlTestHelper.CreateNumberedDocumentStream();
        using var docx = DocxDocument.Open(stream);
        var paragraphs = docx.GetParagraphs().ToList();

        paragraphs[6].ListMarker!.Text.Should().Be("4.");
        paragraphs[7].ListMarker!.Text.Should().Be("5.");
    }
}
