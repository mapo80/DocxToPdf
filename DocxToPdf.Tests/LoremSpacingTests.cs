using DocxToPdf.Sdk.Docx;
using DocxToPdf.Tests.Helpers;
using FluentAssertions;
using System.Linq;
using Xunit;

namespace DocxToPdf.Tests;

public sealed class LoremSpacingTests
{
    [Fact]
    public void ParagraphsInSampleInheritLineSpacing()
    {
        using var doc = DocxDocument.Open(TestFiles.GetSamplePath("lorem.docx"));
        foreach (var paragraph in doc.GetParagraphs())
        {
            paragraph.ParagraphFormatting.LineSpacing.Should().NotBeNull();
        }
    }

    [Fact]
    public void HeadingRunFontSizeMatchesExpected()
    {
        using var doc = DocxDocument.Open(TestFiles.GetSamplePath("lorem.docx"));
        var firstRun = doc.GetParagraphs().ElementAt(0).Runs[0];
        firstRun.Formatting.FontSizePt.Should().BeApproximately(24f, 0.01f);
        firstRun.Formatting.FontFamily.Should().Be("Cambria");
        firstRun.Formatting.CharacterSpacingPt.Should().Be(0f);
    }
}
