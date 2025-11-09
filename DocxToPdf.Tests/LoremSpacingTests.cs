using DocxToPdf.Sdk.Docx;
using DocxToPdf.Tests.Helpers;
using FluentAssertions;
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
}
