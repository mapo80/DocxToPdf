using System.IO;
using System.Linq;
using DocxToPdf.Tests.Helpers;
using FluentAssertions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using Xunit;

namespace DocxToPdf.Tests.Pdf;

public sealed class PdfSelectableTextTests
{
    [Fact]
    public void LoremSample_ShouldExposeSelectableText()
    {
        var sampleDocx = TestFiles.GetSamplePath("lorem.docx");
        var pdfPath = PdfTestHelper.RenderWithSdk(sampleDocx);
        try
        {
            using var pdf = PdfDocument.Open(pdfPath);
            var page = pdf.GetPage(1);

            page.Letters.Should().NotBeEmpty("rendered PDF should contain logical letters instead of path-only glyphs");

            var extracted = string.Concat(page.Letters.Select(l => l.Value));
            extracted.Should().Contain("Lorem ipsum", "the sample DOCX text must remain searchable/selectable in the PDF output");

            var wordBoundingBoxes = page.GetWords().Select(w => w.BoundingBox).ToList();
            wordBoundingBoxes.Should().NotBeEmpty("PdfPig should be able to reconstruct words when ToUnicode maps are present");
        }
        finally
        {
            if (File.Exists(pdfPath))
                File.Delete(pdfPath);
        }
    }
}
