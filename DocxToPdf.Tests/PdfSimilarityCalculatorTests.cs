using DocxToPdf.Tests.Helpers;
using DocxToPdf.Tests.Pdf;
using FluentAssertions;
using System;
using System.IO;
using Xunit;
using Xunit.Abstractions;

namespace DocxToPdf.Tests;

public sealed class PdfSimilarityCalculatorTests
{
    private readonly ITestOutputHelper _output;

    public PdfSimilarityCalculatorTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public void IdenticalPdfHasPerfectSimilarity()
    {
        var pdfPath = TestFiles.GetGoldenPath("lorem.pdf");
        var calculator = new PdfSimilarityCalculator();

        var result = calculator.Calculate(pdfPath, pdfPath);

        result.Pages.Should().NotBeEmpty();
        result.OverallSimilarity.Should().Be(1);
        result.Pages.Should().OnlyContain(p => p.Similarity == 1);
    }

    [Fact]
    public void CanCalculateSimilarityBetweenGoldenAndSdkOutput()
    {
        var docxPath = TestFiles.GetSamplePath("lorem.docx");
        var goldenPdf = TestFiles.GetGoldenPath("lorem.pdf");
        var calculator = new PdfSimilarityCalculator();

        var candidate = PdfTestHelper.RenderWithSdk(docxPath);
        try
        {
            var result = calculator.Calculate(goldenPdf, candidate);

            result.Pages.Should().NotBeEmpty();
            result.OverallSimilarity.Should().BeGreaterThanOrEqualTo(0);
            result.OverallSimilarity.Should().BeLessOrEqualTo(1);
            _output.WriteLine($"Overall similarity: {result.OverallSimilarity:P2}");
            foreach (var page in result.Pages)
            {
                _output.WriteLine($"Page {page.PageNumber}: {page.Similarity:P2}");
            }
        }
        finally
        {
            SafeDelete(candidate);
        }
    }

    private static void SafeDelete(string path)
    {
        try
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        catch
        {
            // ignore cleanup failures
        }
    }
}
