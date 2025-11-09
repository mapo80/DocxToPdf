using DocxToPdf.Tests.Helpers;
using DocxToPdf.Tests.Pdf;
using FluentAssertions;
using System;
using System.Collections.Generic;
using System.IO;
using Xunit;
using Xunit.Abstractions;

namespace DocxToPdf.Tests;

public sealed class GoldenComparisonTests : IDisposable
{
    private readonly PdfSimilarityCalculator _calculator = new();
    private readonly List<string> _generatedFiles = new();
    private readonly ITestOutputHelper _output;

    public GoldenComparisonTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Theory]
    [InlineData("lorem.docx", "lorem.pdf")]
    [InlineData("styles-theme.docx", "styles-theme.pdf")]
    [InlineData("styles-theme-alt.docx", "styles-theme-alt.pdf")]
    [InlineData("numbering-multilevel.docx", "numbering-multilevel.pdf")]
    [InlineData("tabs-alignment.docx", "tabs-alignment.pdf")]
    [InlineData("bullets-basic.docx", "bullets-basic.pdf")]
    public void SampleSimilarityIsMeasured(string docxFile, string goldenPdf)
    {
        var docxPath = TestFiles.GetSamplePath(docxFile);
        var goldenPath = TestFiles.GetGoldenPath(goldenPdf);

        var candidatePath = PdfTestHelper.RenderWithSdk(docxPath);
        _generatedFiles.Add(candidatePath);

        var result = _calculator.Calculate(goldenPath, candidatePath, dpi: 200);

        _output.WriteLine($"[{docxFile}] similarity: {result.OverallSimilarity:P2}");
        result.Pages.Should().NotBeEmpty();
        result.OverallSimilarity.Should().BeGreaterOrEqualTo(0);
        result.OverallSimilarity.Should().BeLessOrEqualTo(1);
    }

    public void Dispose()
    {
        foreach (var path in _generatedFiles)
        {
            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch
            {
                // ignore cleanup
            }
        }
    }
}
