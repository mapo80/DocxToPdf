using PdfVisualDiff.Core;
using Xunit;

namespace PdfVisualDiff.Core.Tests;

public sealed class GeometryDiffAnalyzerTests
{
    [Fact]
    public void CompareWordsDetectsTextMismatch()
    {
        var expected = new[]
        {
            new WordInfo("Hello", new GeometryBox(0, 0, 10, 5)),
            new WordInfo("World", new GeometryBox(15, 0, 12, 5))
        };

        var actual = new[]
        {
            new WordInfo("Hello", new GeometryBox(0.1, 0, 10, 5)),
            new WordInfo("Word", new GeometryBox(15.5, 0, 12, 5))
        };

        var mismatches = GeometryDiffAnalyzer.CompareWords(expected, actual, 0.25);
        Assert.Single(mismatches);
        Assert.Equal("World", mismatches[0].ExpectedText);
        Assert.Equal("Word", mismatches[0].ActualText);
    }

    [Fact]
    public void CompareWordsDetectsGeometryDelta()
    {
        var expected = new[]
        {
            new WordInfo("Hello", new GeometryBox(0, 0, 10, 5))
        };
        var actual = new[]
        {
            new WordInfo("Hello", new GeometryBox(1, 0, 10, 5))
        };

        var mismatches = GeometryDiffAnalyzer.CompareWords(expected, actual, 0.25);
        Assert.Single(mismatches);
        Assert.True(mismatches[0].Delta.MaxDeviation > 0.25);
    }
}
