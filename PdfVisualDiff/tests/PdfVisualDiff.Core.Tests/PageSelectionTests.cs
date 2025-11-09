using PdfVisualDiff.Core;
using Xunit;

namespace PdfVisualDiff.Core.Tests;

public sealed class PageSelectionTests
{
    [Theory]
    [InlineData(null, null)]
    [InlineData("", null)]
    [InlineData("1", "1")]
    [InlineData("1,2,3", "1-3")]
    [InlineData("3,1,2", "1-3")]
    [InlineData("1-3,5,7-9", "1-3,5,7-9")]
    public void NormalizeProducesCanonicalRanges(string? input, string? expected)
    {
        var normalized = PageSelection.Normalize(input);
        Assert.Equal(expected, normalized);
    }

    [Fact]
    public void ParseRejectsInvalidRange()
    {
        Assert.Throws<FormatException>(() => PageSelection.Parse("a-b"));
    }
}
