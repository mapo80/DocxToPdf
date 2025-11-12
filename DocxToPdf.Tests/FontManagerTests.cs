using DocxToPdf.Sdk.Text;
using FluentAssertions;
using Xunit;

namespace DocxToPdf.Tests;

public sealed class FontManagerTests
{
    [Fact]
    public void EmbeddedCaladeaFontIsAvailable()
    {
        var typeface = FontManager.Instance.GetTypeface("Caladea");
        typeface.Should().NotBeNull();
        typeface.FamilyName.Should().NotBeNull();
        typeface.FamilyName.Should().BeEquivalentTo("Caladea");
    }

    [Fact]
    public void CambriaTypefaceIsLoadedFromEmbeddedFonts()
    {
        var typeface = FontManager.Instance.GetTypeface("Cambria");
        typeface.Should().NotBeNull();
        typeface.FamilyName.Should().NotBeNull();
        typeface.FamilyName.Should().BeEquivalentTo("Cambria");
    }

    [Fact]
    public void CalibriTypefaceIsLoadedFromEmbeddedFonts()
    {
        var typeface = FontManager.Instance.GetTypeface("Calibri");
        typeface.Should().NotBeNull();
        typeface.FamilyName.Should().BeEquivalentTo("Calibri");
    }
}
