using DocumentFormat.OpenXml.Wordprocessing;
using DocxToPdf.Sdk.Docx.Formatting;
using DocxToPdf.Sdk.Docx.Styles;
using DocxToPdf.Tests.Helpers;
using FluentAssertions;
using Xunit;

namespace DocxToPdf.Tests;

public sealed class ThemeComponentsTests
{
    [Fact]
    public void ThemeFontSchemeLoadsMajorAndMinorFamilies()
    {
        var scheme = OpenXmlTestHelper.LoadThemeFontScheme("Major Font", "Minor Font");
        scheme.MajorLatin.Should().Be("Major Font");
        scheme.MinorLatin.Should().Be("Minor Font");
    }

    [Fact]
    public void ThemeColorPaletteReadsCustomColors()
    {
        var palette = OpenXmlTestHelper.LoadThemeColorPalette("112233");
        palette.TryGet(ThemeColorValues.Accent1, out var color).Should().BeTrue();
        color.ToHex().Should().Be("112233");
    }

    [Fact]
    public void ColorSchemeMapperRemapsColorsAccordingToSettings()
    {
        var mapping = new ColorSchemeMapping
        {
            Background1 = ColorSchemeIndexValues.Light2,
            Text1 = ColorSchemeIndexValues.Dark2,
            Accent1 = ColorSchemeIndexValues.Accent3
        };

        var mapper = ColorSchemeMapper.Load(mapping);

        mapper.Resolve(ThemeColorValues.Background1).Should().Be(ThemeColorValues.Light2);
        mapper.Resolve(ThemeColorValues.Text1).Should().Be(ThemeColorValues.Dark2);
        mapper.Resolve(ThemeColorValues.Accent1).Should().Be(ThemeColorValues.Accent3);
        mapper.Resolve(ThemeColorValues.Accent2).Should().Be(ThemeColorValues.Accent2);
    }

    [Fact]
    public void ThemePaletteFallsBackToOfficeDefaultsWhenThemeMissing()
    {
        var palette = ThemeColorPalette.Load(themePart: null);
        palette.TryGet(ThemeColorValues.Accent1, out var accent1).Should().BeTrue();
        accent1.ToHex().Should().Be("4F81BD");
        palette.TryGet(ThemeColorValues.Dark1, out var dark1).Should().BeTrue();
        dark1.ToHex().Should().Be("000000");
    }

    [Fact]
    public void ThemeFontSchemeFallsBackWhenThemeMissing()
    {
        var scheme = ThemeFontScheme.Load(themePart: null);
        scheme.MajorLatin.Should().Be("Calibri");
        scheme.MinorLatin.Should().Be("Calibri");
        scheme.MajorComplex.Should().Be("Times New Roman");
    }
}
