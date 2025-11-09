using DocxToPdf.Sdk.Text;
using FluentAssertions;
using SkiaSharp;
using Xunit;

namespace DocxToPdf.Tests;

    public sealed class TextRendererTests
{
    [Theory]
    [InlineData("Caladea", "™,")]
    [InlineData("Carlito", "™,")]
    [InlineData("Caladea", "€")]
    [InlineData("Caladea", "©")]
    public void MeasureTextWithFallback_ShouldReturnPositiveWidth_ForSpecialCharacters(string family, string text)
    {
        var renderer = new TextRenderer();
        var typeface = FontManager.Instance.GetTypeface(family);
        var width = renderer.MeasureTextWithFallback(text, typeface, 12f);
        width.Should().BeGreaterThan(0f, "the glyph '{0}' must be measured for layout decisions", text);
    }

    [Fact]
    public void DrawShapedTextWithLetterSpacingAddsAdditionalWidth()
    {
        var renderer = new TextRenderer();
        var typeface = FontManager.Instance.GetTypeface("Caladea");
        var text = "Test";
        using var surface = SKSurface.Create(new SKImageInfo(200, 200));
        using var canvas = surface.Canvas;
        var baseWidth = renderer.MeasureTextWithFallback(text, typeface, 20f);
        var spacedWidth = renderer.DrawShapedTextWithFallback(
            canvas,
            text,
            0,
            20,
            typeface,
            20f,
            SKColors.Black,
            0.5f);
        spacedWidth.Should().BeGreaterThan(baseWidth);
    }

    [Fact]
    public void KerningToggleAffectsMeasurement()
    {
        var renderer = new TextRenderer();
        var typeface = FontManager.Instance.GetTypeface("Caladea");
        var text = "To";
        var kerningOn = renderer.MeasureTextWithFallback(text, typeface, 28f, enableKerning: true);
        var kerningOff = renderer.MeasureTextWithFallback(text, typeface, 28f, enableKerning: false);
        kerningOff.Should().BeGreaterThan(kerningOn);
    }
}
