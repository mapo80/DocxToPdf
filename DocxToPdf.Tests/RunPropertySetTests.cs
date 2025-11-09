using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxToPdf.Sdk.Docx;
using DocxToPdf.Sdk.Docx.Formatting;
using DocxToPdf.Sdk.Docx.Styles;
using DocxToPdf.Tests.Helpers;
using FluentAssertions;
using System.Linq;
using Xunit;

namespace DocxToPdf.Tests;

public sealed class RunPropertySetTests
{
    [Fact]
    public void FromOpenXml_ParsesFontsSizesAndColors()
    {
        var props = new RunProperties(
            new RunFonts
            {
                Ascii = "Arial",
                HighAnsi = "Helvetica",
                ComplexScriptTheme = ThemeFontValues.MajorBidi
            },
            new FontSize { Val = "26" },
            new Bold { Val = OnOffValue.FromBoolean(true) },
            new Italic(),
            new Strike { Val = OnOffValue.FromBoolean(false) },
            new SmallCaps(),
            new Underline { Val = UnderlineValues.Single },
            new Color
            {
                ThemeColor = ThemeColorValues.Accent4,
                ThemeTint = "80",
                ThemeShade = "40",
                Val = "FF00FF"
            }
        );

        var set = RunPropertySet.FromOpenXml(props);
        set.AsciiFont.Should().Be("Arial");
        set.HighAnsiFont.Should().Be("Helvetica");
        set.ComplexScriptTheme.Should().Be(ThemeFontValues.MajorBidi);
        set.FontSizePt.Should().Be(13);
        set.Bold.Should().BeTrue();
        set.Italic.Should().BeTrue();
        set.Strike.Should().BeFalse();
        set.SmallCaps.Should().BeTrue();
        set.Underline.Should().BeTrue();
        set.ThemeColor.Should().Be(ThemeColorValues.Accent4);
        set.Tint.Should().Be(0x80);
        set.Shade.Should().Be(0x40);
        set.ColorHex.Should().Be("FF00FF");
    }

    [Fact]
    public void Apply_MergesValuesAndCloneProducesDeepCopy()
    {
        var baseSet = RunPropertySet.FromOpenXml(new RunProperties(
            new RunFonts { Ascii = "Calibri" },
            new FontSize { Val = "22" },
            new Bold()
        ));

        var overlay = RunPropertySet.FromOpenXml(new RunProperties(
            new RunFonts { Ascii = "Contoso" },
            new Color { Val = "ABCDEF" },
            new Italic()
        ));

        var cloned = baseSet.Clone();
        cloned.Apply(overlay);

        cloned.AsciiFont.Should().Be("Contoso");
        cloned.Bold.Should().BeTrue();
        cloned.Italic.Should().BeTrue();
        cloned.ColorHex.Should().Be("ABCDEF");

        baseSet.AsciiFont.Should().Be("Calibri");
        baseSet.ColorHex.Should().BeNull();
    }

    [Fact]
    public void ToFormatting_UsesResolverDefaultsWhenValuesMissing()
    {
        using var stream = OpenXmlTestHelper.CreateStyledDocumentStream();
        using var docx = DocxDocument.Open(stream);
        var paragraph = docx.GetParagraphs().Single();

        var formatting = paragraph.Runs[0].Formatting;
        formatting.FontFamily.Should().Be("Contoso Headings");
        formatting.FontSizePt.Should().Be(11);
        formatting.Color.Should().BeEquivalentTo(RgbColor.FromHex("1F497D"));
    }
}
