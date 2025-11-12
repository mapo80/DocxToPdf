using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace DocxToPdf.Sdk.Docx.Styles;

internal sealed class ThemeFontScheme
{
    public string? MajorLatin { get; private set; }
    public string? MajorEastAsia { get; private set; }
    public string? MajorComplex { get; private set; }
    public string? MinorLatin { get; private set; }
    public string? MinorEastAsia { get; private set; }
    public string? MinorComplex { get; private set; }

    public static ThemeFontScheme Load(ThemePart? themePart)
    {
        var scheme = new ThemeFontScheme
        {
            MajorLatin = "Aptos",
            MajorEastAsia = "Aptos",
            MajorComplex = "Times New Roman",
            MinorLatin = "Aptos",
            MinorEastAsia = "Aptos",
            MinorComplex = "Times New Roman"
        };

        if (themePart == null)
            return scheme;

        using var stream = themePart.GetStream(FileMode.Open, FileAccess.Read);
        var xdoc = XDocument.Load(stream);
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";

        var fontScheme = xdoc.Descendants(a + "fontScheme").FirstOrDefault();
        if (fontScheme == null)
            return scheme;

        var majorFont = fontScheme.Element(a + "majorFont");
        var minorFont = fontScheme.Element(a + "minorFont");

        scheme.MajorLatin = majorFont?.Element(a + "latin")?.Attribute("typeface")?.Value ?? scheme.MajorLatin;
        scheme.MajorEastAsia = majorFont?.Element(a + "ea")?.Attribute("typeface")?.Value ?? scheme.MajorEastAsia;
        scheme.MajorComplex = majorFont?.Element(a + "cs")?.Attribute("typeface")?.Value ?? scheme.MajorComplex;

        scheme.MinorLatin = minorFont?.Element(a + "latin")?.Attribute("typeface")?.Value ?? scheme.MinorLatin;
        scheme.MinorEastAsia = minorFont?.Element(a + "ea")?.Attribute("typeface")?.Value ?? scheme.MinorEastAsia;
        scheme.MinorComplex = minorFont?.Element(a + "cs")?.Attribute("typeface")?.Value ?? scheme.MinorComplex;

        return scheme;
    }
}
