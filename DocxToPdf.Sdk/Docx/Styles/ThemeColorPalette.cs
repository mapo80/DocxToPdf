using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxToPdf.Sdk.Docx.Formatting;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace DocxToPdf.Sdk.Docx.Styles;

internal sealed class ThemeColorPalette
{
    private readonly Dictionary<ThemeColorValues, RgbColor> _colors = new();

    public void Set(ThemeColorValues key, RgbColor color) => _colors[key] = color;

    public bool TryGet(ThemeColorValues key, out RgbColor color) =>
        _colors.TryGetValue(key, out color);

    public static ThemeColorPalette Load(ThemePart? themePart)
    {
        var palette = new ThemeColorPalette();

        if (themePart == null)
        {
            ApplyOfficeDefaults(palette);
            return palette;
        }

        using var stream = themePart.GetStream(FileMode.Open, FileAccess.Read);
        var xdoc = XDocument.Load(stream);
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";

        var clrScheme = xdoc.Descendants(a + "clrScheme").FirstOrDefault();
        if (clrScheme == null)
        {
            ApplyOfficeDefaults(palette);
            return palette;
        }

        void Read(string name, ThemeColorValues key)
        {
            var elem = clrScheme.Element(a + name);
            if (elem == null)
                return;

            var srgb = elem.Element(a + "srgbClr")?.Attribute("val")?.Value;
            var sys = elem.Element(a + "sysClr")?.Attribute("lastClr")?.Value;
            var hex = srgb ?? sys;
            if (string.IsNullOrWhiteSpace(hex))
                return;

            palette.Set(key, RgbColor.FromHex(hex));
        }

        Read("dk1", ThemeColorValues.Dark1);
        Read("lt1", ThemeColorValues.Light1);
        Read("dk2", ThemeColorValues.Dark2);
        Read("lt2", ThemeColorValues.Light2);
        Read("accent1", ThemeColorValues.Accent1);
        Read("accent2", ThemeColorValues.Accent2);
        Read("accent3", ThemeColorValues.Accent3);
        Read("accent4", ThemeColorValues.Accent4);
        Read("accent5", ThemeColorValues.Accent5);
        Read("accent6", ThemeColorValues.Accent6);
        Read("hlink", ThemeColorValues.Hyperlink);
        Read("folHlink", ThemeColorValues.FollowedHyperlink);

        return palette;
    }

    private static void ApplyOfficeDefaults(ThemeColorPalette palette)
    {
        palette.Set(ThemeColorValues.Dark1, RgbColor.Black);
        palette.Set(ThemeColorValues.Light1, RgbColor.White);
        palette.Set(ThemeColorValues.Dark2, RgbColor.FromHex("1F497D"));
        palette.Set(ThemeColorValues.Light2, RgbColor.FromHex("EEECE1"));
        palette.Set(ThemeColorValues.Accent1, RgbColor.FromHex("4F81BD"));
        palette.Set(ThemeColorValues.Accent2, RgbColor.FromHex("C0504D"));
        palette.Set(ThemeColorValues.Accent3, RgbColor.FromHex("9BBB59"));
        palette.Set(ThemeColorValues.Accent4, RgbColor.FromHex("8064A2"));
        palette.Set(ThemeColorValues.Accent5, RgbColor.FromHex("4BACC6"));
        palette.Set(ThemeColorValues.Accent6, RgbColor.FromHex("F79646"));
        palette.Set(ThemeColorValues.Hyperlink, RgbColor.FromHex("0000FF"));
        palette.Set(ThemeColorValues.FollowedHyperlink, RgbColor.FromHex("800080"));
    }
}
