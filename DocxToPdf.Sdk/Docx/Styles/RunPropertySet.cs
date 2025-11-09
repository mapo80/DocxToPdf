using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxToPdf.Sdk.Docx.Formatting;
using DocxToPdf.Sdk.Units;
using System;

namespace DocxToPdf.Sdk.Docx.Styles;

/// <summary>
/// Rappresenta le proprietà di run ereditabili (prima della risoluzione finale).
/// </summary>
internal sealed class RunPropertySet
{
    internal string? AsciiFont { get; set; }
    internal string? HighAnsiFont { get; set; }
    internal string? ComplexScriptFont { get; set; }
    internal string? EastAsiaFont { get; set; }
    internal ThemeFontValues? AsciiTheme { get; set; }
    internal ThemeFontValues? HighAnsiTheme { get; set; }
    internal ThemeFontValues? ComplexScriptTheme { get; set; }
    internal ThemeFontValues? EastAsiaTheme { get; set; }
    internal double? FontSizePt { get; set; }
    internal bool? Bold { get; set; }
    internal bool? Italic { get; set; }
    internal bool? Underline { get; set; }
    internal bool? Strike { get; set; }
    internal bool? SmallCaps { get; set; }
    internal string? ColorHex { get; set; }
    internal ThemeColorValues? ThemeColor { get; set; }
    internal byte? Tint { get; set; }
    internal byte? Shade { get; set; }

    public RunPropertySet Clone() =>
        new()
        {
            AsciiFont = AsciiFont,
            HighAnsiFont = HighAnsiFont,
            ComplexScriptFont = ComplexScriptFont,
            EastAsiaFont = EastAsiaFont,
            AsciiTheme = AsciiTheme,
            HighAnsiTheme = HighAnsiTheme,
            ComplexScriptTheme = ComplexScriptTheme,
            EastAsiaTheme = EastAsiaTheme,
            FontSizePt = FontSizePt,
            Bold = Bold,
            Italic = Italic,
            Underline = Underline,
            Strike = Strike,
            SmallCaps = SmallCaps,
            ColorHex = ColorHex,
            ThemeColor = ThemeColor,
            Tint = Tint,
            Shade = Shade
        };

    public void Apply(RunPropertySet? overlay)
    {
        if (overlay == null)
            return;

        if (!string.IsNullOrEmpty(overlay.AsciiFont))
            AsciiFont = overlay.AsciiFont;
        if (!string.IsNullOrEmpty(overlay.HighAnsiFont))
            HighAnsiFont = overlay.HighAnsiFont;
        if (!string.IsNullOrEmpty(overlay.ComplexScriptFont))
            ComplexScriptFont = overlay.ComplexScriptFont;
        if (!string.IsNullOrEmpty(overlay.EastAsiaFont))
            EastAsiaFont = overlay.EastAsiaFont;

        if (overlay.AsciiTheme.HasValue)
            AsciiTheme = overlay.AsciiTheme;
        if (overlay.HighAnsiTheme.HasValue)
            HighAnsiTheme = overlay.HighAnsiTheme;
        if (overlay.ComplexScriptTheme.HasValue)
            ComplexScriptTheme = overlay.ComplexScriptTheme;
        if (overlay.EastAsiaTheme.HasValue)
            EastAsiaTheme = overlay.EastAsiaTheme;

        if (overlay.FontSizePt.HasValue)
            FontSizePt = overlay.FontSizePt;

        if (overlay.Bold.HasValue)
            Bold = overlay.Bold;
        if (overlay.Italic.HasValue)
            Italic = overlay.Italic;
        if (overlay.Underline.HasValue)
            Underline = overlay.Underline;
        if (overlay.Strike.HasValue)
            Strike = overlay.Strike;
        if (overlay.SmallCaps.HasValue)
            SmallCaps = overlay.SmallCaps;

        if (overlay.ColorHex != null)
            ColorHex = overlay.ColorHex;
        if (overlay.ThemeColor.HasValue)
            ThemeColor = overlay.ThemeColor;
        if (overlay.Tint.HasValue)
            Tint = overlay.Tint;
        if (overlay.Shade.HasValue)
            Shade = overlay.Shade;
    }

    public RunFormatting ToFormatting(DocxStyleResolver resolver)
    {
        var fontFamily = resolver.ResolveFontFamily(this);
        var fontSize = (float)(FontSizePt ?? resolver.DefaultFontSizePt);

        var color = resolver.ResolveColor(this);

        return new RunFormatting
        {
            FontFamily = fontFamily,
            FontSizePt = fontSize,
            Bold = Bold ?? false,
            Italic = Italic ?? false,
            Underline = Underline ?? false,
            Strike = Strike ?? false,
            SmallCaps = SmallCaps ?? false,
            Color = color
        };
    }

    public static RunPropertySet Empty => new();

    public static RunPropertySet FromOpenXml(RunProperties? runProps)
    {
        var set = new RunPropertySet();
        if (runProps == null)
            return set;

        if (runProps.RunFonts is { } runFonts)
        {
            set.AsciiFont = runFonts.Ascii?.Value;
            set.HighAnsiFont = runFonts.HighAnsi?.Value;
            set.ComplexScriptFont = runFonts.ComplexScript?.Value;
            set.EastAsiaFont = runFonts.EastAsia?.Value;
            set.AsciiTheme = runFonts.AsciiTheme?.Value;
            set.HighAnsiTheme = runFonts.HighAnsiTheme?.Value;
            set.ComplexScriptTheme = runFonts.ComplexScriptTheme?.Value;
            set.EastAsiaTheme = runFonts.EastAsiaTheme?.Value;
        }

        if (runProps.FontSize?.Val?.Value != null &&
            double.TryParse(runProps.FontSize.Val.Value, out var halfPoints))
        {
            set.FontSizePt = halfPoints / 2d;
        }
        else if (runProps.FontSizeComplexScript?.Val?.Value != null &&
            double.TryParse(runProps.FontSizeComplexScript.Val.Value, out var csHalfPoints))
        {
            set.FontSizePt = csHalfPoints / 2d;
        }

        set.Bold = ExtractOnOff(runProps.Bold);
        set.Italic = ExtractOnOff(runProps.Italic);
        set.Strike = ExtractOnOff(runProps.Strike);
        set.SmallCaps = ExtractOnOff(runProps.SmallCaps);

        if (runProps.Underline is { } underline)
        {
            var underlineVal = underline.Val?.Value ?? UnderlineValues.Single;
            set.Underline = underlineVal != UnderlineValues.None;
        }

        if (runProps.Color is { } color)
        {
            var val = color.Val?.Value;
            if (!string.IsNullOrWhiteSpace(val) &&
                !val.Equals("auto", StringComparison.OrdinalIgnoreCase))
            {
                set.ColorHex = val;
            }

            if (color.ThemeColor?.Value is ThemeColorValues themeColor)
                set.ThemeColor = themeColor;

            if (color.ThemeTint is { } tint)
                set.Tint = HexToByte(tint);
            if (color.ThemeShade is { } shade)
                set.Shade = HexToByte(shade);
        }

        return set;
    }

    private static bool? ExtractOnOff(OnOffType? val)
    {
        if (val == null)
            return null;

        var raw = val.Val;
        if (raw == null)
            return true;

        if (!raw.HasValue)
            return true;

        return raw.Value;
    }

    private static byte? HexToByte(StringValue? hex)
    {
        if (hex == null)
            return null;

        var value = hex.Value ?? string.Empty;
        if (value.Length == 0)
            return null;

        var cleaned = value.Trim();
        if (cleaned.Length == 0)
            return null;

        try
        {
            return Convert.ToByte(cleaned, 16);
        }
        catch
        {
            return null;
        }
    }
}
