using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxToPdf.Sdk.Docx.Formatting;
using System;
using System.Collections.Generic;

namespace DocxToPdf.Sdk.Docx.Styles;

internal sealed class DocxStyleResolver
{
    private readonly Dictionary<string, StyleDefinition> _styles;
    private readonly ParagraphPropertySet _paragraphDefaults;
    private readonly RunPropertySet _runDefaults;
    private readonly ThemeFontScheme _themeFonts;
    private readonly ThemeColorPalette _themeColors;
    private readonly ColorSchemeMapper _colorMapper;
    private readonly string? _defaultParagraphStyleId;
    private readonly string? _defaultCharacterStyleId;

    private const string WordDefaultFontFamily = "Aptos";
    private const double WordDefaultFontSizePt = 12d;

    public string DefaultFontFamily { get; }
    public double DefaultFontSizePt { get; }
    public RgbColor DefaultTextColor { get; }

    private DocxStyleResolver(
        Dictionary<string, StyleDefinition> styles,
        ParagraphPropertySet paragraphDefaults,
        RunPropertySet runDefaults,
        ThemeFontScheme themeFonts,
        ThemeColorPalette themeColors,
        ColorSchemeMapper colorMapper,
        string? defaultParagraphStyleId,
        string? defaultCharacterStyleId)
    {
        _styles = styles;
        _paragraphDefaults = paragraphDefaults;
        _runDefaults = runDefaults;
        _themeFonts = themeFonts;
        _themeColors = themeColors;
        _colorMapper = colorMapper;
        _defaultParagraphStyleId = defaultParagraphStyleId;
        _defaultCharacterStyleId = defaultCharacterStyleId;

        DefaultFontFamily = DetermineDefaultFontFamily();
        DefaultFontSizePt = runDefaults.FontSizePt ?? WordDefaultFontSizePt;
        DefaultTextColor = ResolveThemeColor(ThemeColorValues.Text1, null, null);
    }

    public static DocxStyleResolver Load(WordprocessingDocument document)
    {
        var mainPart = document.MainDocumentPart
            ?? throw new InvalidOperationException("Documento DOCX non valido: manca MainDocumentPart.");

        var styles = new Dictionary<string, StyleDefinition>(StringComparer.OrdinalIgnoreCase);
        string? defaultParagraphStyleId = null;
        string? defaultCharacterStyleId = null;

        if (mainPart.StyleDefinitionsPart?.Styles is { } stylesRoot)
        {
            foreach (var style in stylesRoot.Elements<Style>())
            {
                var def = StyleDefinition.FromStyle(style);
                if (!string.IsNullOrEmpty(def.StyleId))
                {
                    styles[def.StyleId] = def;
                    if (def.StyleType == StyleValues.Paragraph && def.IsDefault)
                        defaultParagraphStyleId ??= def.StyleId;
                    if (def.StyleType == StyleValues.Character && def.IsDefault)
                        defaultCharacterStyleId ??= def.StyleId;
                }
            }
        }

        var docDefaults = mainPart.StyleDefinitionsPart?.Styles?.DocDefaults;
        var paragraphDefaultsElement = ExtractParagraphDefaults(docDefaults?.ParagraphPropertiesDefault);
        var paragraphDefaults = paragraphDefaultsElement != null
            ? ParagraphPropertySet.FromOpenXml(paragraphDefaultsElement)
            : ParagraphPropertySet.CreateWordDefaults();

        var runDefaultsElement = ExtractRunDefaults(docDefaults?.RunPropertiesDefault);
        var runDefaults = runDefaultsElement != null
            ? RunPropertySet.FromOpenXml(runDefaultsElement)
            : RunPropertySet.Empty;

        var themeFonts = ThemeFontScheme.Load(mainPart.ThemePart);
        var themeColors = ThemeColorPalette.Load(mainPart.ThemePart);

        var colorMapping = mainPart.DocumentSettingsPart?.Settings?.GetFirstChild<ColorSchemeMapping>();
        var colorMapper = ColorSchemeMapper.Load(colorMapping);

        return new DocxStyleResolver(
            styles,
            paragraphDefaults,
            runDefaults,
            themeFonts,
            themeColors,
            colorMapper,
            defaultParagraphStyleId,
            defaultCharacterStyleId);
    }

    public ParagraphStyleContext CreateParagraphContext(Paragraph paragraph)
    {
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value
            ?? _defaultParagraphStyleId
            ?? string.Empty;

        var paragraphProps = _paragraphDefaults.Clone();
        var runProps = _runDefaults.Clone();

        foreach (var style in EnumerateStyleChain(styleId))
        {
            if (style.StyleType != StyleValues.Paragraph)
                continue;

            paragraphProps.Apply(style.ParagraphProperties);
            runProps.Apply(style.RunProperties);
        }

        if (paragraph.ParagraphProperties != null)
        {
            var directProps = ParagraphPropertySet.FromOpenXml(paragraph.ParagraphProperties);
            paragraphProps.Apply(directProps);
        }

        if (paragraphProps.RunProperties != null)
        {
            runProps.Apply(paragraphProps.RunProperties);
        }

        return new ParagraphStyleContext(styleId, paragraphProps, runProps);
    }

    public RunFormatting ResolveRunFormatting(ParagraphStyleContext paragraphContext, Run run)
    {
        var runProps = paragraphContext.CreateRunPropertySet();
        var runStyleId = run.RunProperties?.RunStyle?.Val?.Value;

        if (!string.IsNullOrEmpty(runStyleId))
        {
            foreach (var style in EnumerateStyleChain(runStyleId))
            {
                if (style.StyleType != StyleValues.Character)
                    continue;

                runProps.Apply(style.RunProperties);
            }
        }
        else if (!string.IsNullOrEmpty(paragraphContext.ParagraphStyleId) &&
            _styles.TryGetValue(paragraphContext.ParagraphStyleId, out var paragraphStyle) &&
            !string.IsNullOrEmpty(paragraphStyle.LinkedStyleId))
        {
            foreach (var style in EnumerateStyleChain(paragraphStyle.LinkedStyleId))
            {
                if (style.StyleType != StyleValues.Character)
                    continue;
                runProps.Apply(style.RunProperties);
            }
        }

        if (run.RunProperties != null)
        {
            var direct = RunPropertySet.FromOpenXml(run.RunProperties);
            runProps.Apply(direct);
        }

        return runProps.ToFormatting(this);
    }

    internal string ResolveFontFamily(RunPropertySet set)
    {
        return set.AsciiFont
            ?? set.HighAnsiFont
            ?? set.EastAsiaFont
            ?? set.ComplexScriptFont
            ?? ResolveThemeFont(set.AsciiTheme)
            ?? ResolveThemeFont(set.HighAnsiTheme)
            ?? ResolveThemeFont(set.EastAsiaTheme)
            ?? ResolveThemeFont(set.ComplexScriptTheme)
            ?? DefaultFontFamily;
    }

    internal RgbColor ResolveColor(RunPropertySet set)
    {
        if (set.ThemeColor.HasValue)
            return ResolveThemeColor(set.ThemeColor.Value, set.Tint, set.Shade);

        if (!string.IsNullOrWhiteSpace(set.ColorHex))
            return RgbColor.FromHex(set.ColorHex);

        return DefaultTextColor;
    }

    private RgbColor ResolveThemeColor(ThemeColorValues requested, byte? tint, byte? shade)
    {
        var mapped = _colorMapper.Resolve(requested);
        var baseColor = TryGetThemeColor(mapped);
        var colored = ApplyTintShade(baseColor, tint, shade);
        return colored;
    }

    private RgbColor TryGetThemeColor(ThemeColorValues key)
    {
        if (_themeColors.TryGet(key, out var color))
            return color;

        if (key == ThemeColorValues.Dark1) return RgbColor.Black;
        if (key == ThemeColorValues.Light1) return RgbColor.White;
        if (key == ThemeColorValues.Dark2) return RgbColor.FromHex("222222");
        if (key == ThemeColorValues.Light2) return RgbColor.FromHex("DDDDDD");
        if (key == ThemeColorValues.Accent1) return RgbColor.FromHex("4F81BD");
        if (key == ThemeColorValues.Accent2) return RgbColor.FromHex("C0504D");
        if (key == ThemeColorValues.Accent3) return RgbColor.FromHex("9BBB59");
        if (key == ThemeColorValues.Accent4) return RgbColor.FromHex("8064A2");
        if (key == ThemeColorValues.Accent5) return RgbColor.FromHex("4BACC6");
        if (key == ThemeColorValues.Accent6) return RgbColor.FromHex("F79646");
        if (key == ThemeColorValues.Hyperlink) return RgbColor.FromHex("0000FF");
        if (key == ThemeColorValues.FollowedHyperlink) return RgbColor.FromHex("800080");
        return RgbColor.Black;
    }

    private static RgbColor ApplyTintShade(RgbColor baseColor, byte? tint, byte? shade)
    {
        var r = baseColor.R;
        var g = baseColor.G;
        var b = baseColor.B;

        if (tint.HasValue)
        {
            var factor = tint.Value / 255f;
            r = ApplyTintChannel(r, factor);
            g = ApplyTintChannel(g, factor);
            b = ApplyTintChannel(b, factor);
        }

        if (shade.HasValue)
        {
            var factor = shade.Value / 255f;
            r = ApplyShadeChannel(r, factor);
            g = ApplyShadeChannel(g, factor);
            b = ApplyShadeChannel(b, factor);
        }

        return new RgbColor((byte)r, (byte)g, (byte)b);
    }

    private static byte ApplyTintChannel(byte channel, float factor)
    {
        var value = channel + (255 - channel) * factor;
        return (byte)Math.Clamp(value, 0, 255);
    }

    private static byte ApplyShadeChannel(byte channel, float factor)
    {
        var value = channel * (1 - factor);
        return (byte)Math.Clamp(value, 0, 255);
    }

    private string DetermineDefaultFontFamily()
    {
        return _runDefaults.AsciiFont
            ?? _runDefaults.HighAnsiFont
            ?? ResolveThemeFont(_runDefaults.AsciiTheme)
            ?? ResolveThemeFont(_runDefaults.HighAnsiTheme)
            ?? ResolveThemeFont(_runDefaults.EastAsiaTheme)
            ?? ResolveThemeFont(_runDefaults.ComplexScriptTheme)
            ?? ResolveThemeFont(ThemeFontValues.MinorAscii)
            ?? WordDefaultFontFamily;
    }

    private string? ResolveThemeFont(ThemeFontValues? themeFont)
    {
        if (!themeFont.HasValue)
            return null;

        var value = themeFont.Value;
        if (value == ThemeFontValues.MajorAscii || value == ThemeFontValues.MajorHighAnsi)
            return _themeFonts.MajorLatin;
        if (value == ThemeFontValues.MajorEastAsia)
            return _themeFonts.MajorEastAsia ?? _themeFonts.MajorLatin;
        if (value == ThemeFontValues.MajorBidi)
            return _themeFonts.MajorComplex ?? _themeFonts.MajorLatin;
        if (value == ThemeFontValues.MinorAscii || value == ThemeFontValues.MinorHighAnsi)
            return _themeFonts.MinorLatin;
        if (value == ThemeFontValues.MinorEastAsia)
            return _themeFonts.MinorEastAsia ?? _themeFonts.MinorLatin;
        if (value == ThemeFontValues.MinorBidi)
            return _themeFonts.MinorComplex ?? _themeFonts.MinorLatin;
        return _themeFonts.MinorLatin;
    }

    private IEnumerable<StyleDefinition> EnumerateStyleChain(string? styleId)
    {
        if (string.IsNullOrEmpty(styleId))
            yield break;

        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var stack = new Stack<StyleDefinition>();

        var currentId = styleId;
        while (!string.IsNullOrEmpty(currentId) &&
            _styles.TryGetValue(currentId, out var style) &&
            visited.Add(currentId))
        {
            stack.Push(style);
            currentId = style.BasedOn;
        }

        while (stack.Count > 0)
        {
            yield return stack.Pop();
        }
    }

    private static ParagraphProperties? ExtractParagraphDefaults(ParagraphPropertiesDefault? defaultsElement)
    {
        if (defaultsElement == null)
            return null;

        var direct = defaultsElement.GetFirstChild<ParagraphProperties>();
        if (direct != null)
            return (ParagraphProperties)direct.CloneNode(true);

        var baseStyle = defaultsElement.GetFirstChild<ParagraphPropertiesBaseStyle>();
        if (baseStyle != null)
        {
            var clone = new ParagraphProperties();
            foreach (var child in baseStyle.ChildElements)
            {
                clone.Append(child.CloneNode(true));
            }
            return clone;
        }

        return null;
    }

    private static RunProperties? ExtractRunDefaults(RunPropertiesDefault? defaultsElement)
    {
        if (defaultsElement == null)
            return null;

        var direct = defaultsElement.GetFirstChild<RunProperties>();
        if (direct != null)
            return (RunProperties)direct.CloneNode(true);

        var baseStyle = defaultsElement.GetFirstChild<RunPropertiesBaseStyle>();
        if (baseStyle != null)
        {
            var clone = new RunProperties();
            foreach (var child in baseStyle.ChildElements)
            {
                clone.Append(child.CloneNode(true));
            }
            return clone;
        }

        return null;
    }
}
