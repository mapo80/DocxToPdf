using SkiaSharp;
using System;
using System.Collections.Generic;
using System.IO;

namespace DocxToPdf.Sdk.Text;

/// <summary>
/// Gestisce i font e fornisce font fallback per codepoint non supportati.
/// In questa fase iniziale usa i font di sistema; fasi successive implementeranno
/// font embedding e fallback avanzato con SKFontManager.MatchCharacter.
/// </summary>
public sealed class FontManager
{
    private static readonly Lazy<FontManager> _instance = new(() => new FontManager());
    private readonly Dictionary<string, SKTypeface> _cache = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, LocalFontFamily> _embeddedFamilies;
    private readonly Dictionary<string, string> _fontAliases;

    public static FontManager Instance => _instance.Value;

    private FontManager()
    {
        _embeddedFamilies = LoadEmbeddedFamilies();
        _fontAliases = BuildAliases();
    }

    /// <summary>
    /// Ottiene un typeface per famiglia e stile.
    /// Usa i font di sistema disponibili.
    /// </summary>
    public SKTypeface GetTypeface(
        string familyName = "Arial",
        SKFontStyle? style = null)
    {
        style ??= SKFontStyle.Normal;
        var key = $"{familyName}|{style.Weight}|{style.Slant}|{style.Width}";

        if (_cache.TryGetValue(key, out var cached))
            return cached;

        var typeface = CreateTypeface(familyName, style);
        _cache[key] = typeface;
        return typeface;
    }

    /// <summary>
    /// Ottiene il typeface di default (Arial/Helvetica o equivalente sistema).
    /// </summary>
    public SKTypeface GetDefaultTypeface() => GetTypeface();

    public SKTypeface GetTypeface(string familyName, bool bold, bool italic)
    {
        var weight = bold ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal;
        var slant = italic ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright;
        var style = new SKFontStyle(weight, SKFontStyleWidth.Normal, slant);
        return GetTypeface(familyName, style);
    }

    /// <summary>
    /// Trova un font fallback che supporta il codepoint specificato.
    /// Usa SKFontManager.MatchCharacter per cercare nei font di sistema.
    /// </summary>
    /// <param name="codepoint">Unicode codepoint da cercare</param>
    /// <param name="primaryFont">Font primario (usato per estrarre stile preferito)</param>
    /// <returns>Typeface che supporta il codepoint, o null se non trovato</returns>
    public SKTypeface? FindFallbackTypeface(int codepoint, SKTypeface primaryFont)
    {
        var fontManager = SKFontManager.Default;

        var fallback = fontManager.MatchCharacter(codepoint);

        return fallback;
    }

    /// <summary>
    /// Verifica se un typeface contiene il glifo per un codepoint specifico.
    /// </summary>
    public bool TypefaceContainsGlyph(SKTypeface typeface, int codepoint)
    {
        // Converti codepoint in string UTF-16
        var text = char.ConvertFromUtf32(codepoint);

        // Usa GetGlyphs per verificare se il typeface supporta il carattere
        using var font = new SKFont(typeface);
        var glyphs = new ushort[text.Length];
        font.GetGlyphs(text, glyphs);

        // Glyph ID 0 indica "missing glyph" (tofu)
        // Se tutti i glifi sono 0, il carattere non è supportato
        foreach (var glyph in glyphs)
        {
            if (glyph != 0)
                return true; // Almeno un glifo valido trovato
        }

        return false;
    }
    private SKTypeface CreateTypeface(string familyName, SKFontStyle style)
    {
        if (TryGetEmbeddedTypeface(familyName, style, out var typeface))
            return typeface;

        if (TryGetSystemTypeface(familyName, style, out typeface))
            return typeface;

        if (_fontAliases.TryGetValue(familyName, out var alias))
        {
            if (TryGetEmbeddedTypeface(alias, style, out typeface))
                return typeface;
            if (TryGetSystemTypeface(alias, style, out typeface))
                return typeface;
        }

        typeface = SKTypeface.FromFamilyName(familyName, style);
        return typeface ?? SKTypeface.Default;
    }

    private bool TryGetSystemTypeface(string familyName, SKFontStyle style, out SKTypeface typeface)
    {
        typeface = SKTypeface.FromFamilyName(familyName, style);
        if (typeface == null)
            return false;

        if (string.Equals(typeface.FamilyName, familyName, StringComparison.OrdinalIgnoreCase))
            return true;

        typeface.Dispose();
        typeface = null!;
        return false;
    }

    private bool TryGetEmbeddedTypeface(string familyName, SKFontStyle style, out SKTypeface typeface)
    {
        typeface = null!;
        if (!_embeddedFamilies.TryGetValue(familyName, out var family))
            return false;

        var path = family.Resolve(style);
        if (path == null || !File.Exists(path))
            return false;

        typeface = SKTypeface.FromFile(path);
        return typeface != null;
    }

    private Dictionary<string, LocalFontFamily> LoadEmbeddedFamilies()
    {
        var dict = new Dictionary<string, LocalFontFamily>(StringComparer.OrdinalIgnoreCase);
        var baseDir = AppContext.BaseDirectory;
        var fontsDir = Path.Combine(baseDir, "Fonts");
        if (!Directory.Exists(fontsDir))
            return dict;

        void TryAdd(string familyName, string regular, string bold, string italic, string boldItalic)
        {
            var family = LocalFontFamily.Create(fontsDir, regular, bold, italic, boldItalic);
            if (family != null)
                dict[familyName] = family;
        }

        TryAdd("Caladea", "Caladea-Regular.ttf", "Caladea-Bold.ttf", "Caladea-Italic.ttf", "Caladea-BoldItalic.ttf");
        TryAdd("Carlito", "Carlito-Regular.ttf", "Carlito-Bold.ttf", "Carlito-Italic.ttf", "Carlito-BoldItalic.ttf");

        return dict;
    }

    private static Dictionary<string, string> BuildAliases() =>
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["Cambria"] = "Caladea",
            ["Cambria Math"] = "Caladea",
            ["Calibri"] = "Carlito",
            ["Calibri Light"] = "Carlito"
        };

    private sealed record LocalFontFamily(string RegularPath, string? BoldPath, string? ItalicPath, string? BoldItalicPath)
    {
        public string? Resolve(SKFontStyle style)
        {
            var isBold = style.Weight >= (int)SKFontStyleWeight.SemiBold;
            var isItalic = style.Slant == SKFontStyleSlant.Italic || style.Slant == SKFontStyleSlant.Oblique;

            if (isBold && isItalic && !string.IsNullOrEmpty(BoldItalicPath))
                return BoldItalicPath;
            if (isBold && !string.IsNullOrEmpty(BoldPath))
                return BoldPath;
            if (isItalic && !string.IsNullOrEmpty(ItalicPath))
                return ItalicPath;
            return RegularPath;
        }

        public static LocalFontFamily? Create(string fontsDir, string regular, string bold, string italic, string boldItalic)
        {
            var regularPath = Path.Combine(fontsDir, regular);
            if (!File.Exists(regularPath))
                return null;

            string? GetOptional(string relativePath)
            {
                var path = Path.Combine(fontsDir, relativePath);
                return File.Exists(path) ? path : null;
            }

            return new LocalFontFamily(
                regularPath,
                GetOptional(bold),
                GetOptional(italic),
                GetOptional(boldItalic));
        }
    }
}
