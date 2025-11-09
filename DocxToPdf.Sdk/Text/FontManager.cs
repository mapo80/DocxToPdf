using SkiaSharp;
using System;
using System.Collections.Generic;

namespace DocxToPdf.Sdk.Text;

/// <summary>
/// Gestisce i font e fornisce font fallback per codepoint non supportati.
/// In questa fase iniziale usa i font di sistema; fasi successive implementeranno
/// font embedding e fallback avanzato con SKFontManager.MatchCharacter.
/// </summary>
public sealed class FontManager
{
    private static readonly Lazy<FontManager> _instance = new(() => new FontManager());
    private readonly Dictionary<string, SKTypeface> _cache = new();

    public static FontManager Instance => _instance.Value;

    private FontManager()
    {
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

        var typeface = SKTypeface.FromFamilyName(familyName, style);
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

        // SkiaSharp 3.x API: MatchCharacter(int character)
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
}
