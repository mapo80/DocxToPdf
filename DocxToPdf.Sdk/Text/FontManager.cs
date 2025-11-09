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
            LogFontInfo($"Falling back '{familyName}' to alias '{alias}'");
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

        var source = family.Resolve(style);
        if (source == null)
            return false;

        typeface = source.CreateTypeface();
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
            try
            {
                var family = LocalFontFamily.Create(fontsDir, regular, bold, italic, boldItalic);
                if (family != null)
                    dict[familyName] = family;
            }
            catch (Exception ex)
            {
                LogFontWarning($"Failed to load embedded font '{familyName}' from {fontsDir}: {ex.Message}");
            }
        }

        TryAdd("Caladea", "Caladea-Regular.ttf", "Caladea-Bold.ttf", "Caladea-Italic.ttf", "Caladea-BoldItalic.ttf");
        TryAdd("Carlito", "Carlito-Regular.ttf", "Carlito-Bold.ttf", "Carlito-Italic.ttf", "Carlito-BoldItalic.ttf");

        var wordFontsDir = "/Applications/Microsoft Word.app/Contents/Resources/DFonts";
        if (Directory.Exists(wordFontsDir))
        {
            void TryAddWordFont(string familyName, string regular, string bold, string italic, string boldItalic)
            {
                try
                {
                    var family = LocalFontFamily.CreateFromPaths(
                        Path.Combine(wordFontsDir, regular),
                        Path.Combine(wordFontsDir, bold),
                        Path.Combine(wordFontsDir, italic),
                        Path.Combine(wordFontsDir, boldItalic));
                    if (family != null)
                    {
                        dict[familyName] = family;
                        LogFontInfo($"Loaded Word font '{familyName}'");
                    }
                    else
                    {
                        LogFontWarning($"Font '{familyName}' not found in Word resources.");
                    }
                }
                catch (Exception ex)
                {
                    LogFontWarning($"Failed to load Word font '{familyName}': {ex.Message}");
                }
            }

            TryAddWordFont("Aptos", "Aptos.ttf", "Aptos-Bold.ttf", "Aptos-Italic.ttf", "Aptos-Bold-Italic.ttf");
            TryAddWordFont("Cambria", "Cambria.ttc", "Cambriab.ttf", "Cambriai.ttf", "Cambriaz.ttf");
            TryAddWordFont("Calibri", "Calibri.ttf", "Calibrib.ttf", "Calibrii.ttf", "Calibriz.ttf");
        }

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

    private sealed record LocalFontFamily(FontSource Regular, FontSource? Bold, FontSource? Italic, FontSource? BoldItalic)
    {
        public FontSource? Resolve(SKFontStyle style)
        {
            var isBold = style.Weight >= (int)SKFontStyleWeight.SemiBold;
            var isItalic = style.Slant == SKFontStyleSlant.Italic || style.Slant == SKFontStyleSlant.Oblique;

            if (isBold && isItalic && BoldItalic != null)
                return BoldItalic;
            if (isBold && Bold != null)
                return Bold;
            if (isItalic && Italic != null)
                return Italic;
            return Regular;
        }

        public static LocalFontFamily? Create(string fontsDir, string regular, string bold, string italic, string boldItalic)
        {
            var regularPath = Path.Combine(fontsDir, regular);
            if (!File.Exists(regularPath))
                return null;

            var regularSource = FontSource.Create(regularPath);
            if (regularSource == null)
                return null;

            string? GetOptional(string relativePath)
            {
                var path = Path.Combine(fontsDir, relativePath);
                return File.Exists(path) ? path : null;
            }

            return new LocalFontFamily(
                regularSource,
                FontSource.Create(GetOptional(bold)),
                FontSource.Create(GetOptional(italic)),
                FontSource.Create(GetOptional(boldItalic)));
        }

        public static LocalFontFamily? CreateFromPaths(string regularPath, string? boldPath, string? italicPath, string? boldItalicPath)
        {
            if (string.IsNullOrEmpty(regularPath) || !File.Exists(regularPath))
                return null;

            var regularSource = FontSource.Create(regularPath);
            if (regularSource == null)
                return null;

            string? Normalize(string? path) => string.IsNullOrEmpty(path) || !File.Exists(path) ? null : path;

            return new LocalFontFamily(
                regularSource,
                FontSource.Create(Normalize(boldPath)),
                FontSource.Create(Normalize(italicPath)),
                FontSource.Create(Normalize(boldItalicPath)));
        }
    }

    private sealed record FontSource(string Path, int? FaceIndex = null)
    {
        public static FontSource? Create(string? path)
        {
            if (string.IsNullOrEmpty(path))
                return null;

            if (System.IO.Path.GetExtension(path).Equals(".ttc", StringComparison.OrdinalIgnoreCase))
            {
                LogFontInfo($"Using TTC face 0 for '{path}'");
                return new FontSource(path, 0);
            }

            return new FontSource(path, null);
        }

        public SKTypeface? CreateTypeface()
        {
            if (FaceIndex.HasValue)
            {
                if (TryExtractTtcFace(Path, FaceIndex.Value, out var fontData))
                {
                    using var stream = new SKMemoryStream(fontData);
                    var typeface = SKTypeface.FromStream(stream);
                    if (typeface != null)
                    {
                        LogFontInfo($"Embedded TTF slice for '{Path}'");
                        return typeface;
                    }

                    LogFontWarning($"SKTypeface.FromData failed for '{Path}', falling back to direct TTC access.");
                }

                return SKTypeface.FromFile(Path, FaceIndex.Value);
            }

            return SKTypeface.FromFile(Path);
        }

        private static bool TryExtractTtcFace(string path, int faceIndex, out SKData data)
        {
            data = null!;
            try
            {
                using var stream = File.OpenRead(path);
                using var reader = new BinaryReader(stream);

                uint ReadUInt32BE() => (uint)((reader.ReadByte() << 24) | (reader.ReadByte() << 16) | (reader.ReadByte() << 8) | reader.ReadByte());
                ushort ReadUInt16BE() => (ushort)((reader.ReadByte() << 8) | reader.ReadByte());

                if (ReadUInt32BE() != 0x74746366) // 'ttcf'
                    return false;

                ReadUInt32BE(); // version
                var numFonts = ReadUInt32BE();
                if (faceIndex < 0 || faceIndex >= numFonts)
                    return false;

                var offsets = new uint[numFonts];
                for (int i = 0; i < numFonts; i++)
                    offsets[i] = ReadUInt32BE();

                var start = offsets[faceIndex];
                stream.Seek(start, SeekOrigin.Begin);

                ReadUInt32BE(); // sfntVersion
                var numTables = ReadUInt16BE();
                ReadUInt16BE(); // searchRange
                ReadUInt16BE(); // entrySelector
                ReadUInt16BE(); // rangeShift

                long end = start + 12 + numTables * 16;
                for (int i = 0; i < numTables; i++)
                {
                    ReadUInt32BE(); // tag
                    ReadUInt32BE(); // checksum
                    var tableOffset = ReadUInt32BE();
                    var tableLength = ReadUInt32BE();
                    var padded = (tableLength + 3) & ~3u;
                    var tableEnd = start + tableOffset + padded;
                    if (tableEnd > end)
                        end = tableEnd;
                }

                var length = end - start;
                stream.Seek(start, SeekOrigin.Begin);
                var buffer = new byte[length];
                int read = 0;
                while (read < buffer.Length)
                {
                    var bytes = stream.Read(buffer, read, buffer.Length - read);
                    if (bytes == 0)
                        break;
                    read += bytes;
                }

                if (read != buffer.Length)
                    return false;

                data = SKData.CreateCopy(buffer);
                return true;
            }
            catch (Exception ex)
            {
                LogFontWarning($"Failed to extract TTC face from '{path}': {ex.Message}");
                return false;
            }
        }
    }

    private static bool ShouldLogFonts =>
        Environment.GetEnvironmentVariable("DOCXTOPDF_LOG_FONTS") == "1";

    private static void LogFontInfo(string message)
    {
        if (ShouldLogFonts)
            Console.Error.WriteLine($"[FontManager] {message}");
    }

    private static void LogFontWarning(string message)
    {
        if (ShouldLogFonts)
            Console.Error.WriteLine($"[FontManager][WARN] {message}");
    }
}
