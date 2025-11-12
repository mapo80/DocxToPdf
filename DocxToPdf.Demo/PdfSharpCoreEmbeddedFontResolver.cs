using System;
using System.Collections.Generic;
using System.IO;
using PdfSharpCore.Fonts;

namespace DocxToPdf.Demo;

internal sealed class PdfSharpCoreEmbeddedFontResolver : IFontResolver
{
    private readonly Dictionary<string, string> _faceToPath = new(StringComparer.OrdinalIgnoreCase);
    public string DefaultFontName => "Arial";

    public PdfSharpCoreEmbeddedFontResolver(string fontsDir)
    {
        void Map(string face, string file)
        {
            var path = Path.Combine(fontsDir, file);
            if (File.Exists(path))
                _faceToPath[face] = path;
        }

        // Aptos
        Map("Aptos#Regular", "Aptos-Regular.ttf");
        Map("Aptos#Bold", "Aptos-Bold.ttf");
        Map("Aptos#Italic", "Aptos-Italic.ttf");
        Map("Aptos#BoldItalic", "Aptos-BoldItalic.ttf");

        // Calibri
        Map("Calibri#Regular", "Calibri.ttf");
        Map("Calibri#Bold", "Calibrib.ttf");
        Map("Calibri#Italic", "Calibrii.ttf");
        Map("Calibri#BoldItalic", "Calibriz.ttf");

        // Arial
        Map("Arial#Regular", "Arial-Regular.ttf");
        Map("Arial#Bold", "Arial-Bold.ttf");
        Map("Arial#Italic", "Arial-Italic.ttf");
        Map("Arial#BoldItalic", "Arial-BoldItalic.ttf");

        // Times New Roman
        Map("Times New Roman#Regular", "TimesNewRoman-Regular.ttf");
        Map("Times New Roman#Bold", "TimesNewRoman-Bold.ttf");
        Map("Times New Roman#Italic", "TimesNewRoman-Italic.ttf");
        Map("Times New Roman#BoldItalic", "TimesNewRoman-BoldItalic.ttf");

        // Symbol
        Map("Symbol#Regular", "Symbol.ttf");
    }

    public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        var style = isBold && isItalic ? "BoldItalic" : isBold ? "Bold" : isItalic ? "Italic" : "Regular";
        string face = $"{familyName}#{style}";

        if (!_faceToPath.ContainsKey(face))
        {
            if (string.Equals(familyName, "Times", StringComparison.OrdinalIgnoreCase))
                face = $"Times New Roman#{style}";
            else if (string.Equals(familyName, "Calibri Light", StringComparison.OrdinalIgnoreCase))
                face = $"Calibri#{style}";
            else if (string.Equals(familyName, "Cambria", StringComparison.OrdinalIgnoreCase))
                face = style switch
                {
                    "Bold" => "Cambria#Bold",
                    "Italic" => "Cambria#Italic",
                    "BoldItalic" => "Cambria#BoldItalic",
                    _ => "Times New Roman#Regular" // no Cambria regular TTF available
                };
        }

        return new FontResolverInfo(face);
    }

    public byte[] GetFont(string faceName)
    {
        if (_faceToPath.TryGetValue(faceName, out var path) && File.Exists(path))
            return File.ReadAllBytes(path);
        return Array.Empty<byte>();
    }
}

