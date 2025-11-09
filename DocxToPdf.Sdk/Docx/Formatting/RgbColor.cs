using System;

namespace DocxToPdf.Sdk.Docx.Formatting;

/// <summary>
/// Rappresenta un colore RGB (0-255) indipendente dalla libreria grafica.
/// </summary>
public readonly record struct RgbColor(byte R, byte G, byte B)
{
    public static readonly RgbColor Black = new(0, 0, 0);
    public static readonly RgbColor White = new(255, 255, 255);

    public static RgbColor FromHex(string hex)
    {
        if (string.IsNullOrWhiteSpace(hex))
            throw new ArgumentException("Hex color cannot be empty", nameof(hex));

        var cleaned = hex.Trim().TrimStart('#');
        if (cleaned.Length is not (6 or 3))
            throw new ArgumentException($"Hex color '{hex}' must have 3 or 6 characters.", nameof(hex));

        if (cleaned.Length == 3)
        {
            var r = Convert.ToByte(new string(cleaned[0], 2), 16);
            var g = Convert.ToByte(new string(cleaned[1], 2), 16);
            var b = Convert.ToByte(new string(cleaned[2], 2), 16);
            return new RgbColor(r, g, b);
        }

        var red = Convert.ToByte(cleaned[..2], 16);
        var green = Convert.ToByte(cleaned[2..4], 16);
        var blue = Convert.ToByte(cleaned[4..6], 16);
        return new RgbColor(red, green, blue);
    }

    public string ToHex() => $"{R:X2}{G:X2}{B:X2}";
}
