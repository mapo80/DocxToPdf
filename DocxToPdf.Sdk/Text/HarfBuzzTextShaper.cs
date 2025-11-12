using HarfBuzzSharp;
using SkiaSharp;
using SkiaSharp.HarfBuzz;
using System;
using Buffer = HarfBuzzSharp.Buffer;

namespace DocxToPdf.Sdk.Text;

/// <summary>
/// Minimal HarfBuzz-based shaper that mirrors SKShaper but allows passing custom OpenType features.
/// </summary>
internal sealed class HarfBuzzTextShaper : IDisposable
{
    private const int FontSizeScale = 512;

    private readonly Font _hbFont;
    private readonly Buffer _buffer;
    private readonly Feature[] _features;

    public HarfBuzzTextShaper(SKTypeface typeface, Feature[]? features = null)
    {
        Typeface = typeface ?? throw new ArgumentNullException(nameof(typeface));

        int index;
        using var stream = Typeface.OpenStream(out index);
        using var blob = stream.ToHarfBuzzBlob();
        using var face = new Face(blob, index)
        {
            Index = index,
            UnitsPerEm = Typeface.UnitsPerEm
        };

        _hbFont = new Font(face);
        _hbFont.SetScale(FontSizeScale, FontSizeScale);
        _hbFont.SetFunctionsOpenType();

        _buffer = new Buffer
        {
            ContentType = ContentType.Unicode
        };
        _features = features ?? Array.Empty<Feature>();
    }

    public SKTypeface Typeface { get; }

    public SKShaper.Result Shape(string text, SKFont font) =>
        Shape(text, 0f, 0f, font);

    public SKShaper.Result Shape(string text, float xOffset, float yOffset, SKFont font)
    {
        if (string.IsNullOrEmpty(text) || font == null)
            return new SKShaper.Result();

        font.Typeface = Typeface;

        _buffer.ClearContents();
        _buffer.AddUtf16(text);
        _buffer.GuessSegmentProperties();

        return ShapeBuffer(xOffset, yOffset, font);
    }

    private SKShaper.Result ShapeBuffer(float xOffset, float yOffset, SKFont font)
    {
        _hbFont.Shape(_buffer, _features);

        var len = _buffer.Length;
        if (len == 0)
            return new SKShaper.Result();

        var info = _buffer.GlyphInfos;
        var pos = _buffer.GlyphPositions;

        var points = new SKPoint[len];
        var clusters = new uint[len];
        var codepoints = new uint[len];

        var textSizeY = font.Size / FontSizeScale;
        var textSizeX = textSizeY * font.ScaleX;

        var originX = xOffset;
        var originY = yOffset;

        for (int i = 0; i < len; i++)
        {
            codepoints[i] = info[i].Codepoint;
            clusters[i] = info[i].Cluster;
            points[i] = new SKPoint(
                originX + pos[i].XOffset * textSizeX,
                originY - pos[i].YOffset * textSizeY);

            originX += pos[i].XAdvance * textSizeX;
            originY += pos[i].YAdvance * textSizeY;
        }

        var width = originX - xOffset;
        return new SKShaper.Result(codepoints, clusters, points, width);
    }

    public void Dispose()
    {
        _hbFont.Dispose();
        _buffer.Dispose();
    }
}
