using System;
using System.Buffers.Binary;
using System.Collections.Concurrent;
using SkiaSharp;

namespace DocxToPdf.Sdk.Text;

internal static class WordLineMetricsProvider
{
    private static readonly ConcurrentDictionary<nint, float?> Cache = new();
    private const uint Os2Tag =
        ((uint)'O' << 24) |
        ((uint)'S' << 16) |
        ((uint)'/' << 8) |
        (uint)'2';

    public static float? GetLineHeightScale(SKTypeface typeface)
    {
        if (typeface == null || typeface.Handle == IntPtr.Zero)
            return null;

        return Cache.GetOrAdd(typeface.Handle, _ => ComputeLineHeightScale(typeface));
    }

    private static float? ComputeLineHeightScale(SKTypeface typeface)
    {
        if (typeface.UnitsPerEm <= 0)
            return null;

        if (!TryReadOs2Metrics(typeface, out var ascender, out var descender, out var lineGap))
            return null;

        var height = (ascender - descender + lineGap) / (float)typeface.UnitsPerEm;
        if (height <= 0 || height > 10f)
            return null;

        return height;
    }

    private static bool TryReadOs2Metrics(SKTypeface typeface, out short ascender, out short descender, out short lineGap)
    {
        ascender = descender = lineGap = 0;
        var data = typeface.GetTableData(Os2Tag);
        if (data == null)
            return false;

        const int RequiredLength = 74;
        if (data.Length < RequiredLength)
            return false;

        var span = new ReadOnlySpan<byte>(data);
        ascender = ReadInt16(span, 68);
        descender = ReadInt16(span, 70);
        lineGap = ReadInt16(span, 72);
        return true;
    }

    private static short ReadInt16(ReadOnlySpan<byte> source, int offset) =>
        BinaryPrimitives.ReadInt16BigEndian(source.Slice(offset, 2));
}
