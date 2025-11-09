using SkiaSharp;
using SkiaSharp.HarfBuzz;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocxToPdf.Sdk.Text;

/// <summary>
/// Renderer di testo basato su Skia + HarfBuzz con supporto a font fallback e accesso ai glifi shapati.
/// </summary>
public sealed class TextRenderer
{
    private readonly FontManager _fontManager = FontManager.Instance;

    /// <summary>
    /// Shapa il testo completo usando HarfBuzz e restituisce un oggetto riutilizzabile.
    /// </summary>
    public ShapedText Shape(string text, SKTypeface primaryTypeface, float sizePt, bool enableKerning = false)
    {
        text ??= string.Empty;
        var runs = SplitIntoFontRuns(text, primaryTypeface);
        var shapedRuns = new List<ShapedRun>(runs.Count);
        int offset = 0;
        foreach (var run in runs)
        {
            if (run.Text.Length == 0)
                continue;
            shapedRuns.Add(ShapedRun.Create(run.Text, run.Typeface, sizePt, offset, enableKerning));
            offset += run.Text.Length;
        }
        return new ShapedText(text, shapedRuns);
    }

    public float MeasureTextWithFallback(string text, SKTypeface primaryTypeface, float sizePt, bool enableKerning = false) =>
        Shape(text, primaryTypeface, sizePt, enableKerning).Width;

    public float DrawShapedTextWithFallback(
        SKCanvas canvas,
        string text,
        float x,
        float y,
        SKTypeface primaryTypeface,
        float sizePt,
        SKColor color,
        float letterSpacingPt = 0f,
        bool enableKerning = false)
    {
        var shaped = Shape(text, primaryTypeface, sizePt, enableKerning);
        return shaped.DrawRange(canvas, x, y, color, 0, text.Length, letterSpacingPt);
    }

    public float DrawShapedText(SKCanvas canvas, string text, float x, float y, SKTypeface typeface, float sizePt, SKColor color)
    {
        var shaped = Shape(text, typeface, sizePt);
        return shaped.Draw(canvas, x, y, color);
    }

    public void DrawCenteredText(SKCanvas canvas, string text, float centerX, float y, float maxWidth, SKTypeface typeface, float sizePt, SKColor color)
    {
        var shaped = Shape(text, typeface, sizePt);
        var textWidth = shaped.Width;
        var clampedWidth = Math.Min(textWidth, maxWidth);
        var startX = centerX - clampedWidth / 2f;
        shaped.Draw(canvas, startX, y, color);
        DrawInvisibleText(canvas, text, startX, y, typeface, sizePt);
    }

    public void DrawInvisibleText(SKCanvas canvas, string text, float x, float y, SKTypeface typeface, float sizePt)
    {
        if (string.IsNullOrEmpty(text))
            return;

        using var paint = new SKPaint
        {
            Color = new SKColor(0, 0, 0, 1),
            IsStroke = false,
            IsAntialias = false,
            TextEncoding = SKTextEncoding.Utf16
        };

        var runs = SplitIntoFontRuns(text, typeface);
        float cursor = x;
        foreach (var run in runs)
        {
            using var font = new SKFont(run.Typeface, sizePt)
            {
                Subpixel = true,
                Edging = SKFontEdging.Antialias
            };
            canvas.DrawText(run.Text, cursor, y, SKTextAlign.Left, font, paint);
            cursor += font.MeasureText(run.Text);
        }
    }

    public SKFontMetrics GetFontMetrics(SKTypeface typeface, float sizePt)
    {
        using var font = new SKFont(typeface, sizePt);
        return font.Metrics;
    }

    public float GetLineSpacing(SKTypeface typeface, float sizePt)
    {
        var metrics = GetFontMetrics(typeface, sizePt);
        return metrics.Descent - metrics.Ascent + metrics.Leading;
    }

    private List<FontRun> SplitIntoFontRuns(string text, SKTypeface primaryTypeface)
    {
        var runs = new List<FontRun>();
        if (string.IsNullOrEmpty(text))
        {
            runs.Add(new FontRun(string.Empty, primaryTypeface));
            return runs;
        }

        var currentFont = primaryTypeface;
        var buffer = new System.Text.StringBuilder();

        for (int i = 0; i < text.Length; i++)
        {
            var codepoint = char.ConvertToUtf32(text, i);
            if (char.IsHighSurrogate(text[i]))
                i++;

            var typeface = _fontManager.TypefaceContainsGlyph(currentFont, codepoint)
                ? currentFont
                : _fontManager.FindFallbackTypeface(codepoint, currentFont) ?? currentFont;

            if (codepoint == 0x2122)
            {
                Console.WriteLine($"SplitIntoFontRuns selects {typeface.FamilyName} for TM (current {currentFont.FamilyName})");
            }

            if (typeface != currentFont && buffer.Length > 0)
            {
                runs.Add(new FontRun(buffer.ToString(), currentFont));
                buffer.Clear();
            }

            currentFont = typeface;
            buffer.Append(char.ConvertFromUtf32(codepoint));
        }

        if (buffer.Length > 0)
            runs.Add(new FontRun(buffer.ToString(), currentFont));

        return runs;
    }

    private record struct FontRun(string Text, SKTypeface Typeface);

    #region Shaped types

    public sealed class ShapedText
    {
        private readonly IReadOnlyList<ShapedRun> _runs;
        public string Source { get; }
        public float Width { get; }

        internal ShapedText(string source, IReadOnlyList<ShapedRun> runs)
        {
            Source = source;
            _runs = runs;
            float width = 0f;
            foreach (var run in runs)
                width += run.Width;
            Width = width;
        }

        public float Draw(SKCanvas canvas, float x, float y, SKColor color, float letterSpacingPt = 0f)
        {
            float cursor = x;
            foreach (var run in _runs)
            {
                cursor += run.Draw(canvas, cursor, y, color, 0, run.Length, letterSpacingPt);
            }
            return cursor - x;
        }

        public float DrawRange(SKCanvas canvas, float x, float y, SKColor color, int start, int length, float letterSpacingPt = 0f)
        {
            float cursor = x;
            int end = start + length;
            foreach (var run in _runs)
            {
                if (!run.Overlaps(start, end))
                    continue;
                var localStart = Math.Max(start, run.StartIndex);
                var localEnd = Math.Min(end, run.EndIndex);
                cursor += run.Draw(canvas, cursor, y, color, localStart - run.StartIndex, localEnd - localStart, letterSpacingPt);
            }
            return cursor - x;
        }
        public float MeasureRange(int start, int length, float letterSpacingPt = 0f)
        {
            if (length <= 0)
                return 0f;
            float width = 0f;
            int end = start + length;
            foreach (var run in _runs)
            {
                if (!run.Overlaps(start, end))
                    continue;
                var localStart = Math.Max(start, run.StartIndex);
                var localEnd = Math.Min(end, run.EndIndex);
                width += run.Measure(localStart - run.StartIndex, localEnd - localStart, letterSpacingPt);
            }
            return width;
        }

        public IEnumerable<ShapedRun> Runs => _runs;
    }

    public sealed class ShapedRun
    {
        private readonly ushort[] _glyphs;
        private readonly SKPoint[] _positions;
        private readonly uint[] _clusters;
        private readonly float[] _advances;

        private ShapedRun(string text, SKTypeface typeface, float sizePt, int startIndex, ushort[] glyphs, SKPoint[] positions, uint[] clusters, float width)
        {
            Text = text;
            Typeface = typeface;
            SizePt = sizePt;
            StartIndex = startIndex;
            Width = width;
            _glyphs = glyphs;
            _positions = positions;
            _clusters = clusters;
            _advances = new float[_glyphs.Length + 1];
            _advances[0] = 0f;
            for (int i = 0; i < _glyphs.Length; i++)
            {
                var boundary = (i + 1 < _positions.Length) ? _positions[i + 1].X : width;
                _advances[i + 1] = boundary;
            }
            BaselineOffset = ComputeBaselineOffset(typeface, sizePt);
        }

        public string Text { get; }
        public SKTypeface Typeface { get; }
        public float SizePt { get; }
        public int StartIndex { get; }
        public int EndIndex => StartIndex + Text.Length;
        public int Length => Text.Length;
        public float Width { get; }
        public float BaselineOffset { get; }

        public static ShapedRun Create(string text, SKTypeface typeface, float sizePt, int startIndex, bool enableKerning)
        {
            using var font = new SKFont(typeface, sizePt)
            {
                Subpixel = true,
                Edging = SKFontEdging.Antialias
            };
            using var shaper = new SKShaper(typeface);
            var result = shaper.Shape(text, font);
            var glyphs = Array.ConvertAll(result.Codepoints, c => (ushort)c);
            var positions = result.Points;
            var clusters = result.Clusters;
            var width = result.Width;

            if (!enableKerning && glyphs.Length > 0)
            {
                using var widthsPaint = new SKPaint();
                var newPositions = new SKPoint[positions.Length];
                var glyphWidths = font.GetGlyphWidths(glyphs.AsSpan(), widthsPaint);
                float cursor = 0f;
                for (int i = 0; i < glyphs.Length; i++)
                {
                    newPositions[i] = new SKPoint(cursor, positions[i].Y);
                    cursor += glyphWidths[i];
                }
                positions = newPositions;
                width = cursor;
            }

            return new ShapedRun(text, typeface, sizePt, startIndex, glyphs, positions, clusters, width);
        }

        public bool Overlaps(int start, int end) => start < EndIndex && end > StartIndex;

        private int FindGlyphIndex(int localCharIndex)
        {
            for (int i = 0; i < _clusters.Length; i++)
            {
                if (_clusters[i] >= localCharIndex)
                    return i;
            }
            return _clusters.Length;
        }

        public float Measure(int localStart, int localLength, float letterSpacingPt = 0f)
        {
            var localEnd = localStart + localLength;
            var startGlyph = FindGlyphIndex(localStart);
            var endGlyph = FindGlyphIndex(localEnd);
            var startAdvance = startGlyph < _advances.Length ? _advances[startGlyph] : Width;
            var endAdvance = endGlyph < _advances.Length ? _advances[endGlyph] : Width;
            var baseWidth = endAdvance - startAdvance;
            var glyphCount = Math.Max(0, endGlyph - startGlyph);
            if (glyphCount > 1 && Math.Abs(letterSpacingPt) > 0.0001f)
                baseWidth += letterSpacingPt * (glyphCount - 1);
            return baseWidth;
        }

        public float Draw(
            SKCanvas canvas,
            float x,
            float y,
            SKColor color,
            int localStart,
            int localLength,
            float letterSpacingPt = 0f)
        {
            var glyphCount = _glyphs.Length;
            if (glyphCount == 0 || localLength <= 0)
                return 0f;

            var localEnd = localStart + localLength;
            var startGlyph = FindGlyphIndex(localStart);
            var endGlyph = FindGlyphIndex(localEnd);
            if (endGlyph <= startGlyph)
                return 0f;

            using var font = new SKFont(Typeface, SizePt)
            {
                Subpixel = true,
                Edging = SKFontEdging.Antialias
            };
            using var paint = new SKPaint
            {
                IsAntialias = true,
                Color = color
            };
            using var builder = new SKTextBlobBuilder();
            var count = endGlyph - startGlyph;
            var buffer = builder.AllocatePositionedRun(font, count);
            var glyphSpan = buffer.Glyphs;
            var posSpan = buffer.Positions;
            var startAdvance = startGlyph < _advances.Length ? _advances[startGlyph] : Width;
            float extraOffset = 0f;
            for (int i = 0; i < count; i++)
            {
                var glyphIndex = startGlyph + i;
                glyphSpan[i] = _glyphs[glyphIndex];
                var pos = _positions[glyphIndex];
                posSpan[i] = new SKPoint(pos.X - startAdvance + x + extraOffset, y + pos.Y + BaselineOffset);
                if (i < count - 1 && Math.Abs(letterSpacingPt) > 0.0001f)
                    extraOffset += letterSpacingPt;
            }
            using var blob = builder.Build();
            canvas.DrawText(blob, 0, 0, paint);

            return Measure(localStart, localLength, letterSpacingPt);
        }

        private static float ComputeBaselineOffset(SKTypeface typeface, float sizePt)
        {
            if (string.Equals(typeface.FamilyName, "Apple Color Emoji", StringComparison.OrdinalIgnoreCase))
                return sizePt * 0.35f;
            return 0f;
        }
    }

    #endregion
}
