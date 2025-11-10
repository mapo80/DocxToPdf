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
    public Action<string>? DiagnosticsLogger { get; set; }

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
            shapedRuns.Add(ShapedRun.Create(run.Text, run.Typeface, sizePt, offset, enableKerning, DiagnosticsLogger));
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
    }

    public SKFontMetrics GetFontMetrics(SKTypeface typeface, float sizePt)
    {
        using var font = new SKFont(typeface, sizePt);
        return font.Metrics;
    }

    public float GetLineSpacing(SKTypeface typeface, float sizePt)
    {
        var preferredScale = WordLineMetricsProvider.GetLineHeightScale(typeface);
        if (preferredScale.HasValue)
            return sizePt * preferredScale.Value;

        var metrics = GetFontMetrics(typeface, sizePt);
        var naturalSpacing = metrics.Descent - metrics.Ascent + metrics.Leading;
        var fallbackSpacing = sizePt * 1.1667f;
        return Math.Max(naturalSpacing, fallbackSpacing);
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
        private readonly Action<string>? _diagnostics;

        private ShapedRun(
            string text,
            SKTypeface typeface,
            float sizePt,
            int startIndex,
            ushort[] glyphs,
            SKPoint[] positions,
            uint[] clusters,
            float width,
            Action<string>? diagnostics)
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
            _diagnostics = diagnostics;
        }

        public string Text { get; }
        public SKTypeface Typeface { get; }
        public float SizePt { get; }
        public int StartIndex { get; }
        public int EndIndex => StartIndex + Text.Length;
        public int Length => Text.Length;
        public float Width { get; }
        public float BaselineOffset { get; }

        public static ShapedRun Create(
            string text,
            SKTypeface typeface,
            float sizePt,
            int startIndex,
            bool enableKerning,
            Action<string>? diagnostics)
        {
            using var font = new SKFont(typeface, sizePt)
            {
                Subpixel = true,
                Hinting = SKFontHinting.None,
                Edging = SKFontEdging.SubpixelAntialias
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

            return new ShapedRun(text, typeface, sizePt, startIndex, glyphs, positions, clusters, width, diagnostics);
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
            if (localLength <= 0)
                return 0f;

            if (RequiresPathFallback(letterSpacingPt))
            {
                LogPathFallback(letterSpacingPt, localStart, localLength);
                return DrawAsPaths(canvas, x, y, color, localStart, localLength, letterSpacingPt);
            }

            return DrawWithSelectableText(canvas, x, y, color, localStart, localLength);
        }

        private static float ComputeBaselineOffset(SKTypeface typeface, float sizePt)
        {
            if (string.Equals(typeface.FamilyName, "Apple Color Emoji", StringComparison.OrdinalIgnoreCase))
                return sizePt * 0.35f;
            return 0f;
        }

        private static bool RequiresPathFallback(float letterSpacingPt) =>
            Math.Abs(letterSpacingPt) > 0.0001f;

        private void LogPathFallback(float letterSpacingPt, int localStart, int localLength)
        {
            if (_diagnostics == null)
                return;

            var sample = DescribeSegment(localStart, localLength);
            _diagnostics.Invoke(
                $"Selectable text disabled for run '{sample}' ({Text.Length} chars) because letter-spacing {letterSpacingPt:F3} pt requires path rendering.");
        }

        private string DescribeSegment(int localStart, int localLength)
        {
            if (string.IsNullOrEmpty(Text))
                return string.Empty;

            const int maxPreview = 32;
            var safeStart = Math.Clamp(localStart, 0, Math.Max(0, Text.Length - 1));
            var remaining = Text.Length - safeStart;
            var desiredLength = Math.Min(Math.Max(localLength, 1), remaining);
            var previewLength = Math.Min(desiredLength, maxPreview);
            var preview = Text.Substring(safeStart, previewLength);
            return previewLength < desiredLength ? $"{preview}…" : preview;
        }

        private float DrawWithSelectableText(
            SKCanvas canvas,
            float x,
            float y,
            SKColor color,
            int localStart,
            int localLength)
        {
            var localEnd = localStart + localLength;
            var startGlyph = FindGlyphIndex(localStart);
            var endGlyph = FindGlyphIndex(localEnd);
            if (endGlyph <= startGlyph)
                return 0f;

            var glyphCount = endGlyph - startGlyph;
            var spanWidth = Measure(localStart, localLength, 0f);
            var startAdvance = startGlyph < _advances.Length ? _advances[startGlyph] : 0f;
            var originX = x - startAdvance;

            using var paint = CreatePaint(color);
            using var blob = BuildTextBlob(startGlyph, glyphCount);
            SkiaInterop.DrawTextBlob(canvas, blob, originX, y + BaselineOffset, paint);

            return spanWidth;
        }

        private float DrawAsPaths(
            SKCanvas canvas,
            float x,
            float y,
            SKColor color,
            int localStart,
            int localLength,
            float letterSpacingPt)
        {
            var localEnd = localStart + localLength;
            var startGlyph = FindGlyphIndex(localStart);
            var endGlyph = FindGlyphIndex(localEnd);
            if (endGlyph <= startGlyph)
                return 0f;

            using var font = CreateFont();
            using var paint = CreatePaint(color);

            var count = endGlyph - startGlyph;
            var startAdvance = startGlyph < _advances.Length ? _advances[startGlyph] : Width;
            float extraOffset = 0f;
            for (int i = 0; i < count; i++)
            {
                var glyphIndex = startGlyph + i;
                var glyphId = _glyphs[glyphIndex];
                using var path = font.GetGlyphPath(glyphId);
                if (path == null)
                    continue;

                var pos = _positions[glyphIndex];
                var tx = pos.X - startAdvance + x + extraOffset;
                var ty = y + pos.Y + BaselineOffset;
                var matrix = SKMatrix.CreateTranslation(tx, ty);
                path.Transform(matrix);
                canvas.DrawPath(path, paint);

                if (i < count - 1 && Math.Abs(letterSpacingPt) > 0.0001f)
                    extraOffset += letterSpacingPt;
            }

            return Measure(localStart, localLength, letterSpacingPt);
        }

        private SKFont CreateFont() => new(Typeface, SizePt)
        {
            Subpixel = true,
            Hinting = SKFontHinting.None,
            Edging = SKFontEdging.SubpixelAntialias
        };

        private static SKPaint CreatePaint(SKColor color) => new()
        {
            IsAntialias = true,
            Color = color,
            Style = SKPaintStyle.Fill
        };

        private SKTextBlob BuildTextBlob(int startGlyph, int glyphCount)
        {
            using var builder = new SKTextBlobBuilder();
            using var font = CreateFont();
            var run = builder.AllocatePositionedRun(font, glyphCount);
#pragma warning disable CS0618
            var glyphSpan = run.GetGlyphSpan();
            var posSpan = run.GetPositionSpan();
#pragma warning restore CS0618
            for (int i = 0; i < glyphCount; i++)
            {
                var glyphIndex = startGlyph + i;
                glyphSpan[i] = _glyphs[glyphIndex];
                var pos = _positions[glyphIndex];
                posSpan[i] = pos;
            }
            return builder.Build();
        }

    }

    #endregion
}
