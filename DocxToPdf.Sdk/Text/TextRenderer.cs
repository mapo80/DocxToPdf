using SkiaSharp;
using SkiaSharp.HarfBuzz;
using System;
using System.Collections.Generic;
using System.Text;

namespace DocxToPdf.Sdk.Text;

/// <summary>
/// Renderer di testo con shaping avanzato tramite HarfBuzz.
/// Gestisce legature, diacritici, script complessi e posizionamento preciso dei glifi.
/// </summary>
public sealed class TextRenderer
{
    /// <summary>
    /// Disegna testo con full text shaping (legature, kerning, script complessi).
    /// </summary>
    /// <param name="canvas">Canvas SkiaSharp su cui disegnare</param>
    /// <param name="text">Testo da renderizzare</param>
    /// <param name="x">Coordinata X baseline in pt</param>
    /// <param name="y">Coordinata Y baseline in pt</param>
    /// <param name="typeface">Font typeface</param>
    /// <param name="sizePt">Dimensione font in pt</param>
    /// <param name="color">Colore del testo</param>
    /// <returns>Larghezza del testo renderizzato in pt</returns>
    public float DrawShapedText(
        SKCanvas canvas,
        string text,
        float x,
        float y,
        SKTypeface typeface,
        float sizePt,
        SKColor color)
    {
        if (string.IsNullOrEmpty(text))
            return 0f;

        using var font = new SKFont(typeface, sizePt)
        {
            Subpixel = true,
            Edging = SKFontEdging.Antialias // Antialiasing per PDF
        };

        using var paint = new SKPaint
        {
            Color = color,
            IsAntialias = true
        };

        using var shaper = new SKShaper(typeface);

        // Shape del testo: HarfBuzz analizza il testo e restituisce i glifi posizionati
        var result = shaper.Shape(text, font);

        // Disegna i glifi posizionati
        canvas.DrawShapedText(shaper, text, x, y, SKTextAlign.Left, font, paint);

        return GetAdvanceWidth(result, font, text);
    }

    /// <summary>
    /// Misura la larghezza del testo con shaping.
    /// </summary>
    public float MeasureText(string text, SKTypeface typeface, float sizePt)
    {
        if (string.IsNullOrEmpty(text))
            return 0f;

        using var font = new SKFont(typeface, sizePt);
        using var shaper = new SKShaper(typeface);

        var result = shaper.Shape(text, font);
        return GetAdvanceWidth(result, font, text);
    }

    /// <summary>
    /// Disegna testo con font fallback automatico per caratteri non supportati (emoji, CJK, etc).
    /// Spezza il testo in run con font appropriati per ogni segmento.
    /// </summary>
    public float DrawShapedTextWithFallback(
        SKCanvas canvas,
        string text,
        float x,
        float y,
        SKTypeface primaryTypeface,
        float sizePt,
        SKColor color)
    {
        if (string.IsNullOrEmpty(text))
            return 0f;

        var fontManager = FontManager.Instance;
        var runs = SplitTextIntoFontRuns(text, primaryTypeface, fontManager);

        float currentX = x;

        foreach (var run in runs)
        {
            var width = DrawShapedText(canvas, run.Text, currentX, y, run.Typeface, sizePt, color);
            currentX += width;
        }

        return currentX - x; // Ritorna la larghezza totale
    }

    /// <summary>
    /// Misura la larghezza del testo con font fallback automatico.
    /// </summary>
    public float MeasureTextWithFallback(string text, SKTypeface primaryTypeface, float sizePt)
    {
        if (string.IsNullOrEmpty(text))
            return 0f;

        var fontManager = FontManager.Instance;
        var runs = SplitTextIntoFontRuns(text, primaryTypeface, fontManager);

        float totalWidth = 0f;

        foreach (var run in runs)
        {
            totalWidth += MeasureText(run.Text, run.Typeface, sizePt);
        }

        return totalWidth;
    }

    /// <summary>
    /// Calcola le metriche del font per una data typeface e dimensione.
    /// Restituisce ascent (negativo), descent (positivo) e leading.
    /// </summary>
    public SKFontMetrics GetFontMetrics(SKTypeface typeface, float sizePt)
    {
        using var font = new SKFont(typeface, sizePt);
        return font.Metrics;
    }

    /// <summary>
    /// Calcola l'altezza della riga (line spacing) per un dato font.
    /// Line spacing = descent - ascent + leading
    /// </summary>
    public float GetLineSpacing(SKTypeface typeface, float sizePt)
    {
        var metrics = GetFontMetrics(typeface, sizePt);
        return metrics.Descent - metrics.Ascent + metrics.Leading;
    }

    /// <summary>
    /// Disegna testo centrato orizzontalmente nell'area specificata.
    /// </summary>
    public void DrawCenteredText(
        SKCanvas canvas,
        string text,
        float centerX,
        float y,
        float maxWidth,
        SKTypeface typeface,
        float sizePt,
        SKColor color)
    {
        var textWidth = MeasureTextWithFallback(text, typeface, sizePt);
        var x = centerX - (textWidth / 2f);

        // Clamp per evitare overflow
        if (x < 0) x = 0;
        if (x + textWidth > maxWidth) x = maxWidth - textWidth;

        DrawShapedTextWithFallback(canvas, text, x, y, typeface, sizePt, color);
    }

    /// <summary>
    /// Rappresenta un segmento di testo con un font specifico.
    /// </summary>
    private record struct FontRun(string Text, SKTypeface Typeface);

    /// <summary>
    /// Spezza il testo in run, ognuno con un font appropriato che supporta i suoi caratteri.
    /// </summary>
    private List<FontRun> SplitTextIntoFontRuns(string text, SKTypeface primaryTypeface, FontManager fontManager)
    {
        var runs = new List<FontRun>();
        var currentRun = new StringBuilder();
        SKTypeface? currentTypeface = null;

        for (int i = 0; i < text.Length; i++)
        {
            int codepoint;

            // Gestisci surrogate pairs per emoji e caratteri oltre U+FFFF
            if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
            {
                codepoint = char.ConvertToUtf32(text[i], text[i + 1]);
            }
            else
            {
                codepoint = text[i];
            }

            // Determina quale font usare per questo carattere
            SKTypeface typefaceForChar;

            if (fontManager.TypefaceContainsGlyph(primaryTypeface, codepoint))
            {
                typefaceForChar = primaryTypeface;
            }
            else
            {
                // Cerca un font fallback
                var fallback = fontManager.FindFallbackTypeface(codepoint, primaryTypeface);
                typefaceForChar = fallback ?? primaryTypeface; // Usa primario se fallback non trovato
            }

            // Se il font cambia, salva il run corrente e iniziane uno nuovo
            if (currentTypeface != null && currentTypeface != typefaceForChar)
            {
                if (currentRun.Length > 0)
                {
                    runs.Add(new FontRun(currentRun.ToString(), currentTypeface));
                    currentRun.Clear();
                }
            }

            currentTypeface = typefaceForChar;

            // Aggiungi il carattere al run corrente
            if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
            {
                currentRun.Append(text[i]);
                currentRun.Append(text[i + 1]);
                i++; // Salta il low surrogate
            }
            else
            {
                currentRun.Append(text[i]);
            }
        }

        // Aggiungi l'ultimo run
        if (currentRun.Length > 0 && currentTypeface != null)
        {
            runs.Add(new FontRun(currentRun.ToString(), currentTypeface));
        }

        return runs;
    }
    /// <summary>
    /// Restituisce la larghezza avanzata di un testo shappato. HarfBuzz può non produrre
    /// glifi per stringhe composte solo da whitespace, quindi in quel caso stimiamo la
    /// larghezza calcolando direttamente gli advance dei glifi tramite SKFont.
    /// </summary>
    private static float GetAdvanceWidth(
        SKShaper.Result result,
        SKFont font,
        string text)
    {
        if (result.Width > 0)
        {
            return result.Width;
        }

        var points = result.Points;
        if (points.Length > 0)
        {
            return points[^1].X;
        }

        return MeasureWithFont(font, text);
    }

    private static float MeasureWithFont(SKFont font, string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return 0f;
        }

        var textSpan = text.AsSpan();
        var glyphCount = font.CountGlyphs(textSpan);
        if (glyphCount == 0)
        {
            return 0f;
        }

        Span<ushort> glyphBuffer = text.Length <= 256
            ? stackalloc ushort[text.Length]
            : new ushort[text.Length];

        font.GetGlyphs(textSpan, glyphBuffer);

        Span<float> widthBuffer = glyphCount <= 256
            ? stackalloc float[glyphCount]
            : new float[glyphCount];

        font.GetGlyphWidths(glyphBuffer[..glyphCount], widthBuffer, Span<SKRect>.Empty, null);

        float total = 0f;
        for (int i = 0; i < glyphCount; i++)
        {
            total += widthBuffer[i];
        }

        return total;
    }
}
