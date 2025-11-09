using DocxToPdf.Sdk.Docx;
using DocxToPdf.Sdk.Docx.Formatting;
using DocxToPdf.Sdk.Text;
using SkiaSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocxToPdf.Sdk.Layout;

/// <summary>
/// Engine di layout per testo con wrapping automatico.
/// Usa algoritmo greedy per spezzare il testo in righe.
/// </summary>
public sealed class TextLayoutEngine
{
    private readonly TextRenderer _textRenderer;
    private readonly FontManager _fontManager;

    public TextLayoutEngine()
    {
        _textRenderer = new TextRenderer();
        _fontManager = FontManager.Instance;
    }

    /// <summary>
    /// Layout di un paragrafo con wrapping automatico.
    /// Restituisce le righe pronte per il rendering.
    /// </summary>
    public List<LayoutLine> LayoutParagraph(
        DocxParagraph paragraph,
        float maxWidthPt)
    {
        var lines = new List<LayoutLine>();
        var currentLine = new List<LayoutRun>();
        float currentLineWidth = 0f;
        float currentMaxAscent = 0f;
        float currentMaxDescent = 0f;
        float currentMaxLeading = 0f;

        var formatting = paragraph.ParagraphFormatting;
        var firstLineOffset = formatting.GetFirstLineOffsetPt();
        var subsequentOffset = formatting.GetSubsequentLineOffsetPt();
        var rightIndent = formatting.RightIndentPt;
        var hasMarker = paragraph.ListMarker != null;
        var textIndentFirstLine = hasMarker ? subsequentOffset : firstLineOffset;

        float SafeWidth(float width) => Math.Max(1f, width);

        var firstLineWidth = SafeWidth(maxWidthPt - textIndentFirstLine - rightIndent);
        var otherLinesWidth = SafeWidth(maxWidthPt - subsequentOffset - rightIndent);
        var currentLineLimit = firstLineWidth;
        var isFirstLine = true;

        foreach (var docxRun in paragraph.Runs)
        {
            var typeface = GetTypefaceForRun(docxRun);
            var fontSize = docxRun.Formatting.FontSizePt;

            // Spezza il run in parole (greedy word wrapping)
            var words = SplitIntoWords(docxRun.Text);

            foreach (var word in words)
            {
                var wordWidth = _textRenderer.MeasureTextWithFallback(word, typeface, fontSize);

                // Se la parola non sta nella riga corrente, vai a capo
                if (currentLine.Count > 0 && currentLineWidth + wordWidth > currentLineLimit)
                {
                    // Salva la riga corrente (crea una COPIA della lista)
                    lines.Add(new LayoutLine(
                        new List<LayoutRun>(currentLine),
                        currentLineWidth,
                        currentMaxAscent,
                        currentMaxDescent,
                        currentMaxLeading,
                        isFirstLine,
                        currentLineLimit
                    ));

                    // Reset per la nuova riga
                    currentLine.Clear();
                    currentLineWidth = 0f;
                    currentMaxAscent = 0f;
                    currentMaxDescent = 0f;
                    currentMaxLeading = 0f;
                    isFirstLine = false;
                    currentLineLimit = otherLinesWidth;
                }

                // Aggiungi la parola alla riga corrente
                currentLine.Add(new LayoutRun(word, typeface, fontSize, docxRun.Formatting));
                currentLineWidth += wordWidth;

                // Aggiorna le metriche della riga (max ascent/descent/leading)
                var metrics = _textRenderer.GetFontMetrics(typeface, fontSize);
                // Ascent è negativo, quindi usiamo Min per ottenere il valore più grande in valore assoluto
                currentMaxAscent = Math.Min(currentMaxAscent, metrics.Ascent);
                currentMaxDescent = Math.Max(currentMaxDescent, metrics.Descent);
                currentMaxLeading = Math.Max(currentMaxLeading, metrics.Leading);
            }
        }

        // Aggiungi l'ultima riga
        if (currentLine.Count > 0)
        {
            lines.Add(new LayoutLine(
                new List<LayoutRun>(currentLine),
                currentLineWidth,
                currentMaxAscent,
                currentMaxDescent,
                currentMaxLeading,
                isFirstLine,
                currentLineLimit
            ));
        }

        // Se il paragrafo è vuoto, aggiungi una riga vuota
        if (lines.Count == 0)
        {
            var defaultTypeface = _fontManager.GetDefaultTypeface();
            var defaultFontSize = paragraph.Runs.FirstOrDefault()?.Formatting.FontSizePt ?? 11f;
            var defaultMetrics = _textRenderer.GetFontMetrics(defaultTypeface, defaultFontSize);
            lines.Add(new LayoutLine(
                new List<LayoutRun>(),
                0f,
                defaultMetrics.Ascent,
                defaultMetrics.Descent,
                defaultMetrics.Leading,
                true,
                firstLineWidth
            ));
        }

        return lines;
    }

    /// <summary>
    /// Ottiene il typeface appropriato per un run DOCX.
    /// </summary>
    private SKTypeface GetTypefaceForRun(DocxRun run)
    {
        var fontFamily = run.Formatting.FontFamily;
        return _fontManager.GetTypeface(fontFamily, run.Formatting.Bold, run.Formatting.Italic);
    }

    /// <summary>
    /// Spezza il testo in parole (greedy word breaking).
    /// Include gli spazi come parti separate per preservare la spaziatura.
    /// </summary>
    private static List<string> SplitIntoWords(string text)
    {
        var words = new List<string>();
        var currentWord = new StringBuilder();

        foreach (var ch in text)
        {
            if (char.IsWhiteSpace(ch))
            {
                // Salva la parola corrente
                if (currentWord.Length > 0)
                {
                    words.Add(currentWord.ToString());
                    currentWord.Clear();
                }

                // Aggiungi lo spazio come token separato
                words.Add(ch.ToString());
            }
            else
            {
                currentWord.Append(ch);
            }
        }

        // Aggiungi l'ultima parola
        if (currentWord.Length > 0)
        {
            words.Add(currentWord.ToString());
        }

        return words;
    }
}

/// <summary>
/// Rappresenta una riga di testo dopo il layout.
/// </summary>
public sealed record LayoutLine(
    IReadOnlyList<LayoutRun> Runs,
    float WidthPt,
    float MaxAscent,    // Negativo: distanza dalla baseline verso l'alto
    float MaxDescent,   // Positivo: distanza dalla baseline verso il basso
    float MaxLeading,   // Spazio extra tra le righe
    bool IsFirstLine,
    float AvailableWidthPt
)
{
    /// <summary>
    /// Calcola l'altezza totale della riga (line spacing).
    /// Line spacing = descent - ascent + leading
    /// </summary>
    public float GetLineSpacing() => MaxDescent - MaxAscent + MaxLeading;
};

/// <summary>
/// Rappresenta un run di testo con font e dimensione per il rendering.
/// </summary>
public sealed record LayoutRun(
    string Text,
    SKTypeface Typeface,
    float FontSizePt,
    RunFormatting Formatting
);
