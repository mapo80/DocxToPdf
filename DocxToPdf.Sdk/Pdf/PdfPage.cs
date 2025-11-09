using SkiaSharp;

namespace DocxToPdf.Sdk.Pdf;

/// <summary>
/// Rappresenta una pagina PDF attiva, fornendo API per disegnare contenuti.
/// Coordinate in punti tipografici (pt), origine in alto-sinistra.
/// </summary>
public sealed class PdfPage
{
    private readonly SKCanvas _canvas;
    private readonly PaperSize _size;

    internal PdfPage(SKCanvas canvas, PaperSize size)
    {
        _canvas = canvas;
        _size = size;
    }

    public PaperSize Size => _size;
    public SKCanvas Canvas => _canvas;

    /// <summary>
    /// Disegna testo semplice (senza shaping complesso) a coordinate specificate.
    /// Per text shaping avanzato usare <see cref="DocxToPdf.Sdk.Text.TextRenderer"/>.
    /// </summary>
    /// <param name="text">Testo da disegnare</param>
    /// <param name="x">Coordinata X in pt</param>
    /// <param name="y">Coordinata Y in pt (baseline del testo)</param>
    /// <param name="typeface">Font typeface</param>
    /// <param name="sizePt">Dimensione font in pt</param>
    /// <param name="color">Colore del testo</param>
    public void DrawSimpleText(string text, float x, float y, SKTypeface typeface, float sizePt, SKColor color)
    {
        using var font = new SKFont(typeface, sizePt)
        {
            Subpixel = true
        };

        using var paint = new SKPaint
        {
            Color = color,
            IsAntialias = true
        };

        _canvas.DrawText(text, x, y, SKTextAlign.Left, font, paint);
    }
}
