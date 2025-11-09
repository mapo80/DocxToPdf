namespace DocxToPdf.Sdk.Pdf;

/// <summary>
/// Definisce le dimensioni standard della carta in punti tipografici (1 pt = 1/72").
/// </summary>
public readonly record struct PaperSize(float WidthPt, float HeightPt)
{
    /// <summary>
    /// A4: 210 × 297 mm ≈ 595.276 × 841.890 pt
    /// </summary>
    public static readonly PaperSize A4 = new(595.276f, 841.890f);

    /// <summary>
    /// Letter: 8.5 × 11 in = 612 × 792 pt
    /// </summary>
    public static readonly PaperSize Letter = new(612f, 792f);

    /// <summary>
    /// Legal: 8.5 × 14 in = 612 × 1008 pt
    /// </summary>
    public static readonly PaperSize Legal = new(612f, 1008f);
}
