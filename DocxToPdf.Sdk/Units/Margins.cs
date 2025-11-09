namespace DocxToPdf.Sdk.Units;

/// <summary>
/// Rappresenta i margini di una pagina in punti tipografici.
/// </summary>
public readonly record struct Margins(float Top, float Right, float Bottom, float Left)
{
    /// <summary>
    /// Margini di default: 1 inch (72 pt) su tutti i lati, compatibile con Word.
    /// </summary>
    public static readonly Margins Default = new(72f, 72f, 72f, 72f);

    /// <summary>
    /// Margini stretti: 0.5 inch (36 pt) su tutti i lati.
    /// </summary>
    public static readonly Margins Narrow = new(36f, 36f, 36f, 36f);

    /// <summary>
    /// Nessun margine.
    /// </summary>
    public static readonly Margins None = new(0f, 0f, 0f, 0f);

    /// <summary>
    /// Crea margini da valori DXA (unità WordprocessingML).
    /// </summary>
    public static Margins FromDxa(int top, int right, int bottom, int left) =>
        new(
            UnitConverter.DxaToPoints(top),
            UnitConverter.DxaToPoints(right),
            UnitConverter.DxaToPoints(bottom),
            UnitConverter.DxaToPoints(left)
        );

    /// <summary>
    /// Calcola la larghezza del contenuto (pagina - margini orizzontali).
    /// </summary>
    public float GetContentWidth(float pageWidth) => pageWidth - Left - Right;

    /// <summary>
    /// Calcola l'altezza del contenuto (pagina - margini verticali).
    /// </summary>
    public float GetContentHeight(float pageHeight) => pageHeight - Top - Bottom;
}
