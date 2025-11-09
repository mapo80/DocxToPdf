namespace DocxToPdf.Sdk.Units;

/// <summary>
/// Convertitore di unità di misura per layout PDF e DOCX.
/// Unità base: punti tipografici (pt), dove 1 pt = 1/72 inch.
/// </summary>
public static class UnitConverter
{
    // Costanti fondamentali
    public const float PointsPerInch = 72f;
    public const float InchesPerMm = 1f / 25.4f;
    public const float InchesPerCm = 1f / 2.54f;

    // Unità DOCX: dxa (twentieth of a point)
    // 1 pt = 20 dxa, quindi 1 dxa = 0.05 pt
    public const float DxaPerPoint = 20f;
    public const float PointsPerDxa = 1f / DxaPerPoint;

    /// <summary>
    /// Converte millimetri in punti tipografici.
    /// </summary>
    public static float MmToPoints(float mm) => mm * InchesPerMm * PointsPerInch;

    /// <summary>
    /// Converte centimetri in punti tipografici.
    /// </summary>
    public static float CmToPoints(float cm) => cm * InchesPerCm * PointsPerInch;

    /// <summary>
    /// Converte pollici in punti tipografici.
    /// </summary>
    public static float InchesToPoints(float inches) => inches * PointsPerInch;

    /// <summary>
    /// Converte punti tipografici in millimetri.
    /// </summary>
    public static float PointsToMm(float points) => points / PointsPerInch / InchesPerMm;

    /// <summary>
    /// Converte DXA (twentieths of a point, unità WordprocessingML) in punti tipografici.
    /// Esempio: margine Word di 1440 dxa = 72 pt = 1 inch.
    /// </summary>
    public static float DxaToPoints(int dxa) => dxa * PointsPerDxa;

    /// <summary>
    /// Converte punti tipografici in DXA.
    /// </summary>
    public static int PointsToDxa(float points) => (int)(points * DxaPerPoint);
}
