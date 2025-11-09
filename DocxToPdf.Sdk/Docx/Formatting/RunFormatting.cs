namespace DocxToPdf.Sdk.Docx.Formatting;

/// <summary>
/// Proprietà effettive di un run di testo dopo aver risolto stili e formattazioni dirette.
/// </summary>
public sealed record RunFormatting
{
    public static readonly RunFormatting Default = new()
    {
        FontFamily = "Arial",
        FontSizePt = 11f,
        Color = RgbColor.Black,
        CharacterSpacingPt = 0f,
        KerningEnabled = false
    };

    public string FontFamily { get; init; } = "Arial";
    public float FontSizePt { get; init; } = 11f;
    public bool Bold { get; init; }
    public bool Italic { get; init; }
    public bool Underline { get; init; }
    public bool Strike { get; init; }
    public bool SmallCaps { get; init; }
    public float CharacterSpacingPt { get; init; }
    public bool KerningEnabled { get; init; }
    public RgbColor Color { get; init; } = RgbColor.Black;
}
