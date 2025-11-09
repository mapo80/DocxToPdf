using DocumentFormat.OpenXml.Wordprocessing;
using DocxToPdf.Sdk.Pdf;
using DocxToPdf.Sdk.Units;

namespace DocxToPdf.Sdk.Docx;

/// <summary>
/// Rappresenta le proprietà di sezione DOCX (pagina e margini).
/// </summary>
public sealed record DocxSection
{
    public PaperSize PageSize { get; init; }
    public Margins Margins { get; init; }
    public PageOrientationValues Orientation { get; init; }

    /// <summary>
    /// Sezione di default: A4 portrait con margini 1" (compatibile Word).
    /// </summary>
    public static readonly DocxSection Default = new()
    {
        PageSize = PaperSize.A4,
        Margins = Units.Margins.Default,
        Orientation = PageOrientationValues.Portrait
    };

    /// <summary>
    /// Crea una DocxSection dalle SectionProperties del DOCX.
    /// </summary>
    public static DocxSection FromSectionProperties(SectionProperties sectPr)
    {
        var pageSize = ParsePageSize(sectPr);
        var margins = ParseMargins(sectPr);
        var orientation = ParseOrientation(sectPr);

        return new DocxSection
        {
            PageSize = pageSize,
            Margins = margins,
            Orientation = orientation
        };
    }

    private static PaperSize ParsePageSize(SectionProperties sectPr)
    {
        var pgSz = sectPr.GetFirstChild<PageSize>();
        if (pgSz == null)
            return PaperSize.A4;

        // PageSize in DOCX è in twips (1/20 pt, 1440 twips = 1 inch = 72 pt)
        // Width/Height possono essere uint?, dobbiamo gestire null
        var widthTwips = pgSz.Width?.Value ?? 11906; // A4 default: 8.27" * 1440
        var heightTwips = pgSz.Height?.Value ?? 16838; // A4 default: 11.69" * 1440

        var widthPt = TwipsToPoints(widthTwips);
        var heightPt = TwipsToPoints(heightTwips);

        return new PaperSize(widthPt, heightPt);
    }

    private static Margins ParseMargins(SectionProperties sectPr)
    {
        var pgMar = sectPr.GetFirstChild<PageMargin>();
        if (pgMar == null)
            return Units.Margins.Default;

        // Margini in DOCX sono in DXA (twentieths of a point)
        // 1440 dxa = 72 pt = 1 inch
        var topDxa = pgMar.Top?.Value ?? 1440;
        var rightDxa = pgMar.Right?.Value ?? 1440;
        var bottomDxa = pgMar.Bottom?.Value ?? 1440;
        var leftDxa = pgMar.Left?.Value ?? 1440;

        return Units.Margins.FromDxa(
            (int)topDxa,
            (int)rightDxa,
            (int)bottomDxa,
            (int)leftDxa
        );
    }

    private static PageOrientationValues ParseOrientation(SectionProperties sectPr)
    {
        var pgSz = sectPr.GetFirstChild<PageSize>();
        return pgSz?.Orient?.Value ?? PageOrientationValues.Portrait;
    }

    /// <summary>
    /// Converte twips (1/20 pt) in punti tipografici.
    /// </summary>
    private static float TwipsToPoints(uint twips) => twips / 20f;
}
