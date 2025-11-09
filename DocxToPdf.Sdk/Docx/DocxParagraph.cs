using DocumentFormat.OpenXml.Wordprocessing;
using DocxToPdf.Sdk.Docx.Formatting;
using DocxToPdf.Sdk.Docx.Numbering;
using DocxToPdf.Sdk.Docx.Styles;
using System.Collections.Generic;
using System.Linq;

namespace DocxToPdf.Sdk.Docx;

/// <summary>
/// Paragrafo DOCX con testo e formattazioni risolte.
/// </summary>
public sealed record DocxParagraph
{
    public ParagraphFormatting ParagraphFormatting { get; init; } = ParagraphFormatting.Default;
    public IReadOnlyList<DocxRun> Runs { get; init; } = [];
    public DocxListMarker? ListMarker { get; init; }
    public float DefaultTabStopPt { get; init; }

    internal static DocxParagraph FromParagraph(Paragraph paragraph, DocxStyleResolver styleResolver, NumberingResolver numberingResolver, float defaultTabStopPt)
    {
        var context = styleResolver.CreateParagraphContext(paragraph);
        var numberingResult = numberingResolver.Resolve(paragraph, context.ParagraphProperties, styleResolver);
        var runs = new List<DocxRun>();

        foreach (var run in paragraph.Descendants<Run>())
        {
            var runFormatting = styleResolver.ResolveRunFormatting(context, run);
            var docxRun = DocxRun.FromRun(run, runFormatting);
            if (docxRun != null)
                runs.Add(docxRun);
        }

        return new DocxParagraph
        {
            ParagraphFormatting = context.ParagraphFormatting,
            Runs = runs,
            ListMarker = numberingResult?.Marker,
            DefaultTabStopPt = defaultTabStopPt
        };
    }

    public string GetFullText() => string.Concat(Runs.Select(r => r.Text));
}
