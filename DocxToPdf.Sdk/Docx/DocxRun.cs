using DocumentFormat.OpenXml.Wordprocessing;
using DocxToPdf.Sdk.Docx.Formatting;
using System.Linq;
using System.Text;

namespace DocxToPdf.Sdk.Docx;

/// <summary>
/// Rappresenta un run di testo DOCX con il relativo stile risolto.
/// </summary>
public sealed record DocxRun
{
    public string Text { get; init; } = string.Empty;
    public RunFormatting Formatting { get; init; } = RunFormatting.Default;

    public static DocxRun? FromRun(Run run, RunFormatting formatting)
    {
        var textElements = run.Elements<DocumentFormat.OpenXml.Wordprocessing.Text>();
        if (!textElements.Any())
            return null;

        var sb = new StringBuilder();
        foreach (var textElement in textElements)
        {
            sb.Append(textElement.InnerText);
        }

        var text = sb.ToString();
        if (string.IsNullOrEmpty(text))
            return null;

        return new DocxRun
        {
            Text = text,
            Formatting = formatting
        };
    }
}
