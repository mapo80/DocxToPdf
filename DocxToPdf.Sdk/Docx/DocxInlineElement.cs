using DocumentFormat.OpenXml.Wordprocessing;
using DocxToPdf.Sdk.Docx.Formatting;

namespace DocxToPdf.Sdk.Docx;

public abstract record DocxInlineElement;

public sealed record DocxTextInline(string Text, RunFormatting Formatting) : DocxInlineElement;

public sealed record DocxTabInline(RunFormatting Formatting) : DocxInlineElement;

public enum PositionalTabReference
{
    Margin,
    Indent,
    Page
}

public sealed record DocxPositionalTabInline(
    RunFormatting Formatting,
    float PositionPt,
    TabAlignment Alignment,
    TabLeader Leader,
    PositionalTabReference Reference) : DocxInlineElement;
