using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxToPdf.Sdk.Docx.Formatting;

public sealed record DocxListMarker(
    string Text,
    RunFormatting Formatting,
    ParagraphAlignment Alignment,
    LevelSuffixValues Suffix);
