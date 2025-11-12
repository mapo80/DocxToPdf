namespace DocxToPdf.Sdk.Docx;

/// <summary>
/// Rappresenta un blocco di contenuto nel body del documento (paragrafo o tabella).
/// </summary>
internal abstract record DocxBlock;

internal sealed record DocxParagraphBlock(DocxParagraph Paragraph) : DocxBlock;

internal sealed record DocxTableBlock(DocxTable Table) : DocxBlock;
