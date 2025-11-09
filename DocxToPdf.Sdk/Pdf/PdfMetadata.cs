namespace DocxToPdf.Sdk.Pdf;

/// <summary>
/// Metadati del documento PDF (XMP).
/// </summary>
public sealed record PdfMetadata
{
    public string? Title { get; init; }
    public string? Author { get; init; }
    public string? Subject { get; init; }
    public string? Keywords { get; init; }
    public string? Creator { get; init; }
    public DateTime? CreationDate { get; init; }

    public static readonly PdfMetadata Empty = new();
}
