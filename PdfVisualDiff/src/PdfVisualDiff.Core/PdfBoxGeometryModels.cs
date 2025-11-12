using System.Text.Json.Serialization;

namespace PdfVisualDiff.Core;

internal sealed record PdfBoxDocumentGeometry(
    [property: JsonPropertyName("pages")] IReadOnlyList<PdfBoxPageGeometry> Pages);

internal sealed record PdfBoxPageGeometry(
    [property: JsonPropertyName("page")] int Page,
    [property: JsonPropertyName("words")] IReadOnlyList<PdfBoxWordGeometry> Words,
    [property: JsonPropertyName("graphics")] IReadOnlyList<PdfBoxGraphicGeometry> Graphics);

internal sealed record PdfBoxWordGeometry(
    [property: JsonPropertyName("text")] string Text,
    [property: JsonPropertyName("x")] double X,
    [property: JsonPropertyName("y")] double Y,
    [property: JsonPropertyName("width")] double Width,
    [property: JsonPropertyName("height")] double Height,
    [property: JsonPropertyName("fontSize")] double FontSize);

internal sealed record PdfBoxGraphicGeometry(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("x")] double X,
    [property: JsonPropertyName("y")] double Y,
    [property: JsonPropertyName("width")] double Width,
    [property: JsonPropertyName("height")] double Height,
    [property: JsonPropertyName("strokeWidth")] double StrokeWidth,
    [property: JsonPropertyName("strokeColor")] string? StrokeColor,
    [property: JsonPropertyName("fillColor")] string? FillColor);
