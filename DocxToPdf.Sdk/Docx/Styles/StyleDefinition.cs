using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxToPdf.Sdk.Docx.Styles;

internal sealed record StyleDefinition(
    string StyleId,
    StyleValues StyleType,
    bool IsDefault,
    string? BasedOn,
    string? LinkedStyleId,
    ParagraphPropertySet? ParagraphProperties,
    RunPropertySet? RunProperties)
{
    public static StyleDefinition FromStyle(Style style)
    {
        var styleId = style.StyleId?.Value ?? string.Empty;
        var type = style.Type?.Value ?? StyleValues.Paragraph;
        var basedOn = style.BasedOn?.Val?.Value;
        var linkedStyle = style.LinkedStyle?.Val?.Value;
        var isDefault = style.Default?.Value ?? false;

        var paragraphProps = ParagraphPropertySet.FromOpenXml(style.StyleParagraphProperties);
        var runProps = RunPropertySet.FromOpenXml(
            RunPropertyHelpers.CloneRunProperties(style.StyleRunProperties));

        return new StyleDefinition(
            styleId,
            type,
            isDefault,
            basedOn,
            linkedStyle,
            paragraphProps,
            runProps);
    }
}
