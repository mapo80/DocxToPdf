using DocxToPdf.Sdk.Docx.Formatting;

namespace DocxToPdf.Sdk.Docx.Styles;

internal sealed class ParagraphStyleContext
{
    private readonly RunPropertySet _runBase;
    private ParagraphFormatting? _formatting;

    public ParagraphStyleContext(string paragraphStyleId, ParagraphPropertySet paragraphProperties, RunPropertySet runBase)
    {
        ParagraphStyleId = paragraphStyleId;
        ParagraphProperties = paragraphProperties;
        _runBase = runBase;
    }

    public string ParagraphStyleId { get; }
    public ParagraphPropertySet ParagraphProperties { get; }

    public ParagraphFormatting ParagraphFormatting => _formatting ??= ParagraphProperties.ToParagraphFormatting();

    public RunPropertySet CreateRunPropertySet() => _runBase.Clone();
}
