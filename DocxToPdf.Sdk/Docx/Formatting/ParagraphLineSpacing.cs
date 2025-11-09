namespace DocxToPdf.Sdk.Docx.Formatting;

public enum ParagraphLineSpacingRule
{
    Auto,
    Exact,
    AtLeast
}

public readonly record struct ParagraphLineSpacing(ParagraphLineSpacingRule Rule, float Value)
{
    public float Resolve(float defaultSpacing) =>
        Rule switch
        {
            ParagraphLineSpacingRule.Auto => defaultSpacing * Value,
            ParagraphLineSpacingRule.Exact => Value,
            ParagraphLineSpacingRule.AtLeast => Math.Max(defaultSpacing, Value),
            _ => defaultSpacing
        };

    public static ParagraphLineSpacing Auto(float multiple) =>
        new(ParagraphLineSpacingRule.Auto, multiple);

    public static ParagraphLineSpacing Exact(float points) =>
        new(ParagraphLineSpacingRule.Exact, points);

    public static ParagraphLineSpacing AtLeast(float points) =>
        new(ParagraphLineSpacingRule.AtLeast, points);
}
