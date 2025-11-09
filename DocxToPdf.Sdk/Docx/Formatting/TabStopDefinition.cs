namespace DocxToPdf.Sdk.Docx.Formatting;

public enum TabAlignment
{
    Left,
    Center,
    Right,
    Decimal,
    Bar
}

public enum TabLeader
{
    None,
    Dots,
    Dashes,
    Line,
    ThickLine,
    Underscore,
    Heavy
}

public sealed record TabStopDefinition(
    float PositionPt,
    TabAlignment Alignment,
    TabLeader Leader);
