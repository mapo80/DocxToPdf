using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxToPdf.Sdk.Docx.Styles;
using System.IO;

namespace DocxToPdf.Tests.Helpers;

internal static class OpenXmlTestHelper
{
    public static MemoryStream CreateStyledDocumentStream()
    {
        var stream = new MemoryStream();

        using (var wordDoc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            AddStylesPart(mainPart);
            AddThemePart(mainPart);
            AddSettingsPart(mainPart);

            BuildBody(mainPart.Document.Body!);
        }

        stream.Position = 0;
        return stream;
    }

    public static MemoryStream CreateNumberedDocumentStream()
    {
        var stream = new MemoryStream();
        using (var wordDoc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            AddStylesPart(mainPart);

            var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = BuildNumberingDefinition();

            BuildNumberedBody(mainPart.Document.Body!);
        }

        stream.Position = 0;
        return stream;
    }

    public static ThemeFontScheme LoadThemeFontScheme(string majorFont, string minorFont)
    {
        var stream = new MemoryStream();
        using var wordDoc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var mainPart = wordDoc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("tmp")))));

        var themePart = mainPart.AddNewPart<ThemePart>();
        var themeXml =
            $"""
            <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Custom">
              <a:themeElements>
                <a:clrScheme name="Custom">
                  <a:dk1><a:srgbClr val="000000"/></a:dk1>
                  <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                  <a:dk2><a:srgbClr val="111111"/></a:dk2>
                  <a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
                  <a:accent1><a:srgbClr val="123456"/></a:accent1>
                  <a:accent2><a:srgbClr val="654321"/></a:accent2>
                  <a:accent3><a:srgbClr val="ABCDEF"/></a:accent3>
                  <a:accent4><a:srgbClr val="FEDCBA"/></a:accent4>
                  <a:accent5><a:srgbClr val="00FF00"/></a:accent5>
                  <a:accent6><a:srgbClr val="FF00FF"/></a:accent6>
                  <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
                  <a:folHlink><a:srgbClr val="FF0000"/></a:folHlink>
                </a:clrScheme>
                <a:fontScheme name="CustomFonts">
                  <a:majorFont>
                    <a:latin typeface="{majorFont}"/>
                    <a:ea typeface=""/>
                    <a:cs typeface=""/>
                  </a:majorFont>
                  <a:minorFont>
                    <a:latin typeface="{minorFont}"/>
                    <a:ea typeface=""/>
                    <a:cs typeface=""/>
                  </a:minorFont>
                </a:fontScheme>
                <a:fmtScheme name="CustomFmt">
                  <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
                  <a:lnStyleLst><a:ln w="9525"/></a:lnStyleLst>
                  <a:effectStyleLst><a:effectStyle/></a:effectStyleLst>
                  <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
                </a:fmtScheme>
              </a:themeElements>
              <a:objectDefaults/>
              <a:extraClrSchemeLst/>
            </a:theme>
            """;

        using (var writer = new StreamWriter(themePart.GetStream(FileMode.Create, FileAccess.Write)))
        {
            writer.Write(themeXml);
        }

        return DocxToPdf.Sdk.Docx.Styles.ThemeFontScheme.Load(mainPart.ThemePart);
    }

    public static DocxToPdf.Sdk.Docx.Styles.ThemeColorPalette LoadThemeColorPalette(string accentHex)
    {
        var stream = new MemoryStream();
        using var wordDoc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var mainPart = wordDoc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("tmp")))));

        var themePart = mainPart.AddNewPart<ThemePart>();
        var themeXml =
            $"""
            <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Palette">
              <a:themeElements>
                <a:clrScheme name="Palette">
                  <a:dk1><a:srgbClr val="000000"/></a:dk1>
                  <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                  <a:dk2><a:srgbClr val="222222"/></a:dk2>
                  <a:lt2><a:srgbClr val="DDDDDD"/></a:lt2>
                  <a:accent1><a:srgbClr val="{accentHex}"/></a:accent1>
                  <a:accent2><a:srgbClr val="00FFFF"/></a:accent2>
                  <a:accent3><a:srgbClr val="FF0000"/></a:accent3>
                  <a:accent4><a:srgbClr val="00FF00"/></a:accent4>
                  <a:accent5><a:srgbClr val="0000FF"/></a:accent5>
                  <a:accent6><a:srgbClr val="800080"/></a:accent6>
                  <a:hlink><a:srgbClr val="0088CC"/></a:hlink>
                  <a:folHlink><a:srgbClr val="9900CC"/></a:folHlink>
                </a:clrScheme>
                <a:fontScheme name="Placeholder">
                  <a:majorFont><a:latin typeface="Times New Roman"/></a:majorFont>
                  <a:minorFont><a:latin typeface="Arial"/></a:minorFont>
                </a:fontScheme>
                <a:fmtScheme name="Fmt">
                  <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
                  <a:lnStyleLst><a:ln w="9525"/></a:lnStyleLst>
                  <a:effectStyleLst><a:effectStyle/></a:effectStyleLst>
                  <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
                </a:fmtScheme>
              </a:themeElements>
              <a:objectDefaults/>
              <a:extraClrSchemeLst/>
            </a:theme>
            """;

        using (var writer = new StreamWriter(themePart.GetStream(FileMode.Create, FileAccess.Write)))
        {
            writer.Write(themeXml);
        }

        return DocxToPdf.Sdk.Docx.Styles.ThemeColorPalette.Load(mainPart.ThemePart);
    }

    private static void AddStylesPart(MainDocumentPart mainPart)
    {
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles(
            new DocDefaults(
                new RunPropertiesDefault(
                    new RunProperties(
                        new RunFonts
                        {
                            AsciiTheme = ThemeFontValues.MinorAscii,
                            HighAnsiTheme = ThemeFontValues.MinorHighAnsi
                        },
                        new FontSize { Val = "24" },
                        new Color { ThemeColor = ThemeColorValues.Text1 }
                    )
                ),
                new ParagraphPropertiesDefault(
                    new ParagraphProperties(
                        new SpacingBetweenLines { After = "200", Line = "276" }
                    )
                )
            ),
            new Style
            {
                Type = StyleValues.Paragraph,
                StyleId = "Normal",
                Default = true,
                StyleName = new StyleName { Val = "Normal" }
            },
            new Style
            {
                Type = StyleValues.Paragraph,
                StyleId = "Heading1",
                BasedOn = new BasedOn { Val = "Normal" },
                LinkedStyle = new LinkedStyle { Val = "Heading1Char" },
                StyleRunProperties = new StyleRunProperties(
                    new RunFonts
                    {
                        AsciiTheme = ThemeFontValues.MajorAscii,
                        HighAnsiTheme = ThemeFontValues.MajorHighAnsi
                    },
                    new Bold()
                ),
                StyleParagraphProperties = new StyleParagraphProperties(
                    new Indentation { Left = "720", FirstLine = "360" },
                    new SpacingBetweenLines { Before = "240", After = "120" },
                    new Justification { Val = JustificationValues.Left }
                )
            },
            new Style
            {
                Type = StyleValues.Character,
                StyleId = "Heading1Char",
                LinkedStyle = new LinkedStyle { Val = "Heading1" },
                StyleRunProperties = new StyleRunProperties(
                    new Italic()
                )
            },
            new Style
            {
                Type = StyleValues.Character,
                StyleId = "AccentChar",
                StyleRunProperties = new StyleRunProperties(
                    new Color { ThemeColor = ThemeColorValues.Accent1 }
                )
            }
        );
    }

    private static void AddThemePart(MainDocumentPart mainPart)
    {
        var themePart = mainPart.AddNewPart<ThemePart>();
        const string themeXml =
            """
            <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="UnitTestTheme">
              <a:themeElements>
                <a:clrScheme name="UnitTestTheme">
                  <a:dk1><a:srgbClr val="111111"/></a:dk1>
                  <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                  <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
                  <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
                  <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
                  <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
                  <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
                  <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
                  <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
                  <a:accent6><a:srgbClr val="F79646"/></a:accent6>
                  <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
                  <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
                </a:clrScheme>
                <a:fontScheme name="UnitTestFonts">
                  <a:majorFont>
                    <a:latin typeface="Contoso Headings"/>
                    <a:ea typeface=""/>
                    <a:cs typeface=""/>
                  </a:majorFont>
                  <a:minorFont>
                    <a:latin typeface="Contoso Body"/>
                    <a:ea typeface=""/>
                    <a:cs typeface=""/>
                  </a:minorFont>
                </a:fontScheme>
                <a:fmtScheme name="Fmt">
                  <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
                  <a:lnStyleLst><a:ln w="9525"/></a:lnStyleLst>
                  <a:effectStyleLst><a:effectStyle/></a:effectStyleLst>
                  <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
                </a:fmtScheme>
              </a:themeElements>
              <a:objectDefaults/>
              <a:extraClrSchemeLst/>
            </a:theme>
            """;

        using var writer = new StreamWriter(themePart.GetStream(FileMode.Create, FileAccess.Write));
        writer.Write(themeXml);
    }

    private static void AddSettingsPart(MainDocumentPart mainPart)
    {
        var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings = new Settings(
            new ColorSchemeMapping
            {
                Background1 = ColorSchemeIndexValues.Light1,
                Text1 = ColorSchemeIndexValues.Dark2,
                Background2 = ColorSchemeIndexValues.Light2,
                Text2 = ColorSchemeIndexValues.Dark1,
                Accent1 = ColorSchemeIndexValues.Accent1,
                Accent2 = ColorSchemeIndexValues.Accent2,
                Accent3 = ColorSchemeIndexValues.Accent3,
                Accent4 = ColorSchemeIndexValues.Accent4,
                Accent5 = ColorSchemeIndexValues.Accent5,
                Accent6 = ColorSchemeIndexValues.Accent6,
                Hyperlink = ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = ColorSchemeIndexValues.FollowedHyperlink
            }
        );
    }

    private static void BuildBody(Body body)
    {
        var paragraph = new Paragraph(
            new ParagraphProperties(
                new ParagraphStyleId { Val = "Heading1" },
                new SpacingBetweenLines { Before = "360", After = "120" },
                new Indentation { Left = "720", FirstLine = "360" },
                new Justification { Val = JustificationValues.Center }
            ),
            BuildRun("Plain linked", null),
            BuildRun(" Tint", new RunProperties(new Color
            {
                ThemeColor = ThemeColorValues.Accent2,
                ThemeTint = "99"
            })),
            BuildRun(" Hex", new RunProperties(new Color { Val = "3366FF" })),
            BuildRun(" Styled", new RunProperties(new RunStyle { Val = "AccentChar" }))
        );

        body.Append(paragraph);
    }

    private static Run BuildRun(string text, RunProperties? runProperties)
    {
        var run = new Run();
        if (runProperties != null)
            run.RunProperties = (RunProperties)runProperties.CloneNode(true);
        run.Append(new Text(text));
        return run;
    }

    private static Numbering BuildNumberingDefinition()
    {
        return new Numbering(
            new AbstractNum(
                new Level(
                    new StartNumberingValue { Val = 1 },
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = "%1." },
                    new LevelSuffix { Val = LevelSuffixValues.Tab },
                    new LevelJustification { Val = LevelJustificationValues.Left },
                    new ParagraphProperties(new Indentation { Left = "720", Hanging = "360" })
                ) { LevelIndex = 0 },
                new Level(
                    new StartNumberingValue { Val = 1 },
                    new NumberingFormat { Val = NumberFormatValues.LowerLetter },
                    new LevelText { Val = "%1.%2" },
                    new LevelSuffix { Val = LevelSuffixValues.Space },
                    new LevelJustification { Val = LevelJustificationValues.Left },
                    new ParagraphProperties(new Indentation { Left = "1080", Hanging = "360" })
                ) { LevelIndex = 1 },
                new Level(
                    new NumberingFormat { Val = NumberFormatValues.Bullet },
                    new LevelText { Val = "•" },
                    new LevelSuffix { Val = LevelSuffixValues.Tab },
                    new LevelJustification { Val = LevelJustificationValues.Left },
                    new ParagraphProperties(new Indentation { Left = "1440", Hanging = "360" }),
                    new NumberingSymbolRunProperties(new RunFonts { Ascii = "Wingdings", HighAnsi = "Wingdings" })
                ) { LevelIndex = 2 }
            ) { AbstractNumberId = 1 },
            new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 },
            new NumberingInstance(
                new AbstractNumId { Val = 1 },
                new LevelOverride(new StartOverrideNumberingValue { Val = 4 }) { LevelIndex = 0 }
            ) { NumberID = 2 }
        );
    }

    private static void BuildNumberedBody(Body body)
    {
        AppendNumberedParagraph(body, "Kickoff", 1, 0);
        AppendNumberedParagraph(body, "Preparation", 1, 1);
        AppendNumberedParagraph(body, "Assets", 1, 1);
        AppendNumberedParagraph(body, "Checklist", 1, 2);
        AppendNumberedParagraph(body, "Timeline", 1, 0);

        body.Append(new Paragraph(new Run(new Text("Separator"))));

        AppendNumberedParagraph(body, "Restarted", 2, 0);
        AppendNumberedParagraph(body, "Continued", 1, 0);
    }

    private static void AppendNumberedParagraph(Body body, string text, int numId, int level)
    {
        var numProps = new NumberingProperties(
            new NumberingLevelReference { Val = level },
            new NumberingId { Val = numId }
        );

        var paragraph = new Paragraph(
            new ParagraphProperties(numProps),
            new Run(new Text(text))
        );

        body.Append(paragraph);
    }
}
