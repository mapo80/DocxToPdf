using DocxToPdf.Sdk;
using DocxToPdf.Sdk.Pdf;
using DocxToPdf.Sdk.Text;
using DocxToPdf.Sdk.Units;
using SkiaSharp;

namespace DocxToPdf.Demo;

/// <summary>
/// Demo DocxToPdf - supporta sia "Hello, World" che conversione DOCX → PDF.
/// </summary>
class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("DocxToPdf.Sdk - Demo\n");

        // Modalità: comando dedicato oppure scorciatoia .docx
        if (args.Length >= 1 && string.Equals(args[0], "render", StringComparison.OrdinalIgnoreCase))
        {
            HandleRenderCommand(args);
            return;
        }

        if (args.Length >= 1 && args[0].EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
        {
            // Modalità conversione DOCX → PDF
            var docxPath = args[0];
            var pdfPath = args.Length >= 2 ? args[1] : Path.ChangeExtension(docxPath, ".pdf");

        ConvertDocxToPdf(docxPath, pdfPath);
            return;
        }

        // Modalità "Hello, World" (default)
        var helloOutputPath = Path.Combine(Environment.CurrentDirectory, "hello.pdf");
        Console.WriteLine($"Generando PDF: {helloOutputPath}");

        try
        {
            GenerateHelloPdf(helloOutputPath);
            Console.WriteLine("\n✓ PDF generato con successo!");
            Console.WriteLine($"  Percorso: {helloOutputPath}");
            Console.WriteLine($"  Dimensione: {new FileInfo(helloOutputPath).Length:N0} bytes");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"\n✗ Errore durante la generazione: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
            Environment.Exit(1);
        }
    }

    static void HandleRenderCommand(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Uso: DocxToPdf.Demo render <input.docx> [-o <output.pdf>]");
            Environment.Exit(1);
        }

        var inputPath = args[1];
        string? outputPath = null;
        var enableDiagnostics = false;

        for (int i = 2; i < args.Length; i++)
        {
            var current = args[i];
            if (string.Equals(current, "-o", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(current, "--output", StringComparison.OrdinalIgnoreCase))
            {
                if (i + 1 >= args.Length)
                {
                    Console.WriteLine("Manca il percorso dopo -o/--output.");
                    Environment.Exit(1);
                }
                outputPath = args[++i];
            }
            else if (string.Equals(current, "--log-numbering", StringComparison.OrdinalIgnoreCase))
            {
                enableDiagnostics = true;
            }
        }

        outputPath ??= Path.ChangeExtension(inputPath, ".pdf");
        var converter = new DocxToPdfConverter
        {
            DiagnosticsLogger = enableDiagnostics ? message => Console.WriteLine(message) : null
        };
        ConvertDocxToPdf(inputPath, outputPath, converter);
    }

    static void ConvertDocxToPdf(string docxPath, string pdfPath) =>
        ConvertDocxToPdf(docxPath, pdfPath, null);

    static void ConvertDocxToPdf(string docxPath, string pdfPath, DocxToPdfConverter? existingConverter)
    {
        Console.WriteLine($"Conversione DOCX → PDF");
        Console.WriteLine($"  Input:  {docxPath}");
        Console.WriteLine($"  Output: {pdfPath}\n");

        if (!File.Exists(docxPath))
        {
            Console.WriteLine($"✗ File non trovato: {docxPath}");
            Environment.Exit(1);
        }

        var outputDir = Path.GetDirectoryName(pdfPath);
        if (!string.IsNullOrEmpty(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        try
        {
            var converter = existingConverter ?? new DocxToPdfConverter();
            converter.Convert(docxPath, pdfPath);

            Console.WriteLine("\n✓ Conversione completata!");
            Console.WriteLine($"  Percorso: {pdfPath}");
            Console.WriteLine($"  Dimensione: {new FileInfo(pdfPath).Length:N0} bytes");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"\n✗ Errore durante la conversione: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
            Environment.Exit(1);
        }
    }

    static void GenerateHelloPdf(string outputPath)
    {
        // Metadati del documento
        var metadata = new PdfMetadata
        {
            Title = "Hello, World Demo",
            Author = "DocxToPdf.Sdk",
            Creator = "DocxToPdf.Sdk v0.1.0-alpha",
            Subject = "Dimostrazione rendering testo con SkiaSharp + HarfBuzz",
            CreationDate = DateTime.Now
        };

        // Crea il documento PDF
        using var builder = PdfDocumentBuilder.Create(outputPath, metadata);

        // Font e renderer di testo
        var fontManager = FontManager.Instance;
        var textRenderer = new TextRenderer();

        // Parametri tipografici
        var typeface = fontManager.GetDefaultTypeface();
        var fontSize = 48f; // pt
        var textColor = SKColors.Black;

        // Inizia una pagina A4
        var page = builder.BeginPage(PaperSize.A4);
        var margins = Margins.Default; // 72 pt = 1 inch su tutti i lati

        // Calcola posizione centrata orizzontalmente, alto verticalmente
        var contentWidth = margins.GetContentWidth(page.Size.WidthPt);
        var centerX = margins.Left + (contentWidth / 2f);
        var topY = margins.Top + fontSize; // baseline = top margin + font size

        // Disegna "Hello, world!" con text shaping
        Console.WriteLine("\nRenderizzando testo con HarfBuzz shaping...");
        textRenderer.DrawCenteredText(
            page.Canvas,
            "Hello, world!",
            centerX,
            topY,
            contentWidth,
            typeface,
            fontSize,
            textColor
        );

        // Disegna testo aggiuntivo con caratteri speciali per testare lo shaping
        var testTexts = new[]
        {
            ("Testo con àccénti", 24f),
            ("Legatures: fi fl ffi ffl", 20f),
            ("Math: ∫∑∏√∞≠≤≥", 18f),
            ("Emoji: 👋🌍✨", 16f)
        };

        var currentY = topY + 80f;

        foreach (var (text, size) in testTexts)
        {
            textRenderer.DrawCenteredText(
                page.Canvas,
                text,
                centerX,
                currentY,
                contentWidth,
                typeface,
                size,
                SKColors.DarkGray
            );
            currentY += size + 20f;
        }

        // Informazioni tecniche in basso
        var infoY = page.Size.HeightPt - margins.Bottom - 40f;
        var infoFont = fontManager.GetTypeface("Courier New", SKFontStyle.Normal);
        var infoSize = 10f;

        var infoTexts = new[]
        {
            $"Pagina: A4 ({page.Size.WidthPt:F1} × {page.Size.HeightPt:F1} pt)",
            $"Margini: {margins.Top} pt (= {UnitConverter.PointsToMm(margins.Top):F1} mm)",
            $"Font: {typeface.FamilyName} @ {fontSize} pt",
            "Engine: SkiaSharp + HarfBuzz text shaping"
        };

        foreach (var info in infoTexts)
        {
            textRenderer.DrawShapedText(
                page.Canvas,
                info,
                margins.Left,
                infoY,
                infoFont,
                infoSize,
                SKColors.Gray
            );
            infoY += infoSize + 4f;
        }

        // Termina la pagina e chiudi il documento
        builder.EndPage();
        builder.Close();
    }
}
