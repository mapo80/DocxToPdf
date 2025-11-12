using PdfVisualDiff.Core;
using System.CommandLine;

var baselineArg = new Argument<FileInfo>("baseline", "Baseline PDF (A)");
var candidateArg = new Argument<FileInfo>("candidate", "Candidate PDF (B)");
var outputOption = new Option<DirectoryInfo>("--out")
{
    Description = "Output directory (will contain A/, B/, diff/, thumbs/, report.csv, index.html)",
    IsRequired = true
};
var dpiOption = new Option<int>("--dpi", () => 300, "Rasterization DPI passed to pdftoppm");
var pageRangeOption = new Option<string?>("--page-range", "Page selection (e.g. 1-3,5). Defaults to overlap between PDFs.");
var cropBoxOption = new Option<bool>("--cropbox", () => false, "Use CropBox when rasterizing (helps align printable area).");
var fuzzOption = new Option<string?>("--fuzz", "Optional fuzz tolerance for ImageMagick compare (e.g. 0.1% or 2).");
var baselineUpwOption = new Option<string?>("--baseline-user-password", "User password for the baseline PDF (pdftoppm -upw).");
var baselineOpwOption = new Option<string?>("--baseline-owner-password", "Owner password for the baseline PDF (pdftoppm -opw).");
var candidateUpwOption = new Option<string?>("--candidate-user-password", "User password for the candidate PDF.");
var candidateOpwOption = new Option<string?>("--candidate-owner-password", "Owner password for the candidate PDF.");
var thumbnailOption = new Option<int>("--thumbnail-size", () => 280, "Max thumbnail dimension in pixels.");
var pdftoppmPathOption = new Option<string?>("--pdftoppm-path", "Path to the pdftoppm executable.");
var comparePathOption = new Option<string?>("--compare-path", "Path to ImageMagick compare.");
var identifyPathOption = new Option<string?>("--identify-path", "Path to ImageMagick identify.");
var magickPathOption = new Option<string?>("--magick-path", "Path to ImageMagick magick/convert for thumbnails.");
var ghostscriptPathOption = new Option<string?>("--ghostscript-path", "Path to the Ghostscript executable (gs).");
var pdfboxDirOption = new Option<DirectoryInfo?>("--pdfbox-dir", "Directory containing pdfbox-app-*.jar and geometry-extractor.jar.");
var geometryOption = new Option<bool>("--geometry-check", () => false, "Run PdfPig geometry comparison after pixel diff.");
var geometryToleranceOption = new Option<double>("--geometry-tolerance", () => 0.25, "Geometry tolerance in points (default 0.25pt).");
var warningAePercentOption = new Option<double>("--warning-ae-percent", () => 0.005, "Warning threshold for AE percent (default 0.005 = 0.005%).");
var warningSsimOption = new Option<double>("--warning-ssim", () => 0.995, "Warning threshold for SSIM (default 0.995).");
var ssimRadiusOption = new Option<double?>("--ssim-radius", "Sets -define compare:ssim-radius=<value> when computing SSIM/DSSIM.");
var ssimSigmaOption = new Option<double?>("--ssim-sigma", "Sets -define compare:ssim-sigma=<value> when computing SSIM/DSSIM.");
var ssimK1Option = new Option<double?>("--ssim-k1", "Sets -define compare:ssim-k1=<value> when computing SSIM/DSSIM.");
var ssimK2Option = new Option<double?>("--ssim-k2", "Sets -define compare:ssim-k2=<value> when computing SSIM/DSSIM.");

var root = new RootCommand("pdf-visual-diff — pdftoppm + ImageMagick pixel diff workflow")
{
    baselineArg,
    candidateArg,
    outputOption,
    dpiOption,
    pageRangeOption,
    cropBoxOption,
    fuzzOption,
    baselineUpwOption,
    baselineOpwOption,
    candidateUpwOption,
    candidateOpwOption,
    thumbnailOption,
    pdftoppmPathOption,
    comparePathOption,
    identifyPathOption,
    magickPathOption,
    ghostscriptPathOption,
    pdfboxDirOption,
    geometryOption,
    geometryToleranceOption,
    warningAePercentOption,
    warningSsimOption,
    ssimRadiusOption,
    ssimSigmaOption,
    ssimK1Option,
    ssimK2Option
};

root.SetHandler(async context =>
{
    var baseline = context.ParseResult.GetValueForArgument(baselineArg);
    var candidate = context.ParseResult.GetValueForArgument(candidateArg);
    var outputDir = context.ParseResult.GetValueForOption(outputOption)!;
    var dpi = context.ParseResult.GetValueForOption(dpiOption);
    var pageRange = context.ParseResult.GetValueForOption(pageRangeOption);
    var useCropBox = context.ParseResult.GetValueForOption(cropBoxOption);
    var fuzz = context.ParseResult.GetValueForOption(fuzzOption);
    var baselinePasswords = new PdfPasswordOptions(
        context.ParseResult.GetValueForOption(baselineUpwOption),
        context.ParseResult.GetValueForOption(baselineOpwOption));
    var candidatePasswords = new PdfPasswordOptions(
        context.ParseResult.GetValueForOption(candidateUpwOption),
        context.ParseResult.GetValueForOption(candidateOpwOption));
    var thumbnailSize = context.ParseResult.GetValueForOption(thumbnailOption);
    var pdftoppmPath = context.ParseResult.GetValueForOption(pdftoppmPathOption);
    var comparePath = context.ParseResult.GetValueForOption(comparePathOption);
    var identifyPath = context.ParseResult.GetValueForOption(identifyPathOption);
    var magickPath = context.ParseResult.GetValueForOption(magickPathOption);
    var ghostscriptPath = context.ParseResult.GetValueForOption(ghostscriptPathOption);
    var pdfboxDir = context.ParseResult.GetValueForOption(pdfboxDirOption);
    var geometry = context.ParseResult.GetValueForOption(geometryOption);
    var geometryTolerance = context.ParseResult.GetValueForOption(geometryToleranceOption);
    var warningAePercent = context.ParseResult.GetValueForOption(warningAePercentOption);
    var warningSsim = context.ParseResult.GetValueForOption(warningSsimOption);
    var ssimRadius = context.ParseResult.GetValueForOption(ssimRadiusOption);
    var ssimSigma = context.ParseResult.GetValueForOption(ssimSigmaOption);
    var ssimK1 = context.ParseResult.GetValueForOption(ssimK1Option);
    var ssimK2 = context.ParseResult.GetValueForOption(ssimK2Option);

    if (dpi <= 0)
    {
        context.ExitCode = 2;
        Console.Error.WriteLine("--dpi must be greater than zero.");
        return;
    }

    if (thumbnailSize < 32)
    {
        context.ExitCode = 2;
        Console.Error.WriteLine("--thumbnail-size must be at least 32 pixels.");
        return;
    }

    if (warningAePercent < 0)
    {
        context.ExitCode = 2;
        Console.Error.WriteLine("--warning-ae-percent must be non-negative.");
        return;
    }

    if (warningSsim <= 0 || warningSsim > 1)
    {
        context.ExitCode = 2;
        Console.Error.WriteLine("--warning-ssim must be within (0, 1].");
        return;
    }

    var parsedPages = PageSelection.Parse(pageRange);
    var normalizedRange = PageSelection.Normalize(pageRange);

    var runner = new VisualDiffRunner(new PixelDiffToolPaths(pdftoppmPath, comparePath, identifyPath, magickPath, ghostscriptPath));
    try
    {
        var request = new VisualDiffRequest(
            baseline.FullName,
            candidate.FullName,
            outputDir.FullName,
            parsedPages,
            normalizedRange ?? pageRange,
            dpi,
            useCropBox,
            fuzz,
            baselinePasswords,
            candidatePasswords,
            thumbnailSize,
            geometry,
            geometryTolerance,
            new QualityThresholds(warningAePercent, warningSsim),
            new SsimOptions(ssimRadius, ssimSigma, ssimK1, ssimK2),
            ghostscriptPath,
            pdfboxDir?.FullName);

        var result = await runner.RunAsync(request, context.GetCancellationToken());

        Console.WriteLine($"CSV report: {result.CsvPath}");
        Console.WriteLine($"HTML report: {result.HtmlPath}");
        Console.WriteLine($"JSON report: {result.JsonPath}");
        Console.WriteLine($"Pages compared: {result.Summary.ComparedPages.Count}");
        Console.WriteLine($"Total AE pixels: {result.Summary.TotalAePixels}");
        Console.WriteLine($"Weighted AE %: {result.Summary.WeightedAePercent:F6}");
        Console.WriteLine($"Overall status: {result.Summary.OverallStatus}");

        if (result.Summary.MissingFromBaseline.Count > 0)
        {
            Console.WriteLine($"Skipped (missing in baseline): {string.Join(", ", result.Summary.MissingFromBaseline)}");
        }

        if (result.Summary.MissingFromCandidate.Count > 0)
        {
            Console.WriteLine($"Skipped (missing in candidate): {string.Join(", ", result.Summary.MissingFromCandidate)}");
        }

        context.ExitCode = result.Summary.OverallStatus == DiffStatus.Fail ? 1 : 0;
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"Error: {ex.Message}");
        context.ExitCode = 2;
    }
});

return await root.InvokeAsync(args);
