using PdfVisualDiff.Core;
using System.CommandLine;
using System.CommandLine.Invocation;

var baselineArg = new Argument<FileInfo>("baseline", "Baseline PDF");
var candidateArg = new Argument<FileInfo>("candidate", "Candidate PDF");
var outputOption = new Option<DirectoryInfo>("--out", description: "Output directory")
{
    IsRequired = true
};
var dpiOption = new Option<int>("--dpi", () => 300, "Render DPI passed to diff-pdf and used for metrics");
var modeOption = new Option<string>("--mode", () => "overlay", "Diff mode (ignored)");
var thresholdOption = new Option<double>("--threshold", () => 0, "Allowed diff ratio (ignored)");
var toleranceOption = new Option<int>("--antialias-tolerance", () => 0, "AA tolerance (ignored)");
var maxPagesOption = new Option<int?>("--max-pages", description: "Maximum pages to compare (ignored)");
var pageRangeOption = new Option<string?>("--page-range", description: "Specific pages, e.g. 1-3,5");
var masksOption = new Option<FileInfo?>("--masks", description: "Mask JSON (ignored)");
var geometryOption = new Option<bool>("--geometry-check", () => false, "Enable PdfPig geometry comparison");
var geometryToleranceOption = new Option<double>("--geometry-tolerance", () => 0.25, "Geometry tolerance in points");
var parallelOption = new Option<int>("--parallel", () => Math.Max(1, Math.Min(Environment.ProcessorCount, 4)), "Parallelism (ignored)");
var diffPdfPathOption = new Option<string?>("--diff-pdf-path", "Path to diff-pdf executable");

var root = new RootCommand("pdf-visual-diff (diff-pdf wrapper)")
{
    baselineArg,
    candidateArg,
    outputOption,
    dpiOption,
    modeOption,
    thresholdOption,
    toleranceOption,
    maxPagesOption,
    pageRangeOption,
    masksOption,
    geometryOption,
    geometryToleranceOption,
    parallelOption,
    diffPdfPathOption
};

root.SetHandler(async context =>
{
    var baseline = context.ParseResult.GetValueForArgument(baselineArg);
    var candidate = context.ParseResult.GetValueForArgument(candidateArg);
    var outputDir = context.ParseResult.GetValueForOption(outputOption)!;
    var pageRange = context.ParseResult.GetValueForOption(pageRangeOption);
    var dpi = context.ParseResult.GetValueForOption(dpiOption);
    var mode = context.ParseResult.GetValueForOption(modeOption) ?? "overlay";
    var threshold = context.ParseResult.GetValueForOption(thresholdOption);
    var tolerance = context.ParseResult.GetValueForOption(toleranceOption);
    var maxPages = context.ParseResult.GetValueForOption(maxPagesOption);
    var masks = context.ParseResult.GetValueForOption(masksOption);
    var geometry = context.ParseResult.GetValueForOption(geometryOption);
    var geometryTolerance = context.ParseResult.GetValueForOption(geometryToleranceOption);
    var parallel = context.ParseResult.GetValueForOption(parallelOption);
    var diffPdfPath = context.ParseResult.GetValueForOption(diffPdfPathOption);

    WarnIfUnusedOptions(mode, threshold, tolerance, maxPages, masks, parallel);

    int exitCode;
    try
    {
        var normalizedPages = PageSelection.Normalize(pageRange);
        var parsedPages = PageSelection.Parse(pageRange);
        var runner = new DiffPdfRunner(diffPdfPath);
        var request = new DiffPdfRequest(
            BaselinePath: baseline.FullName,
            CandidatePath: candidate.FullName,
            OutputDirectory: outputDir.FullName,
            Pages: normalizedPages,
            PageNumbers: parsedPages,
            Dpi: dpi,
            GeometryCheck: geometry,
            GeometryTolerance: geometryTolerance);

        var result = await Task.Run(() => runner.Run(request));
        Console.Write(result.StdOut);
        Console.Error.Write(result.StdErr);

        exitCode = result.ExitCode switch
        {
            0 => 0,
            1 => 1,
            _ => 2
        };
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"Error: {ex.Message}");
        exitCode = 2;
    }

    context.ExitCode = exitCode;
});

return await root.InvokeAsync(args);

static void WarnIfUnusedOptions(
    string mode,
    double threshold,
    int tolerance,
    int? maxPages,
    FileInfo? masks,
    int parallel)
{
    static void Warn(string name) => Console.Error.WriteLine($"Warning: option {name} is ignored when using diff-pdf.");

    if (!string.Equals(mode, "overlay", StringComparison.OrdinalIgnoreCase)) Warn("--mode");
    if (Math.Abs(threshold) > double.Epsilon) Warn("--threshold");
    if (tolerance != 0) Warn("--antialias-tolerance");
    if (maxPages.HasValue) Warn("--max-pages");
    if (masks != null) Warn("--masks");
    if (parallel != Math.Max(1, Math.Min(Environment.ProcessorCount, 4))) Warn("--parallel");
}
