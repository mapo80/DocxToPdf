using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace PdfVisualDiff.Core;

public sealed class DiffPdfRunner
{
    private readonly string _toolPath;

    public DiffPdfRunner(string? toolPath = null)
    {
        _toolPath = string.IsNullOrWhiteSpace(toolPath) ? "diff-pdf" : toolPath;
    }

    public DiffPdfResult Run(DiffPdfRequest request, CancellationToken cancellationToken = default)
    {
        if (!File.Exists(request.BaselinePath))
            throw new FileNotFoundException("Baseline PDF not found", request.BaselinePath);
        if (!File.Exists(request.CandidatePath))
            throw new FileNotFoundException("Candidate PDF not found", request.CandidatePath);

        Directory.CreateDirectory(request.OutputDirectory);
        var diffPath = Path.Combine(request.OutputDirectory, request.DiffFileName ?? "diff.pdf");

        var psi = new ProcessStartInfo(_toolPath)
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };

        psi.ArgumentList.Add("-v");
        psi.ArgumentList.Add("--output-diff");
        psi.ArgumentList.Add(diffPath);
        if (request.Dpi > 0)
        {
            psi.ArgumentList.Add("--dpi");
            psi.ArgumentList.Add(request.Dpi.ToString(CultureInfo.InvariantCulture));
        }
        if (!string.IsNullOrWhiteSpace(request.Pages))
        {
            psi.ArgumentList.Add("--pages");
            psi.ArgumentList.Add(request.Pages!);
        }

        psi.ArgumentList.Add(request.BaselinePath);
        psi.ArgumentList.Add(request.CandidatePath);

        string stdout;
        string stderr;
        int exitCode;

        try
        {
            using var process = Process.Start(psi)
                ?? throw new InvalidOperationException("Unable to start diff-pdf process.");

            stdout = process.StandardOutput.ReadToEnd();
            stderr = process.StandardError.ReadToEnd();
            process.WaitForExit();
            exitCode = process.ExitCode;
        }
        catch (Win32Exception ex) when (ex.NativeErrorCode == 2)
        {
            throw new InvalidOperationException("diff-pdf executable not found. Install it and ensure it is on PATH.", ex);
        }

        var passed = exitCode == 0;
        var diffExists = File.Exists(diffPath);

        var metrics = PageMetricsCalculator.Calculate(request.BaselinePath, request.PageNumbers, request.Dpi);
        var pageStats = BuildPageStats(metrics, stdout);
        var totalDiffPixels = pageStats.Sum(p => p.DiffPixels);
        var totalPixels = pageStats.Sum(p => p.TotalPixels);
        var diffRatio = totalPixels == 0 ? 0 : (double)totalDiffPixels / totalPixels;

        GeometryComparisonResult? geometryResult = null;
        if (request.GeometryCheck)
        {
            geometryResult = GeometryDiffAnalyzer.Compare(
                request.BaselinePath,
                request.CandidatePath,
                request.PageNumbers,
                request.GeometryTolerance);

            if (!geometryResult.Passed)
            {
                passed = false;
                if (exitCode == 0)
                {
                    exitCode = 1;
                }
            }
        }

        var report = new
        {
            passed,
            exitCode,
            diffPdf = diffExists ? diffPath : null,
            stdout,
            stderr,
            baseline = request.BaselinePath,
            candidate = request.CandidatePath,
            pages = request.Pages,
            dpi = request.Dpi,
            difference = new
            {
                totalPixels,
                differingPixels = totalDiffPixels,
                ratio = diffRatio,
                percent = diffRatio * 100
            },
            pageStats,
            geometry = geometryResult
        };

        var reportPath = Path.Combine(request.OutputDirectory, "report.json");
        File.WriteAllText(reportPath, JsonSerializer.Serialize(report, new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        }));

        return new DiffPdfResult(exitCode, passed, diffExists ? diffPath : null, stdout, stderr, pageStats, diffRatio, geometryResult);
    }

    private static List<PageDiffStat> BuildPageStats(IReadOnlyDictionary<int, PagePixelMetrics> metrics, string stdout)
    {
        var diffPixelsByPage = ParseDiffOutput(stdout);
        var stats = new List<PageDiffStat>(metrics.Count);

        foreach (var (pageNumber, pageMetrics) in metrics.OrderBy(item => item.Key))
        {
            var zeroBasedIndex = pageNumber - 1;
            diffPixelsByPage.TryGetValue(zeroBasedIndex, out var diffPixels);

            stats.Add(new PageDiffStat(
                pageNumber,
                pageMetrics.WidthPx,
                pageMetrics.HeightPx,
                pageMetrics.TotalPixels,
                diffPixels));
        }

        return stats;
    }

    private static IReadOnlyDictionary<int, long> ParseDiffOutput(string stdout)
    {
        var matches = DiffLineRegex.Matches(stdout);
        var map = new Dictionary<int, long>(matches.Count);
        foreach (Match match in matches)
        {
            if (!match.Success)
                continue;

            var pageIndex = int.Parse(match.Groups["page"].Value, CultureInfo.InvariantCulture);
            var diffPixels = long.Parse(match.Groups["pixels"].Value, CultureInfo.InvariantCulture);
            map[pageIndex] = diffPixels;
        }

        return map;
    }

    private static readonly Regex DiffLineRegex = new(
        @"page\s+(?<page>\d+)\s+has\s+(?<pixels>\d+)\s+pixels\s+that\s+differ",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
}

public sealed record DiffPdfRequest(
    string BaselinePath,
    string CandidatePath,
    string OutputDirectory,
    string? Pages,
    IReadOnlyList<int>? PageNumbers,
    int Dpi,
    bool GeometryCheck,
    double GeometryTolerance,
    string? DiffFileName = null);

public sealed record DiffPdfResult(
    int ExitCode,
    bool Passed,
    string? DiffPdfPath,
    string StdOut,
    string StdErr,
    IReadOnlyList<PageDiffStat> PageStats,
    double DifferenceRatio,
    GeometryComparisonResult? Geometry);

public sealed record PageDiffStat(
    int PageNumber,
    int WidthPx,
    int HeightPx,
    long TotalPixels,
    long DiffPixels)
{
    public double Ratio => TotalPixels == 0 ? 0 : (double)DiffPixels / TotalPixels;
    public double Percent => Ratio * 100;
}
