using System.Collections.Generic;
using System.Linq;
using UglyToad.PdfPig;

namespace PdfVisualDiff.Core;

public sealed class VisualDiffRunner
{
    private readonly PixelDiffToolPaths _toolPaths;
    private readonly PdfRasterizer _rasterizer;
    private readonly ImageMagickService _imageService;

    public VisualDiffRunner(PixelDiffToolPaths toolPaths)
    {
        _toolPaths = toolPaths;
        _rasterizer = new PdfRasterizer(toolPaths.PdftoppmPath);
        _imageService = new ImageMagickService(toolPaths.ComparePath, toolPaths.IdentifyPath, toolPaths.MagickPath);
    }

    public async Task<VisualDiffResult> RunAsync(VisualDiffRequest request, CancellationToken cancellationToken = default)
    {
        if (!File.Exists(request.BaselinePath))
            throw new FileNotFoundException("Baseline PDF not found", request.BaselinePath);
        if (!File.Exists(request.CandidatePath))
            throw new FileNotFoundException("Candidate PDF not found", request.CandidatePath);

        Directory.CreateDirectory(request.OutputDirectory);

        using var baselineDoc = PdfDocument.Open(request.BaselinePath);
        using var candidateDoc = PdfDocument.Open(request.CandidatePath);
        var baselinePages = baselineDoc.NumberOfPages;
        var candidatePages = candidateDoc.NumberOfPages;

        var desiredPages = request.PageNumbers is { Count: > 0 }
            ? request.PageNumbers
            : Enumerable.Range(1, Math.Min(baselinePages, candidatePages)).ToArray();

        var comparablePages = desiredPages.Where(page => page >= 1 && page <= baselinePages && page <= candidatePages)
            .Distinct()
            .OrderBy(p => p)
            .ToArray();

        if (comparablePages.Length == 0)
            throw new InvalidOperationException("No overlapping pages to compare between the two PDFs.");

        var missingFromBaseline = desiredPages.Where(page => page > baselinePages).Distinct().OrderBy(p => p).ToArray();
        var missingFromCandidate = desiredPages.Where(page => page > candidatePages).Distinct().OrderBy(p => p).ToArray();

        var rendererContexts = CreateRendererContexts(request.OutputDirectory);
        var pageResults = new List<PageComparisonResult>(comparablePages.Length);

        foreach (var page in comparablePages)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var popplerResult = await ProcessRendererPageAsync(
                    rendererContexts[RendererEngine.Poppler],
                    page,
                    request,
                    request.BaselinePasswords,
                    request.CandidatePasswords,
                    cancellationToken)
                .ConfigureAwait(false);

            var ghostscriptResult = await ProcessRendererPageAsync(
                    rendererContexts[RendererEngine.Ghostscript],
                    page,
                    request,
                    request.BaselinePasswords,
                    request.CandidatePasswords,
                    cancellationToken)
                .ConfigureAwait(false);

            var classification = Classify(popplerResult.Status, ghostscriptResult.Status);

            pageResults.Add(new PageComparisonResult(
                page,
                popplerResult,
                ghostscriptResult,
                classification));
        }

        var totalAe = pageResults.Sum(r => r.Poppler.AePixels);
        var totalPixelsAll = pageResults.Sum(r => r.Poppler.TotalPixels);
        var weightedPercent = totalPixelsAll == 0 ? 0d : totalAe / (double)totalPixelsAll * 100d;

        var (consistentPages, engineOnlyPages, cleanPages) = CountClassifications(pageResults);

        var overallStatus = pageResults
            .SelectMany(p => p.Renderers)
            .Select(r => r.Status)
            .DefaultIfEmpty(DiffStatus.Pass)
            .Aggregate(DiffStatus.Pass, DiffStatusHelper.Merge);

        GeometryAnalysisResult? geometryAnalysis = null;
        string? pdfBoxVersion = null;
        if (request.GeometryCheck)
        {
            var extractor = new PdfBoxGeometryExtractor(ResolveToolDirectory(request.PdfBoxDirectory, "pdfbox"));
            pdfBoxVersion = extractor.PdfBoxVersion;
            var analyzer = new PdfBoxGeometryAnalyzer(extractor);
            geometryAnalysis = await analyzer.AnalyzeAsync(
                    request.BaselinePath,
                    request.CandidatePath,
                    comparablePages,
                    request.GeometryTolerance,
                    cancellationToken)
                .ConfigureAwait(false);

            if (!geometryAnalysis.Passed)
            {
                overallStatus = DiffStatus.Fail;
            }
        }

        var tools = new ToolMetadata(
            ImageMagickVersion: await _imageService.GetVersionAsync(cancellationToken).ConfigureAwait(false),
            PopplerVersion: await _rasterizer.GetVersionAsync(cancellationToken).ConfigureAwait(false),
            GhostscriptVersion: await GetGhostscriptVersionAsync(request.GhostscriptPath, cancellationToken).ConfigureAwait(false),
            PdfBoxVersion: pdfBoxVersion);

        var summary = new VisualDiffSummary(
            request.BaselinePath,
            request.CandidatePath,
            baselinePages,
            candidatePages,
            comparablePages,
            missingFromBaseline,
            missingFromCandidate,
            totalAe,
            weightedPercent,
            consistentPages,
            engineOnlyPages,
            cleanPages,
            overallStatus,
            request.PageRangeExpression,
            request.Fuzz,
            request.UseCropBox,
            request.Dpi,
            request.GeometryCheck,
            request.Thresholds,
            request.SsimOptions,
            tools,
            DateTimeOffset.UtcNow,
            geometryAnalysis);

        var csvPath = Path.Combine(request.OutputDirectory, "report.csv");
        ReportBuilder.WriteCsv(csvPath, pageResults, request.OutputDirectory);

        var htmlPath = Path.Combine(request.OutputDirectory, "index.html");
        ReportBuilder.WriteHtml(htmlPath, summary, pageResults, geometryAnalysis, request.OutputDirectory);

        var jsonPath = Path.Combine(request.OutputDirectory, "report.json");
        ReportBuilder.WriteJson(jsonPath, summary, pageResults, geometryAnalysis);

        return new VisualDiffResult(summary, pageResults, csvPath, htmlPath, jsonPath, geometryAnalysis);
    }

    private async Task<RendererComparisonResult> ProcessRendererPageAsync(
        RendererContext context,
        int page,
        VisualDiffRequest request,
        PdfPasswordOptions baselinePasswords,
        PdfPasswordOptions candidatePasswords,
        CancellationToken cancellationToken)
    {
        var baselineImage = await RasterizeAsync(
                context,
                request.BaselinePath,
                page,
                true,
                request,
                baselinePasswords,
                cancellationToken)
            .ConfigureAwait(false);

        var candidateImage = await RasterizeAsync(
                context,
                request.CandidatePath,
                page,
                false,
                request,
                candidatePasswords,
                cancellationToken)
            .ConfigureAwait(false);

        var (width, height) = await _imageService.IdentifyAsync(baselineImage, cancellationToken)
            .ConfigureAwait(false);
        var (candidateWidth, candidateHeight) = await _imageService.IdentifyAsync(candidateImage, cancellationToken)
            .ConfigureAwait(false);

        if (width != candidateWidth || height != candidateHeight)
        {
            var targetWidth = Math.Max(width, candidateWidth);
            var targetHeight = Math.Max(height, candidateHeight);

            if (width != targetWidth || height != targetHeight)
            {
                await _imageService.PadToSizeAsync(baselineImage, targetWidth, targetHeight, cancellationToken).ConfigureAwait(false);
                width = targetWidth;
                height = targetHeight;
            }

            if (candidateWidth != targetWidth || candidateHeight != targetHeight)
            {
                await _imageService.PadToSizeAsync(candidateImage, targetWidth, targetHeight, cancellationToken).ConfigureAwait(false);
                candidateWidth = targetWidth;
                candidateHeight = targetHeight;
            }
        }

        var diffImage = Path.Combine(context.DiffDir, $"page-{page:D4}.png");
        var metrics = await _imageService.CompareAsync(
                baselineImage,
                candidateImage,
                diffImage,
                request.Fuzz,
                request.SsimOptions,
                cancellationToken)
            .ConfigureAwait(false);

        var overlayImage = Path.Combine(context.OverlayDir, $"page-{page:D4}.png");
        await _imageService.CreateOverlayAsync(baselineImage, candidateImage, overlayImage, cancellationToken)
            .ConfigureAwait(false);

        var totalPixels = (long)width * height;
        var aePercent = totalPixels == 0
            ? 0d
            : metrics.AePixels / (double)totalPixels * 100d;
        var status = DiffStatusHelper.FromMetrics(metrics.AePixels, aePercent, metrics.Psnr, metrics.Ssim, metrics.Dssim, request.Thresholds);

        var baselineThumb = Path.Combine(context.ThumbsDir, $"page-{page:D4}-A.png");
        var candidateThumb = Path.Combine(context.ThumbsDir, $"page-{page:D4}-B.png");
        var diffThumb = Path.Combine(context.ThumbsDir, $"page-{page:D4}-diff.png");
        var overlayThumb = Path.Combine(context.ThumbsDir, $"page-{page:D4}-overlay.png");

        await _imageService.CreateThumbnailAsync(baselineImage, baselineThumb, request.ThumbnailMaxSize, cancellationToken).ConfigureAwait(false);
        await _imageService.CreateThumbnailAsync(candidateImage, candidateThumb, request.ThumbnailMaxSize, cancellationToken).ConfigureAwait(false);
        await _imageService.CreateThumbnailAsync(diffImage, diffThumb, request.ThumbnailMaxSize, cancellationToken).ConfigureAwait(false);
        await _imageService.CreateThumbnailAsync(overlayImage, overlayThumb, request.ThumbnailMaxSize, cancellationToken).ConfigureAwait(false);

        return new RendererComparisonResult(
            context.Engine,
            width,
            height,
            request.Dpi,
            metrics.AePixels,
            aePercent,
            metrics.Mae,
            metrics.Rmse,
            metrics.Psnr,
            metrics.Ssim,
            metrics.Dssim,
            status,
            baselineImage,
            candidateImage,
            diffImage,
            baselineThumb,
            candidateThumb,
            diffThumb,
            overlayImage,
            overlayThumb);
    }

    private static PageDiffClassification Classify(DiffStatus popplerStatus, DiffStatus ghostscriptStatus)
    {
        var popplerDiffers = popplerStatus != DiffStatus.Pass;
        var ghostscriptDiffers = ghostscriptStatus != DiffStatus.Pass;

        if (popplerDiffers && ghostscriptDiffers)
            return PageDiffClassification.Consistent;
        if (popplerDiffers || ghostscriptDiffers)
            return PageDiffClassification.EngineOnly;
        return PageDiffClassification.Clean;
    }

    private static (int Consistent, int EngineOnly, int Clean) CountClassifications(IEnumerable<PageComparisonResult> pages)
    {
        int consistent = 0, engineOnly = 0, clean = 0;
        foreach (var page in pages)
        {
            switch (page.Classification)
            {
                case PageDiffClassification.Consistent:
                    consistent++;
                    break;
                case PageDiffClassification.EngineOnly:
                    engineOnly++;
                    break;
                default:
                    clean++;
                    break;
            }
        }

        return (consistent, engineOnly, clean);
    }

    private async Task<string> RasterizeAsync(
        RendererContext context,
        string pdfPath,
        int page,
        bool isBaseline,
        VisualDiffRequest request,
        PdfPasswordOptions passwords,
        CancellationToken cancellationToken)
    {
        var targetDir = isBaseline ? context.BaselineDir : context.CandidateDir;
        var basePath = Path.Combine(targetDir, $"page-{page:D4}");

        return context.Engine switch
        {
            RendererEngine.Poppler => await _rasterizer.RasterizeAsync(
                new RasterizeRequest(
                    pdfPath,
                    page,
                    basePath,
                    request.Dpi,
                    request.UseCropBox,
                    passwords.UserPassword,
                    passwords.OwnerPassword),
                cancellationToken).ConfigureAwait(false),
            RendererEngine.Ghostscript => await _imageService.RenderPdfPageAsync(
                pdfPath,
                page,
                basePath + ".png",
                request.Dpi,
                request.UseCropBox,
                SelectPassword(passwords),
                cancellationToken).ConfigureAwait(false),
            _ => throw new NotSupportedException($"Renderer {context.Engine} is not supported.")
        };
    }

    private static string? SelectPassword(PdfPasswordOptions options) =>
        options.UserPassword ?? options.OwnerPassword;

    private async Task<string?> GetGhostscriptVersionAsync(string? ghostscriptPath, CancellationToken cancellationToken)
    {
        var executable = string.IsNullOrWhiteSpace(ghostscriptPath) ? "gs" : ghostscriptPath!;

        try
        {
            var result = await ProcessRunner.RunAsync(executable, new[] { "--version" }, cancellationToken: cancellationToken)
                .ConfigureAwait(false);
            return $"{executable} {result.StdOut.Trim()}";
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Ghostscript executable not found. Install Ghostscript or provide --ghostscript-path.", ex);
        }
    }

    private Dictionary<RendererEngine, RendererContext> CreateRendererContexts(string outputRoot)
    {
        var contexts = new[]
        {
            CreateRendererContext(RendererEngine.Poppler, outputRoot, "poppler"),
            CreateRendererContext(RendererEngine.Ghostscript, outputRoot, "ghostscript")
        };

        return contexts.ToDictionary(c => c.Engine);
    }

    private static RendererContext CreateRendererContext(RendererEngine engine, string outputRoot, string folderName)
    {
        var root = Path.Combine(outputRoot, folderName);
        var context = new RendererContext(
            engine,
            folderName,
            Path.Combine(root, "A"),
            Path.Combine(root, "B"),
            Path.Combine(root, "diff"),
            Path.Combine(root, "overlay"),
            Path.Combine(root, "thumbs"));

        Directory.CreateDirectory(context.BaselineDir);
        Directory.CreateDirectory(context.CandidateDir);
        Directory.CreateDirectory(context.DiffDir);
        Directory.CreateDirectory(context.OverlayDir);
        Directory.CreateDirectory(context.ThumbsDir);

        return context;
    }

    private static string ResolveToolDirectory(string? overridePath, string folderName)
    {
        if (!string.IsNullOrWhiteSpace(overridePath))
            return overridePath!;

        var probe = Path.Combine(AppContext.BaseDirectory, "tools", folderName);
        if (Directory.Exists(probe))
            return probe;

        var current = AppContext.BaseDirectory;
        for (int i = 0; i < 8; i++)
        {
            current = Path.GetFullPath(Path.Combine(current, ".."));
            var candidate = Path.Combine(current, "tools", folderName);
            if (Directory.Exists(candidate))
                return candidate;
        }

        return probe;
    }

    private sealed record RendererContext(
        RendererEngine Engine,
        string FolderName,
        string BaselineDir,
        string CandidateDir,
        string DiffDir,
        string OverlayDir,
        string ThumbsDir);
}
