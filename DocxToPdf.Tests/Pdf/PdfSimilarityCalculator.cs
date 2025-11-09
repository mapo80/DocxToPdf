using SkiaSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DocxToPdf.Tests.Pdf;

internal sealed class PdfSimilarityCalculator
{
    public PdfSimilarityResult Calculate(string baselinePdfPath, string candidatePdfPath, int dpi = 200)
    {
        using var baseline = PdfRasterizer.Rasterize(baselinePdfPath, dpi);
        using var candidate = PdfRasterizer.Rasterize(candidatePdfPath, dpi);

        var pageCount = Math.Max(baseline.ImagePaths.Count, candidate.ImagePaths.Count);
        if (pageCount == 0)
            return new PdfSimilarityResult(1, Array.Empty<PageSimilarity>());

        var pageResults = new List<PageSimilarity>(pageCount);

        for (int i = 0; i < pageCount; i++)
        {
            double pageSimilarity;

            if (i >= baseline.ImagePaths.Count || i >= candidate.ImagePaths.Count)
            {
                pageSimilarity = 0;
            }
            else
            {
                var comparison = CompareImages(baseline.ImagePaths[i], candidate.ImagePaths[i]);
                pageSimilarity = comparison.TotalPixels == 0
                    ? 1
                    : (double)comparison.MatchingPixels / comparison.TotalPixels;
            }

            pageResults.Add(new PageSimilarity(i + 1, pageSimilarity));
        }

        var overall = pageResults.Average(p => p.Similarity);
        return new PdfSimilarityResult(overall, pageResults);
    }

    private static ImageComparison CompareImages(string baselineImagePath, string candidateImagePath)
    {
        using var baselineBitmap = SKBitmap.Decode(baselineImagePath)
            ?? throw new InvalidOperationException($"Unable to decode PNG: {baselineImagePath}");
        using var candidateBitmap = SKBitmap.Decode(candidateImagePath)
            ?? throw new InvalidOperationException($"Unable to decode PNG: {candidateImagePath}");

        var overlapWidth = Math.Min(baselineBitmap.Width, candidateBitmap.Width);
        var overlapHeight = Math.Min(baselineBitmap.Height, candidateBitmap.Height);

        if (overlapWidth <= 0 || overlapHeight <= 0)
        {
            var total = (long)Math.Max(baselineBitmap.Width, candidateBitmap.Width) *
                        Math.Max(baselineBitmap.Height, candidateBitmap.Height);
            return new ImageComparison(0, total);
        }

        var baselinePixels = baselineBitmap.Pixels;
        var candidatePixels = candidateBitmap.Pixels;

        if (baselinePixels is null || candidatePixels is null)
            throw new InvalidOperationException("Unable to access bitmap pixel data.");

        long matching = 0;
        for (int y = 0; y < overlapHeight; y++)
        {
            var baseRowStart = y * baselineBitmap.Width;
            var candRowStart = y * candidateBitmap.Width;

            for (int x = 0; x < overlapWidth; x++)
            {
                if (baselinePixels[baseRowStart + x] == candidatePixels[candRowStart + x])
                    matching++;
            }
        }

        var totalPixels = (long)Math.Max(baselineBitmap.Width, candidateBitmap.Width) *
                          Math.Max(baselineBitmap.Height, candidateBitmap.Height);
        return new ImageComparison(matching, totalPixels);
    }

    private readonly record struct ImageComparison(long MatchingPixels, long TotalPixels);
}

internal sealed record PdfSimilarityResult(double OverallSimilarity, IReadOnlyList<PageSimilarity> Pages);

internal sealed record PageSimilarity(int PageNumber, double Similarity);
