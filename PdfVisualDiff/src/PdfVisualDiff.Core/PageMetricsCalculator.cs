using UglyToad.PdfPig;

namespace PdfVisualDiff.Core;

internal static class PageMetricsCalculator
{
    public static IReadOnlyDictionary<int, PagePixelMetrics> Calculate(
        string pdfPath,
        IReadOnlyList<int>? pageNumbers,
        int dpi)
    {
        if (dpi <= 0)
            throw new ArgumentOutOfRangeException(nameof(dpi), "DPI must be greater than zero.");

        using var document = PdfDocument.Open(pdfPath);
        var numbers = pageNumbers is { Count: > 0 }
            ? pageNumbers
            : Enumerable.Range(1, document.NumberOfPages).ToArray();

        var map = new Dictionary<int, PagePixelMetrics>(numbers.Count);
        foreach (var pageNumber in numbers)
        {
            if (pageNumber < 1 || pageNumber > document.NumberOfPages)
                throw new ArgumentOutOfRangeException(nameof(pageNumbers), $"Page {pageNumber} is out of range.");

            var page = document.GetPage(pageNumber);
            var widthPx = ConvertPointsToPixels(page.Width, dpi);
            var heightPx = ConvertPointsToPixels(page.Height, dpi);
            map[pageNumber] = new PagePixelMetrics(widthPx, heightPx);
        }

        return map;
    }

    private static int ConvertPointsToPixels(double points, int dpi)
    {
        var pixels = points / 72d * dpi;
        return Math.Max(1, (int)Math.Round(pixels, MidpointRounding.AwayFromZero));
    }
}

internal readonly record struct PagePixelMetrics(int WidthPx, int HeightPx)
{
    public long TotalPixels => (long)WidthPx * HeightPx;
}
