using System.Globalization;

namespace PdfVisualDiff.Core;

public static class PageSelection
{
    public static IReadOnlyList<int>? Parse(string? input)
    {
        if (string.IsNullOrWhiteSpace(input))
            return null;

        var pages = new SortedSet<int>();
        var segments = input.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        foreach (var segment in segments)
        {
            if (segment.Contains('-', StringComparison.Ordinal))
            {
                var bounds = segment.Split('-', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                if (bounds.Length != 2 ||
                    !int.TryParse(bounds[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out var start) ||
                    !int.TryParse(bounds[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var end))
                {
                    throw new FormatException($"Invalid page range '{segment}'.");
                }

                if (end < start)
                    (start, end) = (end, start);

                for (int page = start; page <= end; page++)
                    pages.Add(page);
            }
            else
            {
                if (!int.TryParse(segment, NumberStyles.Integer, CultureInfo.InvariantCulture, out var page))
                    throw new FormatException($"Invalid page identifier '{segment}'.");
                pages.Add(page);
            }
        }

        return pages.Count == 0 ? null : pages.ToArray();
    }

    public static string? Normalize(string? input)
    {
        var pages = Parse(input);
        if (pages == null || pages.Count == 0)
            return null;

        var ranges = new List<string>();
        int start = pages[0];
        int prev = start;

        for (int i = 1; i < pages.Count; i++)
        {
            var current = pages[i];
            if (current == prev + 1)
            {
                prev = current;
                continue;
            }

            ranges.Add(FormatRange(start, prev));
            start = prev = current;
        }

        ranges.Add(FormatRange(start, prev));
        return string.Join(',', ranges);
    }

    private static string FormatRange(int start, int end) =>
        start == end ? start.ToString(CultureInfo.InvariantCulture) : $"{start}-{end}";
}
