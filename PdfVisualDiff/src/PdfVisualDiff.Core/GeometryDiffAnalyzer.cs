using System;
using System.Collections.Generic;
using System.Linq;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace PdfVisualDiff.Core;

public static class GeometryDiffAnalyzer
{
    public static GeometryComparisonResult Compare(
        string baselinePath,
        string candidatePath,
        IReadOnlyList<int>? pageNumbers,
        double tolerancePt)
    {
        using var baselineDoc = PdfDocument.Open(baselinePath);
        using var candidateDoc = PdfDocument.Open(candidatePath);

        var totalPages = Math.Max(baselineDoc.NumberOfPages, candidateDoc.NumberOfPages);
        var numbers = pageNumbers is { Count: > 0 }
            ? pageNumbers.Where(p => p >= 1 && p <= totalPages).ToArray()
            : Enumerable.Range(1, totalPages).ToArray();

        var results = new List<GeometryPageResult>(numbers.Length);
        var passed = true;

        foreach (var pageNumber in numbers)
        {
            var expectedWords = pageNumber <= baselineDoc.NumberOfPages
                ? ExtractWords(baselineDoc.GetPage(pageNumber))
                : Array.Empty<WordInfo>();

            var actualWords = pageNumber <= candidateDoc.NumberOfPages
                ? ExtractWords(candidateDoc.GetPage(pageNumber))
                : Array.Empty<WordInfo>();

            var mismatches = CompareWords(expectedWords, actualWords, tolerancePt);
            if (expectedWords.Count != actualWords.Count || mismatches.Count > 0)
            {
                passed = false;
            }

            results.Add(new GeometryPageResult(pageNumber, expectedWords.Count, actualWords.Count, mismatches));
        }

        return new GeometryComparisonResult(passed, tolerancePt, results);
    }

    public static List<WordGeometryMismatch> CompareWords(
        IReadOnlyList<WordInfo> expected,
        IReadOnlyList<WordInfo> actual,
        double tolerancePt)
    {
        var mismatches = new List<WordGeometryMismatch>();
        var count = Math.Min(expected.Count, actual.Count);

        for (var i = 0; i < count; i++)
        {
            var exp = expected[i];
            var act = actual[i];

            if (!string.Equals(exp.Text, act.Text, StringComparison.Ordinal))
            {
                mismatches.Add(new WordGeometryMismatch(
                    i,
                    exp.Text,
                    act.Text,
                    exp.Bounds,
                    act.Bounds,
                    GeometryDelta.From(exp.Bounds, act.Bounds)));
                continue;
            }

            var delta = GeometryDelta.From(exp.Bounds, act.Bounds);
            if (delta.MaxDeviation > tolerancePt)
            {
                mismatches.Add(new WordGeometryMismatch(
                    i,
                    exp.Text,
                    act.Text,
                    exp.Bounds,
                    act.Bounds,
                    delta));
            }
        }

        if (expected.Count != actual.Count)
        {
            var mismatchIndex = count;
            mismatches.Add(new WordGeometryMismatch(
                mismatchIndex,
                mismatchIndex < expected.Count ? expected[mismatchIndex].Text : "<missing>",
                mismatchIndex < actual.Count ? actual[mismatchIndex].Text : "<missing>",
                mismatchIndex < expected.Count ? expected[Math.Min(mismatchIndex, expected.Count - 1)].Bounds : GeometryBox.Empty,
                mismatchIndex < actual.Count ? actual[Math.Min(mismatchIndex, actual.Count - 1)].Bounds : GeometryBox.Empty,
                GeometryDelta.Zero));
        }

        return mismatches;
    }

    private static IReadOnlyList<WordInfo> ExtractWords(Page page)
    {
        var list = new List<WordInfo>();
        foreach (var word in page.GetWords())
        {
            var text = word.Text?.Trim();
            if (string.IsNullOrEmpty(text))
                continue;

            var bounds = new GeometryBox(word.BoundingBox.Left, word.BoundingBox.Bottom, word.BoundingBox.Width, word.BoundingBox.Height);
            list.Add(new WordInfo(text, bounds));
        }

        return list;
    }
}

public sealed record GeometryComparisonResult(
    bool Passed,
    double TolerancePt,
    IReadOnlyList<GeometryPageResult> Pages);

public sealed record GeometryPageResult(
    int PageNumber,
    int ExpectedWordCount,
    int ActualWordCount,
    IReadOnlyList<WordGeometryMismatch> Mismatches);

public sealed record WordGeometryMismatch(
    int Index,
    string ExpectedText,
    string ActualText,
    GeometryBox Expected,
    GeometryBox Actual,
    GeometryDelta Delta);

public readonly record struct GeometryBox(double X, double Y, double Width, double Height)
{
    public double CenterX => X + (Width / 2d);
    public double CenterY => Y + (Height / 2d);

    public static GeometryBox Empty => new(0, 0, 0, 0);
}

public readonly record struct GeometryDelta(double DeltaX, double DeltaY, double DeltaWidth, double DeltaHeight)
{
    public double MaxDeviation => Math.Max(
        Math.Max(Math.Abs(DeltaX), Math.Abs(DeltaY)),
        Math.Max(Math.Abs(DeltaWidth), Math.Abs(DeltaHeight)));

    public static GeometryDelta From(GeometryBox expected, GeometryBox actual) =>
        new(actual.CenterX - expected.CenterX,
            actual.CenterY - expected.CenterY,
            actual.Width - expected.Width,
            actual.Height - expected.Height);

    public static GeometryDelta Zero => new(0, 0, 0, 0);
}

public readonly record struct WordInfo(string Text, GeometryBox Bounds);
