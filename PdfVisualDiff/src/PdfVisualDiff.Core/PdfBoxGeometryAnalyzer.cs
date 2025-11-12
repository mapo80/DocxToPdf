namespace PdfVisualDiff.Core;

internal sealed class PdfBoxGeometryAnalyzer
{
    private readonly PdfBoxGeometryExtractor _extractor;

    public PdfBoxGeometryAnalyzer(PdfBoxGeometryExtractor extractor)
    {
        _extractor = extractor;
    }

    public async Task<GeometryAnalysisResult> AnalyzeAsync(
        string baselinePath,
        string candidatePath,
        IReadOnlyList<int> pages,
        double tolerancePt,
        CancellationToken cancellationToken)
    {
        var baseline = await _extractor.ExtractAsync(baselinePath, pages, cancellationToken).ConfigureAwait(false);
        var candidate = await _extractor.ExtractAsync(candidatePath, pages, cancellationToken).ConfigureAwait(false);

        var baselineMap = baseline.Pages.ToDictionary(p => p.Page);
        var candidateMap = candidate.Pages.ToDictionary(p => p.Page);

        var reports = new List<GeometryPageReport>(pages.Count);
        var moveThreshold = Math.Max(tolerancePt * 10, 6.0);
        var matchThreshold = Math.Max(moveThreshold * 1.5, 10.0);
        var graphicNoiseThreshold = Math.Max(tolerancePt * 10, 5.0);
        var passed = true;

        foreach (var page in pages)
        {
            baselineMap.TryGetValue(page, out var baselinePage);
            candidateMap.TryGetValue(page, out var candidatePage);

            var report = AnalyzePage(
                page,
                baselinePage,
                candidatePage,
                moveThreshold,
                matchThreshold,
                graphicNoiseThreshold);

            reports.Add(report);
            if (!report.IsClean)
            {
                passed = false;
            }
        }

        return new GeometryAnalysisResult(
            tolerancePt,
            passed,
            reports);
    }

    private static GeometryPageReport AnalyzePage(
        int pageNumber,
        PdfBoxPageGeometry? baseline,
        PdfBoxPageGeometry? candidate,
        double moveThreshold,
        double matchThreshold,
        double graphicNoiseThreshold)
    {
        var baselineWords = baseline?.Words ?? Array.Empty<PdfBoxWordGeometry>();
        var candidateWords = candidate?.Words ?? Array.Empty<PdfBoxWordGeometry>();
        var baselineGraphics = baseline?.Graphics ?? Array.Empty<PdfBoxGraphicGeometry>();
        var candidateGraphics = candidate?.Graphics ?? Array.Empty<PdfBoxGraphicGeometry>();

        var missingWords = new List<WordDifference>();
        var movedWords = new List<WordDifference>();
        var extraWords = new List<WordDifference>();

        MatchWords(baselineWords, candidateWords, moveThreshold, matchThreshold, missingWords, extraWords, movedWords);

        var addedGraphics = new List<GraphicDifference>();
        var removedGraphics = new List<GraphicDifference>();
        var changedGraphics = new List<GraphicDifference>();
        MatchGraphics(
            baselineGraphics,
            candidateGraphics,
            moveThreshold,
            matchThreshold,
            addedGraphics,
            removedGraphics,
            changedGraphics);
        FilterTrivialGraphics(addedGraphics, removedGraphics, changedGraphics, graphicNoiseThreshold);

        var (wordStatus, wordStatusNote) = EvaluateWordStatus(
            baselineWords.Count,
            candidateWords.Count,
            missingWords.Count,
            extraWords.Count);
        if (wordStatus != GeometryWordStatus.Ok)
        {
            missingWords.Clear();
            extraWords.Clear();
            movedWords.Clear();
        }

        return new GeometryPageReport(
            pageNumber,
            missingWords,
            extraWords,
            movedWords,
            addedGraphics,
            removedGraphics,
            changedGraphics,
            wordStatus,
            wordStatusNote);
    }

    private static void MatchWords(
        IReadOnlyList<PdfBoxWordGeometry> baseline,
        IReadOnlyList<PdfBoxWordGeometry> candidate,
        double moveThreshold,
        double matchThreshold,
        List<WordDifference> missing,
        List<WordDifference> extra,
        List<WordDifference> moved)
    {
        var candidateBuckets = BucketWords(candidate);

        foreach (var word in baseline)
        {
            var key = Normalize(word.Text);
            if (string.IsNullOrEmpty(key))
                continue;

            if (!candidateBuckets.TryGetValue(key, out var list) || list.Count == 0)
            {
                missing.Add(WordOnlyDifference(word));
                continue;
            }

            var match = TakeNearest(word, list);
            if (match == null || match.Distance > matchThreshold)
            {
                missing.Add(WordOnlyDifference(word));
                continue;
            }

            if (match.Distance > moveThreshold)
            {
                moved.Add(CreateWordDifference(word, match.Value));
            }
        }

        foreach (var bucket in candidateBuckets.Values)
        {
            foreach (var word in bucket)
            {
                extra.Add(WordOnlyDifference(word, isBaseline: false));
            }
        }
    }

    private static void MatchGraphics(
        IReadOnlyList<PdfBoxGraphicGeometry> baseline,
        IReadOnlyList<PdfBoxGraphicGeometry> candidate,
        double moveThreshold,
        double matchThreshold,
        List<GraphicDifference> added,
        List<GraphicDifference> removed,
        List<GraphicDifference> changed)
    {
        var candidateBuckets = BucketGraphics(candidate);

        foreach (var graphic in baseline)
        {
            var key = graphic.Type ?? "unknown";
            if (!candidateBuckets.TryGetValue(key, out var list) || list.Count == 0)
            {
                removed.Add(GraphicOnlyDifference(graphic));
                continue;
            }

            var match = TakeNearest(graphic, list);
            if (match == null || match.Distance > matchThreshold)
            {
                removed.Add(GraphicOnlyDifference(graphic));
                continue;
            }

            if (match.Distance > moveThreshold ||
                Math.Abs(graphic.Width - match.Value.Width) > moveThreshold ||
                Math.Abs(graphic.Height - match.Value.Height) > moveThreshold)
            {
                changed.Add(GraphicDifference(graphic, match.Value));
            }
        }

        foreach (var bucket in candidateBuckets.Values)
        {
            foreach (var graphic in bucket)
            {
                added.Add(GraphicOnlyDifference(graphic, isBaseline: false));
            }
        }
    }

    private static void FilterTrivialGraphics(
        List<GraphicDifference> added,
        List<GraphicDifference> removed,
        List<GraphicDifference> changed,
        double noiseThreshold)
    {
        static bool WithinThreshold(BoundingBox? box, double threshold) =>
            box != null && box.Width <= threshold && box.Height <= threshold;

        added.RemoveAll(g => WithinThreshold(g.Candidate, noiseThreshold));
        removed.RemoveAll(g => WithinThreshold(g.Baseline, noiseThreshold));
        changed.RemoveAll(g =>
            WithinThreshold(g.Baseline, noiseThreshold) &&
            WithinThreshold(g.Candidate, noiseThreshold));
    }

    private static Dictionary<string, List<PdfBoxWordGeometry>> BucketWords(IReadOnlyList<PdfBoxWordGeometry> words)
    {
        var buckets = new Dictionary<string, List<PdfBoxWordGeometry>>(StringComparer.Ordinal);
        foreach (var word in words)
        {
            var key = Normalize(word.Text);
            if (string.IsNullOrEmpty(key))
                continue;

            if (!buckets.TryGetValue(key, out var list))
            {
                list = new List<PdfBoxWordGeometry>();
                buckets[key] = list;
            }

            list.Add(word);
        }

        return buckets;
    }

    private static Dictionary<string, List<PdfBoxGraphicGeometry>> BucketGraphics(IReadOnlyList<PdfBoxGraphicGeometry> graphics)
    {
        var buckets = new Dictionary<string, List<PdfBoxGraphicGeometry>>(StringComparer.OrdinalIgnoreCase);
        foreach (var graphic in graphics)
        {
            var key = string.IsNullOrWhiteSpace(graphic.Type) ? "unknown" : graphic.Type;
            if (!buckets.TryGetValue(key, out var list))
            {
                list = new List<PdfBoxGraphicGeometry>();
                buckets[key] = list;
            }

            list.Add(graphic);
        }

        return buckets;
    }

    private static MatchResult<PdfBoxWordGeometry>? TakeNearest(PdfBoxWordGeometry reference, List<PdfBoxWordGeometry> candidates)
    {
        if (candidates.Count == 0)
            return null;

        var referenceBox = ToBoundingBox(reference);
        var refCenterX = referenceBox.CenterX;
        var refCenterY = referenceBox.CenterY;

        double bestDistance = double.MaxValue;
        int bestIndex = -1;
        double deltaX = 0;
        double deltaY = 0;

        for (int i = 0; i < candidates.Count; i++)
        {
            var candidate = candidates[i];
            var candidateBox = ToBoundingBox(candidate);
            var dx = candidateBox.CenterX - refCenterX;
            var dy = candidateBox.CenterY - refCenterY;
            var distance = Math.Sqrt((dx * dx) + (dy * dy));
            if (distance < bestDistance)
            {
                bestDistance = distance;
                bestIndex = i;
                deltaX = dx;
                deltaY = dy;
            }
        }

        if (bestIndex < 0)
            return null;

        var match = candidates[bestIndex];
        candidates.RemoveAt(bestIndex);
        return new MatchResult<PdfBoxWordGeometry>(match, deltaX, deltaY, bestDistance);
    }

    private static MatchResult<PdfBoxGraphicGeometry>? TakeNearest(PdfBoxGraphicGeometry reference, List<PdfBoxGraphicGeometry> candidates)
    {
        if (candidates.Count == 0)
            return null;

        var referenceBox = ToBoundingBox(reference);
        var refCenterX = referenceBox.CenterX;
        var refCenterY = referenceBox.CenterY;

        double bestDistance = double.MaxValue;
        int bestIndex = -1;
        double deltaX = 0;
        double deltaY = 0;

        for (int i = 0; i < candidates.Count; i++)
        {
            var candidate = candidates[i];
            var candidateBox = ToBoundingBox(candidate);
            var dx = candidateBox.CenterX - refCenterX;
            var dy = candidateBox.CenterY - refCenterY;
            var distance = Math.Sqrt((dx * dx) + (dy * dy));
            if (distance < bestDistance)
            {
                bestDistance = distance;
                bestIndex = i;
                deltaX = dx;
                deltaY = dy;
            }
        }

        if (bestIndex < 0)
            return null;

        var match = candidates[bestIndex];
        candidates.RemoveAt(bestIndex);
        return new MatchResult<PdfBoxGraphicGeometry>(match, deltaX, deltaY, bestDistance);
    }

    private static WordDifference WordOnlyDifference(PdfBoxWordGeometry word, bool isBaseline = true) =>
        isBaseline
            ? new WordDifference(word.Text, ToBoundingBox(word), null, GeometryDelta.Zero, word.FontSize)
            : new WordDifference(word.Text, null, ToBoundingBox(word), GeometryDelta.Zero, word.FontSize);

    private static WordDifference CreateWordDifference(PdfBoxWordGeometry baseline, PdfBoxWordGeometry candidate) =>
        new(
            baseline.Text,
            ToBoundingBox(baseline),
            ToBoundingBox(candidate),
            GeometryDelta.From(ToBox(baseline), ToBox(candidate)),
            candidate.FontSize);

    private static GraphicDifference GraphicOnlyDifference(PdfBoxGraphicGeometry graphic, bool isBaseline = true) =>
        isBaseline
            ? new GraphicDifference(graphic.Type, ToBoundingBox(graphic), null, GeometryDelta.Zero, graphic.StrokeWidth, graphic.StrokeColor, graphic.FillColor)
            : new GraphicDifference(graphic.Type, null, ToBoundingBox(graphic), GeometryDelta.Zero, graphic.StrokeWidth, graphic.StrokeColor, graphic.FillColor);

    private static GraphicDifference GraphicDifference(PdfBoxGraphicGeometry baseline, PdfBoxGraphicGeometry candidate) =>
        new GraphicDifference(
            baseline.Type,
            ToBoundingBox(baseline),
            ToBoundingBox(candidate),
            GeometryDelta.From(ToBox(baseline), ToBox(candidate)),
            candidate.StrokeWidth,
            candidate.StrokeColor,
            candidate.FillColor);

    private static BoundingBox ToBoundingBox(PdfBoxWordGeometry word) =>
        new(word.X, word.Y, word.Width, word.Height);

    private static BoundingBox ToBoundingBox(PdfBoxGraphicGeometry graphic) =>
        new(graphic.X, graphic.Y, graphic.Width, graphic.Height);

    private static GeometryBox ToBox(PdfBoxWordGeometry word) =>
        new(word.X, word.Y, word.Width, word.Height);

    private static GeometryBox ToBox(PdfBoxGraphicGeometry graphic) =>
        new(graphic.X, graphic.Y, graphic.Width, graphic.Height);

    private static string Normalize(string text) =>
        string.IsNullOrWhiteSpace(text)
            ? string.Empty
            : text.Trim().ToUpperInvariant();

    private static (GeometryWordStatus Status, string? Note) EvaluateWordStatus(
        int baselineCount,
        int candidateCount,
        int missingCount,
        int extraCount)
    {
        if (baselineCount == 0 && candidateCount == 0)
            return (GeometryWordStatus.NoText, "Nessun testo estraibile su questa pagina (probabilmente solo grafica).");

        if (baselineCount <= 5 && candidateCount > 20)
            return (GeometryWordStatus.BaselineTextUnavailable, "Il PDF di riferimento non espone testo estraibile, quindi il confronto parole non è disponibile.");

        if (candidateCount <= 5 && baselineCount > 20)
            return (GeometryWordStatus.CandidateTextUnavailable, "Il PDF candidato non espone testo estraibile (spesso accade quando il testo viene convertito in tracciati). Il riepilogo \"Missing words\" è disattivato per evitare falsi positivi.");

        double baselineCoverage = baselineCount == 0 ? 1 : 1 - (double)missingCount / baselineCount;
        double candidateCoverage = candidateCount == 0 ? 1 : 1 - (double)extraCount / Math.Max(1, candidateCount);

        if (baselineCoverage < 0.7)
            return (GeometryWordStatus.ComparisonUnreliable, "Oltre la metà del testo del PDF di riferimento non è stato trovato nel candidato. Il confronto parole è disattivato per evitare falsi positivi (probabile testo rasterizzato/vettoriale).");

        if (candidateCoverage < 0.7)
            return (GeometryWordStatus.ComparisonUnreliable, "Oltre la metà del testo del candidato non è mappabile sul riferimento. Il confronto parole è disattivato per evitare falsi positivi.");

        return (GeometryWordStatus.Ok, null);
    }

    private sealed record MatchResult<T>(T Value, double DeltaX, double DeltaY, double Distance);
}
