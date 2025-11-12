namespace PdfVisualDiff.Core;

public sealed record VisualDiffRequest(
    string BaselinePath,
    string CandidatePath,
    string OutputDirectory,
    IReadOnlyList<int>? PageNumbers,
    string? PageRangeExpression,
    int Dpi,
    bool UseCropBox,
    string? Fuzz,
    PdfPasswordOptions BaselinePasswords,
    PdfPasswordOptions CandidatePasswords,
    int ThumbnailMaxSize,
    bool GeometryCheck,
    double GeometryTolerance,
    QualityThresholds Thresholds,
    SsimOptions SsimOptions,
    string? GhostscriptPath,
    string? PdfBoxDirectory);

public sealed record PdfPasswordOptions(string? UserPassword, string? OwnerPassword);

public sealed record VisualDiffResult(
    VisualDiffSummary Summary,
    IReadOnlyList<PageComparisonResult> Pages,
    string CsvPath,
    string HtmlPath,
    string JsonPath,
    GeometryAnalysisResult? GeometryAnalysis);

public sealed record VisualDiffSummary(
    string BaselinePath,
    string CandidatePath,
    int BaselinePageCount,
    int CandidatePageCount,
    IReadOnlyList<int> ComparedPages,
    IReadOnlyList<int> MissingFromBaseline,
    IReadOnlyList<int> MissingFromCandidate,
    long TotalAePixels,
    double WeightedAePercent,
    int ConsistentPages,
    int EngineOnlyPages,
    int CleanPages,
    DiffStatus OverallStatus,
    string? PageRangeExpression,
    string? Fuzz,
    bool UseCropBox,
    int Dpi,
    bool GeometryRequested,
    QualityThresholds Thresholds,
    SsimOptions SsimOptions,
    ToolMetadata Tools,
    DateTimeOffset GeneratedAtUtc,
    GeometryAnalysisResult? GeometryAnalysis);

public sealed record PageComparisonResult(
    int PageNumber,
    RendererComparisonResult Poppler,
    RendererComparisonResult Ghostscript,
    PageDiffClassification Classification)
{
    public IEnumerable<RendererComparisonResult> Renderers => new[] { Poppler, Ghostscript };
}

public sealed record RendererComparisonResult(
    RendererEngine Engine,
    int Width,
    int Height,
    int Dpi,
    long AePixels,
    double AePercent,
    double Mae,
    double Rmse,
    double Psnr,
    double Ssim,
    double Dssim,
    DiffStatus Status,
    string BaselineImagePath,
    string CandidateImagePath,
    string DiffImagePath,
    string BaselineThumbPath,
    string CandidateThumbPath,
    string DiffThumbPath,
    string OverlayImagePath,
    string OverlayThumbPath)
{
    public long TotalPixels => (long)Width * Height;
}

public enum RendererEngine
{
    Poppler,
    Ghostscript
}

public enum PageDiffClassification
{
    Clean,
    Consistent,
    EngineOnly
}

public sealed record QualityThresholds(
    double WarningAePercent,
    double WarningSsim,
    double PassSsimTolerance = 1e-6,
    double PassDssimTolerance = 1e-6);

public sealed record SsimOptions(double? Radius, double? Sigma, double? K1, double? K2)
{
    public bool HasValues =>
        Radius.HasValue || Sigma.HasValue || K1.HasValue || K2.HasValue;
}

public sealed record ToolMetadata(
    string? ImageMagickVersion,
    string? PopplerVersion,
    string? GhostscriptVersion,
    string? PdfBoxVersion);

public sealed record GeometryAnalysisResult(
    double TolerancePt,
    bool Passed,
    IReadOnlyList<GeometryPageReport> Pages);

public sealed record GeometryPageReport(
    int PageNumber,
    IReadOnlyList<WordDifference> MissingWords,
    IReadOnlyList<WordDifference> ExtraWords,
    IReadOnlyList<WordDifference> MovedWords,
    IReadOnlyList<GraphicDifference> AddedGraphics,
    IReadOnlyList<GraphicDifference> RemovedGraphics,
    IReadOnlyList<GraphicDifference> ChangedGraphics,
    GeometryWordStatus WordStatus,
    string? WordStatusNote)
{
    public bool IsClean =>
        WordStatus == GeometryWordStatus.Ok &&
        MissingWords.Count == 0 &&
        ExtraWords.Count == 0 &&
        MovedWords.Count == 0 &&
        AddedGraphics.Count == 0 &&
        RemovedGraphics.Count == 0 &&
        ChangedGraphics.Count == 0;
}

public sealed record WordDifference(
    string Text,
    BoundingBox? Baseline,
    BoundingBox? Candidate,
    GeometryDelta Delta,
    double FontSize);

public sealed record GraphicDifference(
    string Type,
    BoundingBox? Baseline,
    BoundingBox? Candidate,
    GeometryDelta Delta,
    double StrokeWidth,
    string? StrokeColor,
    string? FillColor);

public sealed record BoundingBox(double X, double Y, double Width, double Height)
{
    public double CenterX => X + (Width / 2d);
    public double CenterY => Y + (Height / 2d);

    public GeometryBox ToGeometryBox() => new(X, Y, Width, Height);
}

public enum GeometryWordStatus
{
    Ok = 0,
    NoText = 1,
    BaselineTextUnavailable = 2,
    CandidateTextUnavailable = 3,
    ComparisonUnreliable = 4
}

public enum DiffStatus
{
    Pass = 0,
    Warning = 1,
    Fail = 2
}
