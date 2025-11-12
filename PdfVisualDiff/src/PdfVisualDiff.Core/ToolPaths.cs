namespace PdfVisualDiff.Core;

public sealed record PixelDiffToolPaths(
    string? PdftoppmPath,
    string? ComparePath,
    string? IdentifyPath,
    string? MagickPath,
    string? GhostscriptPath);
