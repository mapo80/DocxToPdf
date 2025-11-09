namespace PdfVisualDiff.Tests;

internal static class TestPaths
{
    public static string RepoRoot { get; } = Path.GetFullPath(
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", ".."));

    public static string SolutionPath => Path.Combine(RepoRoot, "PdfVisualDiff.sln");
    public static string SamplesDirectory => Path.Combine(RepoRoot, "PdfVisualDiff", "samples");
    public static string CliProject => Path.Combine(RepoRoot, "PdfVisualDiff", "src", "PdfVisualDiff", "PdfVisualDiff.csproj");
}
