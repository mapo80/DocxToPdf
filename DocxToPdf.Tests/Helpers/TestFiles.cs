using System;
using System.IO;

namespace DocxToPdf.Tests.Helpers;

internal static class TestFiles
{
    private static readonly string RepoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));

    public static string GetSamplePath(string relative) =>
        Path.Combine(RepoRoot, "samples", relative);

    public static string GetGoldenPath(string relative) =>
        Path.Combine(RepoRoot, "samples", "golden", relative);
}
