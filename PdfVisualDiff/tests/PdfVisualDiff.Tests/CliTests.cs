using System.Diagnostics;
using System.Text;
using System.Text.Json;
using Xunit;

namespace PdfVisualDiff.Tests;

 [Collection("Cli")]
public sealed class CliTests
{
    private static readonly string? DiffPdfPath = GetToolPath("diff-pdf");
    private static readonly bool ToolAvailable = DiffPdfPath != null;

    [Fact]
    public void IdenticalPdfsPass()
    {
        if (!ToolAvailable)
            return;

        using var temp = new TempDirectory();
        var pdf = Path.Combine(TestPaths.SamplesDirectory, "A.pdf");

        var result = RunCli(pdf, pdf, temp.Path);
        Assert.True(result.ExitCode == 0, $"CLI exited with {result.ExitCode}. STDOUT: {result.StdOut} STDERR: {result.StdErr}");

        var report = ReadReport(Path.Combine(temp.Path, "report.json"));
        Assert.True(report.GetProperty("passed").GetBoolean());
        Assert.Equal(0, report.GetProperty("difference").GetProperty("percent").GetDouble());
        Assert.Equal(JsonValueKind.Array, report.GetProperty("pageStats").ValueKind);
    }

    [Fact]
    public void DifferentPdfsFail()
    {
        if (!ToolAvailable)
            return;

        using var temp = new TempDirectory();
        var baseline = Path.Combine(TestPaths.SamplesDirectory, "A.pdf");
        var candidate = Path.Combine(TestPaths.SamplesDirectory, "B.pdf");

        var result = RunCli(baseline, candidate, temp.Path);
        Assert.True(result.ExitCode == 1, $"CLI exited with {result.ExitCode}. STDOUT: {result.StdOut} STDERR: {result.StdErr}");

        var report = ReadReport(Path.Combine(temp.Path, "report.json"));
        Assert.False(report.GetProperty("passed").GetBoolean());
        Assert.NotNull(report.GetProperty("diffPdf").GetString());
        Assert.True(report.GetProperty("difference").GetProperty("percent").GetDouble() > 0);
    }

    [Fact]
    public void GeometryCheckProducesReport()
    {
        if (!ToolAvailable)
            return;

        using var temp = new TempDirectory();
        var pdf = Path.Combine(TestPaths.SamplesDirectory, "A.pdf");

        var result = RunCli(pdf, pdf, temp.Path, "--geometry-check");
        Assert.Equal(0, result.ExitCode);

        var report = ReadReport(Path.Combine(temp.Path, "report.json"));
        Assert.True(report.TryGetProperty("geometry", out var geometry));
        Assert.True(geometry.GetProperty("passed").GetBoolean());
        Assert.Equal(0.25, geometry.GetProperty("tolerancePt").GetDouble(), 3);
    }

    private static CliRunResult RunCli(string baseline, string candidate, string output, params string[] extraArgs)
    {
        var psi = new ProcessStartInfo("dotnet")
        {
            WorkingDirectory = TestPaths.RepoRoot,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };

        psi.ArgumentList.Add("run");
        psi.ArgumentList.Add("--project");
        psi.ArgumentList.Add(TestPaths.CliProject);
        psi.ArgumentList.Add(baseline);
        psi.ArgumentList.Add(candidate);
        psi.ArgumentList.Add("--out");
        psi.ArgumentList.Add(output);

        foreach (var arg in extraArgs)
        {
            psi.ArgumentList.Add(arg);
        }

        if (!string.IsNullOrEmpty(DiffPdfPath))
        {
            psi.ArgumentList.Add("--diff-pdf-path");
            psi.ArgumentList.Add(DiffPdfPath!);
        }

        using var process = Process.Start(psi)!;

        var stdout = new StringBuilder();
        var stderr = new StringBuilder();
        process.OutputDataReceived += (_, e) => { if (e.Data != null) stdout.AppendLine(e.Data); };
        process.ErrorDataReceived += (_, e) => { if (e.Data != null) stderr.AppendLine(e.Data); };
        process.BeginOutputReadLine();
        process.BeginErrorReadLine();

        process.WaitForExit();

        return new CliRunResult(process.ExitCode, stdout.ToString(), stderr.ToString());
    }

    private static JsonElement ReadReport(string path)
    {
        using var stream = File.OpenRead(path);
        using var document = JsonDocument.Parse(stream);
        return document.RootElement.Clone();
    }

    private static string? GetToolPath(string name)
    {
        var psi = new ProcessStartInfo("which", name)
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        try
        {
            using var process = Process.Start(psi);
            if (process == null)
                return null;
            var output = process.StandardOutput.ReadToEnd().Trim();
            process.WaitForExit();
            return process.ExitCode == 0 ? output : null;
        }
        catch
        {
            return null;
        }
    }
}

internal readonly record struct CliRunResult(int ExitCode, string StdOut, string StdErr);

[CollectionDefinition("Cli", DisableParallelization = true)]
public sealed class CliCollection : ICollectionFixture<object>;

internal sealed class TempDirectory : IDisposable
{
    public string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"pdfdiff-{Guid.NewGuid():N}");

    public TempDirectory()
    {
        Directory.CreateDirectory(Path);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(Path))
                Directory.Delete(Path, recursive: true);
        }
        catch
        {
            // ignore
        }
    }
}
