using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace DocxToPdf.Tests.Pdf;

internal static class PdfRasterizer
{
    public static RasterizationResult Rasterize(string pdfPath, int dpi)
    {
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF not found", pdfPath);

        var tempDir = Path.Combine(Path.GetTempPath(), $"pdf-raster-{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);

        var prefixName = "page";
        var outputPrefix = Path.Combine(tempDir, prefixName);

        var arguments = $"-r {dpi} -png \"{pdfPath}\" \"{outputPrefix}\"";
        var psi = new ProcessStartInfo("pdftoppm", arguments)
        {
            CreateNoWindow = true,
            UseShellExecute = false,
            RedirectStandardError = true,
            RedirectStandardOutput = true
        };

        try
        {
            using var process = Process.Start(psi)
                ?? throw new InvalidOperationException("Unable to start pdftoppm.");

            var stdOut = process.StandardOutput.ReadToEnd();
            var stdErr = process.StandardError.ReadToEnd();
            process.WaitForExit();

            if (process.ExitCode != 0)
                throw new InvalidOperationException($"pdftoppm failed ({process.ExitCode}): {stdErr}{Environment.NewLine}{stdOut}");

            var files = Directory.GetFiles(tempDir, $"{prefixName}-*.png")
                .OrderBy(ExtractPageNumber)
                .ToArray();

            if (files.Length == 0)
                throw new InvalidOperationException("pdftoppm did not generate any rasterized pages.");

            return new RasterizationResult(tempDir, files);
        }
        catch
        {
            Directory.Delete(tempDir, recursive: true);
            throw;
        }
    }

    private static int ExtractPageNumber(string path)
    {
        var name = Path.GetFileNameWithoutExtension(path);
        var dashIndex = name.LastIndexOf('-');
        if (dashIndex < 0)
            return 0;

        var suffix = name[(dashIndex + 1)..];
        return int.TryParse(suffix, out var value) ? value : 0;
    }
}

internal sealed class RasterizationResult : IDisposable
{
    private readonly string _directory;

    public RasterizationResult(string directory, IReadOnlyList<string> imagePaths)
    {
        _directory = directory;
        ImagePaths = imagePaths;
    }

    public IReadOnlyList<string> ImagePaths { get; }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_directory))
                Directory.Delete(_directory, recursive: true);
        }
        catch
        {
            // Ignore cleanup failures
        }
    }
}
