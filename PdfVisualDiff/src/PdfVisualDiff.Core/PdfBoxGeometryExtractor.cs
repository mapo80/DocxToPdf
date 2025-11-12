using System.Text.Json;

namespace PdfVisualDiff.Core;

internal sealed class PdfBoxGeometryExtractor
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web)
    {
        PropertyNameCaseInsensitive = true
    };

    private readonly string _javaPath;
    private readonly string _geometryJar;
    private readonly string _pdfboxJar;
    private readonly string _classPathArgument;

    public PdfBoxGeometryExtractor(string? toolsDirectory, string? javaPath = null)
    {
        if (string.IsNullOrWhiteSpace(toolsDirectory))
            throw new InvalidOperationException("PDFBox tools directory is not configured. Set --pdfbox-dir to point to the folder containing pdfbox-app-*.jar.");

        if (!Directory.Exists(toolsDirectory))
            throw new DirectoryNotFoundException($"PDFBox tools directory not found: {toolsDirectory}");

        _javaPath = string.IsNullOrWhiteSpace(javaPath) ? "java" : javaPath;

        _geometryJar = Path.Combine(toolsDirectory, "geometry-extractor.jar");
        if (!File.Exists(_geometryJar))
            throw new FileNotFoundException("Missing geometry-extractor.jar. Rebuild the Java helper under tools/pdfbox.", _geometryJar);

        _pdfboxJar = FindPdfBoxJar(toolsDirectory);
        _classPathArgument = string.Join(Path.PathSeparator, new[] { _geometryJar, _pdfboxJar });
    }

    public string PdfBoxVersion => Path.GetFileName(_pdfboxJar);

    public async Task<PdfBoxDocumentGeometry> ExtractAsync(
        string pdfPath,
        IReadOnlyList<int> pages,
        CancellationToken cancellationToken)
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"pdfgeom-{Guid.NewGuid():N}.json");
        try
        {
            var args = new List<string>
            {
                "-cp",
                _classPathArgument,
                "GeometryExtractor",
                "--pdf",
                pdfPath,
                "--output",
                tempFile
            };

            if (pages is { Count: > 0 })
            {
                args.Add("--pages");
                args.Add(string.Join(",", pages));
            }

            await ProcessRunner.RunAsync(_javaPath, args, cancellationToken: cancellationToken)
                .ConfigureAwait(false);

            await using var stream = File.OpenRead(tempFile);
            var document = await JsonSerializer.DeserializeAsync<PdfBoxDocumentGeometry>(stream, JsonOptions, cancellationToken)
                .ConfigureAwait(false);

            return document ?? new PdfBoxDocumentGeometry(Array.Empty<PdfBoxPageGeometry>());
        }
        finally
        {
            if (File.Exists(tempFile))
            {
                File.Delete(tempFile);
            }
        }
    }

    private static string FindPdfBoxJar(string directory)
    {
        var jar = Directory
            .EnumerateFiles(directory, "pdfbox-app-*.jar", SearchOption.TopDirectoryOnly)
            .OrderByDescending(Path.GetFileName)
            .FirstOrDefault();

        if (jar is null)
            throw new FileNotFoundException("pdfbox-app-*.jar not found in the specified directory.", directory);

        return jar;
    }
}
