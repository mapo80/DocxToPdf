using System.Globalization;

namespace PdfVisualDiff.Core;

internal sealed class PdfRasterizer
{
    private readonly string _toolPath;
    private string? _version;

    public PdfRasterizer(string? toolPath = null)
    {
        _toolPath = string.IsNullOrWhiteSpace(toolPath) ? "pdftoppm" : toolPath;
    }

    public async Task<string?> GetVersionAsync(CancellationToken cancellationToken = default)
    {
        if (_version != null)
            return _version;

        try
        {
            var result = await ProcessRunner.RunAsync(_toolPath, new[] { "-v" }, cancellationToken: cancellationToken)
                .ConfigureAwait(false);
            _version = NormalizeVersionOutput(result.StdOut, result.StdErr);
        }
        catch
        {
            _version = null;
        }

        return _version;
    }

    public async Task<string> RasterizeAsync(
        RasterizeRequest request,
        CancellationToken cancellationToken = default)
    {
        if (!File.Exists(request.PdfPath))
            throw new FileNotFoundException("PDF not found", request.PdfPath);

        var outputDir = Path.GetDirectoryName(request.OutputPrefix)
            ?? request.OutputPrefix;
        Directory.CreateDirectory(outputDir);

        var outputPath = request.OutputPrefix + ".png";

        var args = new List<string>
        {
            "-png",
            "-singlefile",
            "-r", request.Dpi.ToString(CultureInfo.InvariantCulture),
            "-aa", "yes",
            "-aaVector", "yes",
            "-f", request.PageNumber.ToString(CultureInfo.InvariantCulture),
            "-l", request.PageNumber.ToString(CultureInfo.InvariantCulture)
        };

        if (request.UseCropBox)
            args.Add("-cropbox");

        if (!string.IsNullOrEmpty(request.OwnerPassword))
        {
            args.Add("-opw");
            args.Add(request.OwnerPassword!);
        }

        if (!string.IsNullOrEmpty(request.UserPassword))
        {
            args.Add("-upw");
            args.Add(request.UserPassword!);
        }

        args.Add(request.PdfPath);
        args.Add(request.OutputPrefix);

        await ProcessRunner.RunAsync(_toolPath, args, cancellationToken: cancellationToken)
            .ConfigureAwait(false);

        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"pdftoppm did not produce '{outputPath}'.");

        return outputPath;
    }

    private static string? NormalizeVersionOutput(string stdout, string stderr)
    {
        var text = string.IsNullOrWhiteSpace(stdout) ? stderr : stdout;
        var line = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
        return line?.Trim();
    }
}

internal sealed record RasterizeRequest(
    string PdfPath,
    int PageNumber,
    string OutputPrefix,
    int Dpi,
    bool UseCropBox,
    string? UserPassword,
    string? OwnerPassword);
