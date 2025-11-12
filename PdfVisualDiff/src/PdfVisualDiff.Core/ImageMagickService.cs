using System.Globalization;

namespace PdfVisualDiff.Core;

internal sealed class ImageMagickService
{
    private static readonly IReadOnlySet<int> CompareAllowedExitCodes = new HashSet<int> { 0, 1 };

    private readonly string _comparePath;
    private readonly string _identifyPath;
    private readonly string _magickPath;
    private string? _magickVersion;

    public ImageMagickService(string? comparePath = null, string? identifyPath = null, string? magickPath = null)
    {
        _comparePath = string.IsNullOrWhiteSpace(comparePath) ? "compare" : comparePath;
        _identifyPath = string.IsNullOrWhiteSpace(identifyPath) ? "identify" : identifyPath;
        _magickPath = string.IsNullOrWhiteSpace(magickPath) ? "magick" : magickPath;
    }

    public async Task<string?> GetVersionAsync(CancellationToken cancellationToken = default)
    {
        if (_magickVersion != null)
            return _magickVersion;

        var result = await ProcessRunner.RunAsync(_magickPath, new[] { "-version" }, cancellationToken: cancellationToken)
            .ConfigureAwait(false);

        var line = result.StdOut.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
        _magickVersion = line ?? result.StdOut.Trim();
        return _magickVersion;
    }

    public async Task<(int Width, int Height)> IdentifyAsync(string imagePath, CancellationToken cancellationToken = default)
    {
        var args = new List<string>
        {
            "-format",
            "%w %h",
            imagePath
        };

        var result = await ProcessRunner.RunAsync(_identifyPath, args, cancellationToken: cancellationToken)
            .ConfigureAwait(false);

        var tokens = result.StdOut.Split(new[] { ' ', '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length < 2 ||
            !int.TryParse(tokens[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out var width) ||
            !int.TryParse(tokens[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var height))
        {
            throw new InvalidOperationException($"Unable to parse identify output for '{imagePath}': '{result.StdOut}'.");
        }

        return (width, height);
    }

    public async Task<ImageComparisonMetrics> CompareAsync(
        string baselineImage,
        string candidateImage,
        string diffOutput,
        string? fuzz,
        SsimOptions ssimOptions,
        CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(diffOutput)!);

        var rmse = await RunCompareAsync("RMSE", baselineImage, candidateImage, diffOutput, fuzz, ssimOptions, applySsimDefines: false, cancellationToken)
            .ConfigureAwait(false);
        var rmseValue = rmse.Second ?? rmse.First;

        var ae = await RunCompareAsync("AE", baselineImage, candidateImage, "null:", fuzz, ssimOptions, applySsimDefines: false, cancellationToken)
            .ConfigureAwait(false);
        var aePixels = (long)Math.Round(ae.First, MidpointRounding.AwayFromZero);

        var mae = await RunCompareAsync("MAE", baselineImage, candidateImage, "null:", fuzz, ssimOptions, applySsimDefines: false, cancellationToken)
            .ConfigureAwait(false);
        var maeValue = mae.Second ?? mae.First;

        var psnr = await RunCompareAsync("PSNR", baselineImage, candidateImage, "null:", fuzz, ssimOptions, applySsimDefines: false, cancellationToken)
            .ConfigureAwait(false);
        var psnrValue = psnr.First;

        var dssimValue = await ComputeSsimDssimAsync(baselineImage, candidateImage, fuzz, ssimOptions, cancellationToken)
            .ConfigureAwait(false);
        var ssimValue = double.IsNaN(dssimValue) ? double.NaN : Math.Clamp(1 - dssimValue, 0, 1);

        return new ImageComparisonMetrics(aePixels, maeValue, rmseValue, psnrValue, ssimValue, dssimValue);
    }

    public async Task CreateThumbnailAsync(
        string source,
        string destination,
        int maxDimension,
        CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(destination)!);

        var args = new List<string>
        {
            source,
            "-thumbnail",
            $"{maxDimension}x{maxDimension}",
            "-quality",
            "95",
            destination
        };

        await ProcessRunner.RunAsync(_magickPath, args, cancellationToken: cancellationToken)
            .ConfigureAwait(false);
    }

    public async Task PadToSizeAsync(
        string source,
        int width,
        int height,
        CancellationToken cancellationToken = default)
    {
        var tempPath = Path.Combine(
            Path.GetDirectoryName(source)!,
            $"{Path.GetFileNameWithoutExtension(source)}-{Guid.NewGuid():N}{Path.GetExtension(source)}");
        var args = new List<string>
        {
            source,
            "-background",
            "white",
            "-gravity",
            "northwest",
            "-extent",
            $"{width}x{height}",
            tempPath
        };

        await ProcessRunner.RunAsync(_magickPath, args, cancellationToken: cancellationToken)
            .ConfigureAwait(false);

        File.Delete(source);
        File.Move(tempPath, source);
    }

    public async Task CreateOverlayAsync(
        string baselineImage,
        string candidateImage,
        string destination,
        CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(destination)!);

        var args = new List<string>
        {
            "(",
            candidateImage,
            "-colorspace", "Gray",
            ")",
            "(",
            baselineImage,
            "-colorspace", "Gray",
            ")",
            "(",
            baselineImage,
            "-colorspace", "Gray",
            ")",
            "-combine",
            "-set", "colorspace", "sRGB",
            destination
        };

        await ProcessRunner.RunAsync(_magickPath, args, cancellationToken: cancellationToken)
            .ConfigureAwait(false);
    }

    public async Task<string> RenderPdfPageAsync(
        string pdfPath,
        int pageNumber,
        string destinationPath,
        int dpi,
        bool useCropBox,
        string? password,
        CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(destinationPath)!);

        var args = new List<string>
        {
            "-density",
            dpi.ToString(CultureInfo.InvariantCulture)
        };

        if (useCropBox)
        {
            args.Add("-define");
            args.Add("pdf:use-cropbox=true");
        }

        if (!string.IsNullOrWhiteSpace(password))
        {
            args.Add("-authenticate");
            args.Add(password!);
        }

        args.Add(pdfPath + $"[{pageNumber - 1}]");
        args.Add("+repage");
        args.Add("-background");
        args.Add("white");
        args.Add("-alpha");
        args.Add("remove");
        args.Add("-alpha");
        args.Add("off");
        args.Add("-colorspace");
        args.Add("sRGB");
        args.Add(destinationPath);

        await ProcessRunner.RunAsync(_magickPath, args, cancellationToken: cancellationToken)
            .ConfigureAwait(false);

        return destinationPath;
    }

    private async Task<(double First, double? Second)> RunCompareAsync(
        string metric,
        string baseline,
        string candidate,
        string output,
        string? fuzz,
        SsimOptions ssimOptions,
        bool applySsimDefines,
        CancellationToken cancellationToken)
    {
        var args = new List<string>();

        if (applySsimDefines && ssimOptions is not null && ssimOptions.HasValues)
        {
            AddSsimDefines(args, ssimOptions);
        }

        if (!string.IsNullOrWhiteSpace(fuzz))
        {
            args.Add("-fuzz");
            args.Add(fuzz!);
        }

        args.Add("-metric");
        args.Add(metric);
        args.Add(baseline);
        args.Add(candidate);
        args.Add(output);

        var result = await ProcessRunner.RunAsync(
                _comparePath,
                args,
                allowedExitCodes: CompareAllowedExitCodes,
                cancellationToken: cancellationToken)
            .ConfigureAwait(false);

        return ParseMetricValues(result.StdErr, metric);
    }

    private async Task<double> ComputeSsimDssimAsync(
        string baseline,
        string candidate,
        string? fuzz,
        SsimOptions ssimOptions,
        CancellationToken cancellationToken)
    {
        var args = new List<string>
        {
            baseline,
            candidate
        };

        if (ssimOptions is not null && ssimOptions.HasValues)
        {
            AddSsimDefines(args, ssimOptions);
        }

        if (!string.IsNullOrWhiteSpace(fuzz))
        {
            args.Add("-fuzz");
            args.Add(fuzz!);
        }

        args.Add("-metric");
        args.Add("SSIM");
        args.Add("-compare");
        args.Add("-format");
        args.Add("%[distortion]");
        args.Add("info:");

        var result = await ProcessRunner.RunAsync(
                _magickPath,
                args,
                allowedExitCodes: CompareAllowedExitCodes,
                cancellationToken: cancellationToken)
            .ConfigureAwait(false);

        if (double.TryParse(result.StdOut.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
            return value;

        throw new InvalidOperationException($"Unable to parse SSIM/DSSIM value from: '{result.StdOut}'.");
    }

    private static void AddSsimDefines(List<string> args, SsimOptions ssimOptions)
    {
        AddSsimDefine(args, "radius", ssimOptions.Radius);
        AddSsimDefine(args, "sigma", ssimOptions.Sigma);
        AddSsimDefine(args, "k1", ssimOptions.K1);
        AddSsimDefine(args, "k2", ssimOptions.K2);
    }

    private static void AddSsimDefine(List<string> args, string name, double? value)
    {
        if (!value.HasValue)
            return;

        args.Add("-define");
        args.Add($"compare:ssim-{name}={value.Value.ToString("G", CultureInfo.InvariantCulture)}");
    }

    private static (double First, double? Second) ParseMetricValues(string input, string metric)
    {
        var values = new List<double>();
        var tokens = input.Split(new[] { ' ', '\r', '\n', '\t', '(', ')', ':' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var token in tokens)
        {
            if (token.Equals("inf", StringComparison.OrdinalIgnoreCase))
            {
                values.Add(double.PositiveInfinity);
                continue;
            }

            if (double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
            {
                values.Add(value);
                continue;
            }
        }

        if (values.Count == 0)
            throw new InvalidOperationException($"Unable to parse {metric} value from: '{input}'.");

        return (values[0], values.Count > 1 ? values[1] : null);
    }
}

internal sealed record ImageComparisonMetrics(long AePixels, double Mae, double Rmse, double Psnr, double Ssim, double Dssim);
