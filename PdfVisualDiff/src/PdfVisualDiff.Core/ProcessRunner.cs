using System.ComponentModel;
using System.Diagnostics;

namespace PdfVisualDiff.Core;

internal static class ProcessRunner
{
    public static async Task<ProcessResult> RunAsync(
        string fileName,
        IReadOnlyList<string> arguments,
        IEnumerable<int>? allowedExitCodes = null,
        CancellationToken cancellationToken = default)
    {
        var psi = new ProcessStartInfo(fileName)
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };

        foreach (var argument in arguments)
            psi.ArgumentList.Add(argument);

        try
        {
            using var process = Process.Start(psi)
                ?? throw new InvalidOperationException($"Unable to start process '{fileName}'.");

            var stdoutTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
            var stderrTask = process.StandardError.ReadToEndAsync(cancellationToken);

            await process.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
            var stdout = await stdoutTask.ConfigureAwait(false);
            var stderr = await stderrTask.ConfigureAwait(false);

            if (allowedExitCodes != null)
            {
                var allowed = allowedExitCodes as IReadOnlySet<int> ?? new HashSet<int>(allowedExitCodes);
                if (!allowed.Contains(process.ExitCode))
                {
                    throw new InvalidOperationException(
                        $"Tool '{fileName}' exited with code {process.ExitCode}.{Environment.NewLine}{stderr}");
                }
            }
            else if (process.ExitCode != 0)
            {
                throw new InvalidOperationException(
                    $"Tool '{fileName}' exited with code {process.ExitCode}.{Environment.NewLine}{stderr}");
            }

            return new ProcessResult(process.ExitCode, stdout, stderr);
        }
        catch (Win32Exception ex)
        {
            throw new InvalidOperationException(
                $"Unable to start '{fileName}'. Ensure it is installed and available on PATH.",
                ex);
        }
    }
}

internal sealed record ProcessResult(int ExitCode, string StdOut, string StdErr);
