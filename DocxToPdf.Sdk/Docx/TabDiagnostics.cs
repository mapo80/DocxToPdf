using System;

namespace DocxToPdf.Sdk.Docx;

internal static class TabDiagnostics
{
    public static Action<string>? Sink { get; set; }

    public static void Write(string message)
    {
        Sink?.Invoke($"[tabs] {message}");
    }
}
