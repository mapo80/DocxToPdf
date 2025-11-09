using DocxToPdf.Sdk;
using System;
using System.IO;

namespace DocxToPdf.Tests.Helpers;

internal static class PdfTestHelper
{
    public static string RenderWithSdk(string docxPath)
    {
        var outputPath = Path.Combine(Path.GetTempPath(), $"docx-sdk-{Guid.NewGuid():N}.pdf");
        var converter = new DocxToPdfConverter();
        converter.Convert(docxPath, outputPath);
        return outputPath;
    }
}
