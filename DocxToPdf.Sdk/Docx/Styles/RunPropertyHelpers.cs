using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxToPdf.Sdk.Docx.Styles;

internal static class RunPropertyHelpers
{
    public static RunProperties? CloneRunProperties(OpenXmlCompositeElement? source)
    {
        if (source == null)
            return null;

        var runProps = new RunProperties();
        foreach (var child in source.ChildElements)
        {
            runProps.Append(child.CloneNode(true));
        }

        return runProps.HasChildren ? runProps : null;
    }
}
