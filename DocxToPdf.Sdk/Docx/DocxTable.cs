using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using DocxToPdf.Sdk.Docx.Formatting;
using DocxToPdf.Sdk.Units;
using System;
using System.Collections.Generic;
using System.Linq;

namespace DocxToPdf.Sdk.Docx;

internal sealed record DocxTable(
    IReadOnlyList<float> ColumnGridPt,
    IReadOnlyList<DocxTableRow> Rows,
    TablePreferredWidth PreferredWidth,
    TableLayoutType LayoutType,
    TablePadding DefaultCellPadding,
    TableBorderSet Borders)
{
    public static DocxTable? FromOpenXmlTable(
        Table table,
        Styles.DocxStyleResolver styleResolver,
        Numbering.NumberingResolver numberingResolver,
        float defaultTabStopPt,
        char decimalSymbol)
    {
        var grid = table.GetFirstChild<TableGrid>();
        var gridColumns = grid?.Elements<GridColumn>().ToList() ?? new List<GridColumn>();
        if (gridColumns.Count == 0)
            return null;

        var tableProps = table.GetFirstChild<TableProperties>();
        var preferredWidth = ParsePreferredWidth(tableProps?.GetFirstChild<TableWidth>()) ?? TablePreferredWidth.Auto;
        var layout = ParseLayout(tableProps?.GetFirstChild<TableLayout>());
        var defaultPaddingOverride = ParsePadding(tableProps?.GetFirstChild<TableCellMarginDefault>());
        var defaultPadding = TablePadding.Default.ApplyOverride(defaultPaddingOverride);
        var tableBorders = ParseBorders(tableProps?.GetFirstChild<TableBorders>()) ?? TableBorderSet.None;

        var columnWidths = new List<float>();
        foreach (var column in gridColumns)
        {
            int dxa = 0;
            var raw = column.Width?.Value;
            if (!string.IsNullOrEmpty(raw) && int.TryParse(raw, out var parsed))
                dxa = parsed;
            columnWidths.Add(UnitConverter.DxaToPoints(dxa));
        }

        var rows = new List<DocxTableRow>();
        foreach (var row in table.Elements<TableRow>())
        {
            var cells = new List<DocxTableCell>();
            int colIndex = 0;
            foreach (var cell in row.Elements<TableCell>())
            {
                int span = 1;
                var gridSpan = cell.GetFirstChild<TableCellProperties>()?.GridSpan;
                if (gridSpan?.Val?.Value > 1)
                    span = gridSpan.Val.Value;

                var width = 0f;
                for (int s = 0; s < span && colIndex + s < columnWidths.Count; s++)
                    width += columnWidths[colIndex + s];

                var cellProps = cell.GetFirstChild<TableCellProperties>();
                var cellPadding = ParsePadding(cellProps?.GetFirstChild<TableCellMargin>()) ??
                    ParsePadding(cellProps?.GetFirstChild<TableCellMarginDefault>());
                var preferredCellWidth = ParsePreferredWidth(cellProps?.GetFirstChild<TableCellWidth>());
                var vMerge = ParseVerticalMerge(cellProps?.GetFirstChild<VerticalMerge>());
                var vAlign = ParseVerticalAlignment(cellProps?.GetFirstChild<TableCellVerticalAlignment>());
                var cellBorders = ParseBorders(cellProps?.GetFirstChild<TableCellBorders>());

                var paragraphs = cell.Elements<Paragraph>()
                    .Select(p => DocxParagraph.FromParagraph(
                        p,
                        styleResolver,
                        numberingResolver,
                        defaultTabStopPt,
                        decimalSymbol,
                        isInTable: true))
                    .ToArray();
                if (paragraphs.Length == 0)
                {
                    paragraphs = new[]
                    {
                        DocxParagraph.FromParagraph(
                            new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(string.Empty))),
                            styleResolver,
                            numberingResolver,
                            defaultTabStopPt,
                            decimalSymbol,
                            isInTable: true)
                    };
                }

                var cellModel = new DocxTableCell(
                    colIndex,
                    span,
                    width,
                    paragraphs,
                    preferredCellWidth,
                    cellPadding,
                    vAlign,
                    vMerge,
                    cellBorders);

                cells.Add(cellModel);
                colIndex += span;
            }

            if (cells.Count > 0)
                rows.Add(new DocxTableRow(cells));
        }

        if (rows.Count == 0)
            return null;

        return new DocxTable(
            columnWidths,
            rows,
            preferredWidth,
            layout,
            defaultPadding,
            tableBorders);
    }

    private static TablePreferredWidth? ParsePreferredWidth(TableWidth? width)
    {
        if (width == null)
            return null;

        var unit = width.Type != null ? width.Type.Value : (TableWidthUnitValues?)null;
        if (unit == TableWidthUnitValues.Dxa)
            return new TablePreferredWidth(TableWidthUnit.Dxa, ParseTwipString(width.Width));
        if (unit == TableWidthUnitValues.Pct)
            return new TablePreferredWidth(TableWidthUnit.Percent, ParseTwipString(width.Width));

        return TablePreferredWidth.Auto;
    }

    private static TablePreferredWidth? ParsePreferredWidth(TableCellWidth? width)
    {
        if (width == null)
            return null;

        var unit = width.Type != null ? width.Type.Value : (TableWidthUnitValues?)null;
        if (unit == TableWidthUnitValues.Dxa)
            return new TablePreferredWidth(TableWidthUnit.Dxa, ParseTwipString(width.Width));
        if (unit == TableWidthUnitValues.Pct)
            return new TablePreferredWidth(TableWidthUnit.Percent, ParseTwipString(width.Width));

        return null;
    }

    private static TableLayoutType ParseLayout(TableLayout? layout) =>
        layout?.Type?.Value == TableLayoutValues.Fixed
            ? TableLayoutType.Fixed
            : TableLayoutType.Auto;

    private static TablePaddingOverride? ParsePadding(OpenXmlElement? marginElement)
    {
        if (marginElement == null)
            return null;

        var top = ParseMarginValue(FindMarginElement(marginElement, "top"));
        var right = ParseMarginValue(FindMarginElement(marginElement, "right", "end"));
        var bottom = ParseMarginValue(FindMarginElement(marginElement, "bottom"));
        var left = ParseMarginValue(FindMarginElement(marginElement, "left", "start"));

        if (top == null && right == null && bottom == null && left == null)
            return null;

        return new TablePaddingOverride(top, right, bottom, left);
    }

    private static OpenXmlElement? FindMarginElement(OpenXmlElement? parent, params string[] localNames)
    {
        if (parent == null || localNames.Length == 0)
            return null;

        foreach (var child in parent.ChildElements)
        {
            foreach (var name in localNames)
            {
                if (string.Equals(child.LocalName, name, StringComparison.OrdinalIgnoreCase))
                    return child;
            }
        }

        return null;
    }

    private static float? ParseMarginValue(OpenXmlElement? margin)
    {
        if (margin == null)
            return null;

        string? widthValue = null;
        string? typeValue = null;
        foreach (var attribute in margin.GetAttributes())
        {
            if (string.Equals(attribute.LocalName, "w", StringComparison.OrdinalIgnoreCase))
                widthValue = attribute.Value;
            else if (string.Equals(attribute.LocalName, "type", StringComparison.OrdinalIgnoreCase))
                typeValue = attribute.Value;
        }

        var unit = TableWidthUnitValues.Dxa;
        if (!string.IsNullOrEmpty(typeValue) && Enum.TryParse<TableWidthUnitValues>(typeValue, true, out var parsedUnit))
            unit = parsedUnit;

        if (unit == TableWidthUnitValues.Nil)
            return 0f;

        if (string.IsNullOrEmpty(widthValue) || !int.TryParse(widthValue, out var twips))
            return null;

        return UnitConverter.DxaToPoints(twips);
    }

    private static TableVerticalMerge ParseVerticalMerge(VerticalMerge? merge)
    {
        var state = merge != null && merge.Val != null ? merge.Val.Value : (MergedCellValues?)null;
        if (state == MergedCellValues.Restart)
            return TableVerticalMerge.Restart;
        if (state == MergedCellValues.Continue)
            return TableVerticalMerge.Continue;
        return TableVerticalMerge.None;
    }

    private static VerticalAlignment ParseVerticalAlignment(TableCellVerticalAlignment? align)
    {
        var value = align != null && align.Val != null ? align.Val.Value : (TableVerticalAlignmentValues?)null;
        if (value == TableVerticalAlignmentValues.Center)
            return VerticalAlignment.Center;
        if (value == TableVerticalAlignmentValues.Bottom)
            return VerticalAlignment.Bottom;
        return VerticalAlignment.Top;
    }

    private static TableBorderSet? ParseBorders(OpenXmlElement? element)
    {
        if (element == null)
            return null;

        return new TableBorderSet(
            ParseBorder(element.GetFirstChild<TopBorder>()),
            ParseBorder(element.GetFirstChild<LeftBorder>()),
            ParseBorder(element.GetFirstChild<BottomBorder>()),
            ParseBorder(element.GetFirstChild<RightBorder>()),
            ParseBorder(element.GetFirstChild<InsideHorizontalBorder>()),
            ParseBorder(element.GetFirstChild<InsideVerticalBorder>()));
    }

    private static TableBorderStyle? ParseBorder(BorderType? border)
    {
        if (border == null)
            return null;

        var size = border.Size?.Value ?? 0;
        if (size <= 0)
            return null;

        var widthPt = size / 8f;
        var color = border.Color?.Value ?? "B4B4B4";
        var rgb = HexToColor(color);
        return new TableBorderStyle(widthPt, rgb);
    }

    private static RgbColor HexToColor(string hex)
    {
        if (hex == "auto" || hex.Length < 6)
            return RgbColor.Black;

        var r = byte.Parse(hex.Substring(0, 2), System.Globalization.NumberStyles.HexNumber);
        var g = byte.Parse(hex.Substring(2, 2), System.Globalization.NumberStyles.HexNumber);
        var b = byte.Parse(hex.Substring(4, 2), System.Globalization.NumberStyles.HexNumber);
        return new RgbColor(r, g, b);
    }

    private static int ParseTwipString(StringValue? value) =>
        value != null && int.TryParse(value.Value, out var result) ? result : 0;
}

internal sealed record DocxTableRow(IReadOnlyList<DocxTableCell> Cells);

internal sealed record DocxTableCell(
    int ColumnIndex,
    int ColumnSpan,
    float WidthPt,
    IReadOnlyList<DocxParagraph> Paragraphs,
    TablePreferredWidth? PreferredWidth,
    TablePaddingOverride? PaddingOverride,
    VerticalAlignment VerticalAlignment,
    TableVerticalMerge VerticalMerge,
    TableBorderSet? Borders);

internal enum TableLayoutType
{
    Auto,
    Fixed
}

internal enum TableWidthUnit
{
    Auto,
    Dxa,
    Percent
}

internal sealed record TablePreferredWidth(TableWidthUnit Unit, int Value)
{
    public static readonly TablePreferredWidth Auto = new(TableWidthUnit.Auto, 0);
}

internal sealed record TablePadding(float Top, float Right, float Bottom, float Left)
{
    // Word default cell margins (tblCellMar default): Left/Right = 108 twips (≈5.4pt), Top/Bottom = 0
    public static readonly TablePadding Default = new(0f, 5.4f, 0f, 5.4f);

    public TablePadding ApplyOverride(TablePaddingOverride? overridePadding)
    {
        if (overridePadding == null)
            return this;

        return new TablePadding(
            overridePadding.Top ?? Top,
            overridePadding.Right ?? Right,
            overridePadding.Bottom ?? Bottom,
            overridePadding.Left ?? Left);
    }
}

internal sealed record TablePaddingOverride(float? Top, float? Right, float? Bottom, float? Left);

internal sealed record TableBorderSet(
    TableBorderStyle? Top,
    TableBorderStyle? Left,
    TableBorderStyle? Bottom,
    TableBorderStyle? Right,
    TableBorderStyle? InsideHorizontal,
    TableBorderStyle? InsideVertical)
{
    public static readonly TableBorderSet None = new(null, null, null, null, null, null);
}

internal sealed record TableBorderStyle(float WidthPt, RgbColor Color);

internal enum VerticalAlignment
{
    Top,
    Center,
    Bottom
}

internal enum TableVerticalMerge
{
    None,
    Restart,
    Continue
}
