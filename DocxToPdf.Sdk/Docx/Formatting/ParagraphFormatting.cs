using System;
using System.Collections.Generic;

namespace DocxToPdf.Sdk.Docx.Formatting;

/// <summary>
/// Proprietà tipografiche di un paragrafo risolte dopo il cascading di Word.
/// </summary>
public sealed record ParagraphFormatting
{
    public static readonly ParagraphFormatting Default = new();

    public float SpacingBeforePt { get; init; }
    public float SpacingAfterPt { get; init; }
    public float? LineSpacingPt { get; init; }
    public ParagraphAlignment Alignment { get; init; } = ParagraphAlignment.Left;
    public float LeftIndentPt { get; init; }
    public float RightIndentPt { get; init; }
    public float FirstLineIndentPt { get; init; }
    public float HangingIndentPt { get; init; }
    public IReadOnlyList<TabStopDefinition> TabStops { get; init; } = Array.Empty<TabStopDefinition>();

    /// <summary>
    /// Calcola l'indentazione applicata alla prima riga considerando special e hanging.
    /// </summary>
    public float GetFirstLineOffsetPt()
    {
        if (HangingIndentPt > 0)
        {
            var offset = LeftIndentPt - HangingIndentPt;
            return offset < 0 ? 0 : offset;
        }

        return LeftIndentPt + FirstLineIndentPt;
    }

    /// <summary>
    /// Indentazione applicata alle righe successive alla prima.
    /// </summary>
    public float GetSubsequentLineOffsetPt()
    {
        if (HangingIndentPt > 0)
        {
            return LeftIndentPt;
        }

        return LeftIndentPt;
    }
}
