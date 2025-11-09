using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxToPdf.Sdk.Docx.Formatting;
using DocxToPdf.Sdk.Docx.Styles;
using System;
using System.Collections.Generic;
using System.Linq;

namespace DocxToPdf.Sdk.Docx.Numbering;

internal sealed class NumberingDefinitions
{
    private readonly Dictionary<int, AbstractNumberingDefinition> _abstract = new();
    private readonly Dictionary<int, NumberingInstanceDefinition> _instances = new();

    private NumberingDefinitions()
    {
    }

    public static NumberingDefinitions Load(WordprocessingDocument document)
    {
        var definitions = new NumberingDefinitions();
        var numberingPart = document.MainDocumentPart?.NumberingDefinitionsPart;
        if (numberingPart?.Numbering == null)
            return definitions;

        foreach (var abstractNum in numberingPart.Numbering.Elements<AbstractNum>())
        {
            var def = AbstractNumberingDefinition.FromOpenXml(abstractNum);
            definitions._abstract[def.Id] = def;
        }

        foreach (var number in numberingPart.Numbering.Elements<DocumentFormat.OpenXml.Wordprocessing.NumberingInstance>())
        {
            var def = NumberingInstanceDefinition.FromOpenXml(number);
            if (def != null)
                definitions._instances[def.Id] = def;
        }

        return definitions;
    }

    public bool TryGetInstance(int numId, out NumberingInstanceDefinition instance) =>
        _instances.TryGetValue(numId, out instance!);

    public bool TryGetAbstract(int abstractId, out AbstractNumberingDefinition definition) =>
        _abstract.TryGetValue(abstractId, out definition!);
}

internal sealed record AbstractNumberingDefinition(int Id, Dictionary<int, NumberingLevelDefinition> Levels)
{
    public static AbstractNumberingDefinition FromOpenXml(AbstractNum abstractNum)
    {
        var id = (int)(abstractNum.AbstractNumberId?.Value ?? 0);
        var levels = new Dictionary<int, NumberingLevelDefinition>();

        foreach (var level in abstractNum.Elements<Level>())
        {
            var def = NumberingLevelDefinition.FromOpenXml(level);
            levels[def.LevelIndex] = def;
        }

        return new AbstractNumberingDefinition(id, levels);
    }
}

internal sealed class NumberingInstanceDefinition
{
    public int Id { get; }
    public int AbstractId { get; }
    private readonly Dictionary<int, NumberingLevelOverride> _overrides = new();

    private NumberingInstanceDefinition(int id, int abstractId)
    {
        Id = id;
        AbstractId = abstractId;
    }

    public static NumberingInstanceDefinition? FromOpenXml(DocumentFormat.OpenXml.Wordprocessing.NumberingInstance instance)
    {
        var id = instance.NumberID?.Value;
        var abstractId = instance.AbstractNumId?.Val?.Value;
        if (!id.HasValue || !abstractId.HasValue)
            return null;

        var def = new NumberingInstanceDefinition((int)id.Value, (int)abstractId.Value);

        foreach (var lvlOverride in instance.Elements<LevelOverride>())
        {
            var ov = NumberingLevelOverride.FromOpenXml(lvlOverride);
            if (ov != null)
                def._overrides[ov.Level] = ov;
        }

        return def;
    }

    public NumberingLevelDefinition? ResolveLevel(NumberingDefinitions definitions, int level)
    {
        if (!definitions.TryGetAbstract(AbstractId, out var abstractDef))
            return null;

        abstractDef.Levels.TryGetValue(level, out var baseLevel);

        if (_overrides.TryGetValue(level, out var overrideLevel) && overrideLevel.LevelDefinition != null)
        {
            return overrideLevel.LevelDefinition.WithOverride(overrideLevel.StartOverride);
        }

        if (baseLevel == null)
            return null;

        if (_overrides.TryGetValue(level, out var @override))
        {
            return baseLevel.WithOverride(@override.StartOverride);
        }

        return baseLevel;
    }

    public int? GetStartOverride(int level)
    {
        return _overrides.TryGetValue(level, out var ov) ? ov.StartOverride : null;
    }
}

internal sealed record NumberingLevelOverride(int Level, int? StartOverride, NumberingLevelDefinition? LevelDefinition)
{
    public static NumberingLevelOverride? FromOpenXml(LevelOverride levelOverride)
    {
        var lvl = levelOverride.LevelIndex?.Value;
        if (!lvl.HasValue)
            return null;

        int? start = levelOverride.StartOverrideNumberingValue?.Val;
        NumberingLevelDefinition? definition = null;

        if (levelOverride.Level is Level overrideLevel)
        {
            definition = NumberingLevelDefinition.FromOpenXml(overrideLevel);
        }

        return new NumberingLevelOverride((int)lvl.Value, start, definition);
    }
}

internal sealed record NumberingLevelDefinition(
    int LevelIndex,
    NumberFormatValues NumberFormat,
    string LevelText,
    LevelSuffixValues Suffix,
    ParagraphAlignment Alignment,
    int StartNumber,
    ParagraphPropertySet? ParagraphProperties,
    RunPropertySet? RunProperties)
{
    public static NumberingLevelDefinition FromOpenXml(Level level)
    {
        var index = (int)(level.LevelIndex?.Value ?? 0);
        var numFmt = level.NumberingFormat?.Val?.Value ?? NumberFormatValues.Decimal;
        var lvlText = level.LevelText?.Val?.Value ?? "%1.";
        var suffix = level.LevelSuffix?.Val?.Value ?? LevelSuffixValues.Tab;
        var justification = level.LevelJustification?.Val?.Value;
        var alignment = ParagraphAlignment.Left;
        if (justification == LevelJustificationValues.Center)
            alignment = ParagraphAlignment.Center;
        else if (justification == LevelJustificationValues.Right)
            alignment = ParagraphAlignment.Right;
        var start = (int)(level.StartNumberingValue?.Val?.Value ?? 1);

        ParagraphPropertySet? paragraphProperties = null;
        var levelParagraphProps = ExtractParagraphProperties(level);
        if (levelParagraphProps != null)
        {
            paragraphProperties = ParagraphPropertySet.FromOpenXml(levelParagraphProps);
        }

        RunPropertySet? runProperties = null;
        if (RunPropertyHelpers.CloneRunProperties(level.GetFirstChild<NumberingSymbolRunProperties>()) is RunProperties lvlRunProps)
        {
            runProperties = RunPropertySet.FromOpenXml(lvlRunProps);
        }

        return new NumberingLevelDefinition(
            index,
            numFmt,
            lvlText,
            suffix,
            alignment,
            start,
            paragraphProperties,
            runProperties);
    }

    public NumberingLevelDefinition WithOverride(int? startOverride) =>
        startOverride.HasValue ? this with { StartNumber = startOverride.Value } : this;
    private static ParagraphProperties? ExtractParagraphProperties(Level level)
    {
        if (level.PreviousParagraphProperties is PreviousParagraphProperties prev)
        {
            var paraProps = new ParagraphProperties();
            foreach (var child in prev.ChildElements)
            {
                paraProps.Append(child.CloneNode(true));
            }
            return paraProps;
        }

        return null;
    }
}
