using DocumentFormat.OpenXml.Wordprocessing;
using DocxToPdf.Sdk.Docx.Formatting;
using DocxToPdf.Sdk.Docx.Styles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace DocxToPdf.Sdk.Docx.Numbering;

internal sealed class NumberingResolver
{
    private readonly NumberingDefinitions _definitions;
    private readonly Dictionary<int, NumberingCounterState> _states = new();
    private readonly HashSet<(int NumId, int Level)> _appliedStartOverrides = new();

    public NumberingResolver(NumberingDefinitions definitions)
    {
        _definitions = definitions;
    }

    public NumberingResult? Resolve(Paragraph paragraph, ParagraphPropertySet baseParagraphProps, DocxStyleResolver styleResolver)
    {
        var numPr = paragraph.ParagraphProperties?.NumberingProperties;
        int? numId = numPr?.NumberingId?.Val?.Value;
        int? level = numPr?.NumberingLevelReference?.Val?.Value;

        if (!numId.HasValue)
            numId = baseParagraphProps.NumberingId;
        if (!level.HasValue)
            level = baseParagraphProps.NumberingLevel;

        if (!numId.HasValue)
            return null;

        var effectiveLevel = (int)(level ?? 0);
        if (!_definitions.TryGetInstance((int)numId.Value, out var instance))
            return null;

        var levelDef = instance.ResolveLevel(_definitions, effectiveLevel);
        if (levelDef == null)
            return null;

        if (!_states.TryGetValue(instance.AbstractId, out var state))
        {
            state = new NumberingCounterState();
            _states[instance.AbstractId] = state;
        }
        state.ResetLevelsAbove(effectiveLevel);

        var counter = state.GetCounter(effectiveLevel);

        var startOverride = instance.GetStartOverride(effectiveLevel);
        if (startOverride.HasValue && _appliedStartOverrides.Add((instance.Id, effectiveLevel)))
        {
            counter.ForceStart(startOverride.Value);
        }
        var start = levelDef.StartNumber;
        var value = counter.Next(start);

        ApplyIndentOverrides(levelDef, baseParagraphProps);

        var text = BuildLevelText(instance, effectiveLevel, levelDef, value, state);
        var runFormatting = levelDef.RunProperties?.ToFormatting(styleResolver) ?? RunFormatting.Default;

        NumberingDiagnostics.Write($"numId={instance.Id}, ilvl={level}, value={value}, text='{text}'");

        return new NumberingResult(new DocxListMarker(text, runFormatting, levelDef.Alignment, levelDef.Suffix));
    }

    private static void ApplyIndentOverrides(NumberingLevelDefinition levelDef, ParagraphPropertySet paragraphProps)
    {
        if (levelDef.ParagraphProperties == null)
            return;

        paragraphProps.Apply(levelDef.ParagraphProperties);
    }

    private string BuildLevelText(NumberingInstanceDefinition instance, int currentLevel, NumberingLevelDefinition levelDef, int currentValue, NumberingCounterState state)
    {
        if (levelDef.NumberFormat == NumberFormatValues.Bullet)
        {
            return NormalizeBulletText(levelDef.LevelText);
        }

        var pattern = levelDef.LevelText ?? "%1.";
        var regex = new Regex("%[1-9]");

        string Replace(Match match)
        {
            var idx = match.Value[1] - '1';
            var counter = state.TryGetCounter(idx);
            if (counter == null || !counter.HasValue)
            {
                return string.Empty;
            }

            var lvlDef = instance.ResolveLevel(_definitions, idx);
            var format = lvlDef?.NumberFormat ?? levelDef.NumberFormat;
            return NumberFormatConverter.Format(counter.Current, format);
        }

        return regex.Replace(pattern, Replace);
    }

    private static string NormalizeBulletText(string text)
    {
        if (string.IsNullOrEmpty(text))
            return "\u2022";

        var buffer = text.ToCharArray();
        bool modified = false;
        for (int i = 0; i < buffer.Length; i++)
        {
            var mapped = buffer[i] switch
            {
                '\uf0b7' => '\u2022', // Wingdings bullet
                '\uf0d8' => '\u25C6', // diamond
                '\uf0a7' => '\u25AA', // square bullet
                _ => buffer[i]
            };
            if (mapped != buffer[i])
            {
                buffer[i] = mapped;
                modified = true;
            }
        }

        return modified ? new string(buffer) : text;
    }
}

internal sealed record NumberingResult(DocxListMarker Marker);

internal sealed class NumberingCounter
{
    public int Current { get; private set; }
    public bool HasValue { get; private set; }
    private int? _pendingStart;

    public int Next(int start)
    {
        if (_pendingStart.HasValue)
        {
            Current = _pendingStart.Value;
            HasValue = true;
            _pendingStart = null;
            return Current;
        }

        if (!HasValue)
        {
            Current = start;
            HasValue = true;
            return Current;
        }

        Current += 1;
        return Current;
    }

    public void Reset()
    {
        HasValue = false;
        Current = 0;
    }

    public void SetCurrent(int value)
    {
        Current = value;
        HasValue = true;
    }

    public void ForceStart(int start)
    {
        _pendingStart = start;
        HasValue = false;
    }
}

internal sealed class NumberingCounterState
{
    private readonly Dictionary<int, NumberingCounter> _counters = new();

    public NumberingCounter GetCounter(int level)
    {
        if (!_counters.TryGetValue(level, out var counter))
        {
            counter = new NumberingCounter();
            _counters[level] = counter;
        }

        return counter;
    }

    public NumberingCounter? TryGetCounter(int level) => _counters.TryGetValue(level, out var counter) ? counter : null;

    public void ResetLevelsAbove(int level)
    {
        foreach (var key in _counters.Keys.ToList())
        {
            if (key > level)
            {
                _counters[key].Reset();
            }
        }
    }
}

internal static class NumberingDiagnostics
{
    public static Action<string>? Sink { get; set; }

    public static void Write(string message)
    {
        Sink?.Invoke($"[Numbering] {message}");
    }
}

internal static class NumberFormatConverter
{
    public static string Format(int value, NumberFormatValues format)
    {
        if (format == NumberFormatValues.Decimal)
            return value.ToString();
        if (format == NumberFormatValues.DecimalZero)
            return value.ToString("00");
        if (format == NumberFormatValues.LowerLetter)
            return ToAlpha(value).ToLowerInvariant();
        if (format == NumberFormatValues.UpperLetter)
            return ToAlpha(value).ToUpperInvariant();
        if (format == NumberFormatValues.LowerRoman)
            return ToRoman(value).ToLowerInvariant();
        if (format == NumberFormatValues.UpperRoman)
            return ToRoman(value).ToUpperInvariant();

        return value.ToString();
    }

    private static string ToAlpha(int value)
    {
        if (value <= 0)
            return value.ToString();

        var chars = new Stack<char>();
        var current = value;
        while (current > 0)
        {
            current--;
            chars.Push((char)('a' + (current % 26)));
            current /= 26;
        }
        return new string(chars.ToArray());
    }

    private static string ToRoman(int value)
    {
        if (value <= 0)
            return value.ToString();

        (int, string)[] map =
        {
            (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
            (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
            (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")
        };

        var remaining = value;
        var result = string.Empty;

        foreach (var (number, numeral) in map)
        {
            while (remaining >= number)
            {
                result += numeral;
                remaining -= number;
            }
        }

        return result;
    }
}
