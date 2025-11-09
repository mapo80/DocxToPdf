using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;

namespace DocxToPdf.Sdk.Docx.Styles;

internal sealed class ColorSchemeMapper
{
    private readonly Dictionary<ThemeColorValues, ThemeColorValues> _map =
        new()
        {
            [ThemeColorValues.Background1] = ThemeColorValues.Light1,
            [ThemeColorValues.Text1] = ThemeColorValues.Dark1,
            [ThemeColorValues.Background2] = ThemeColorValues.Light2,
            [ThemeColorValues.Text2] = ThemeColorValues.Dark2,
            [ThemeColorValues.Accent1] = ThemeColorValues.Accent1,
            [ThemeColorValues.Accent2] = ThemeColorValues.Accent2,
            [ThemeColorValues.Accent3] = ThemeColorValues.Accent3,
            [ThemeColorValues.Accent4] = ThemeColorValues.Accent4,
            [ThemeColorValues.Accent5] = ThemeColorValues.Accent5,
            [ThemeColorValues.Accent6] = ThemeColorValues.Accent6,
            [ThemeColorValues.Hyperlink] = ThemeColorValues.Hyperlink,
            [ThemeColorValues.FollowedHyperlink] = ThemeColorValues.FollowedHyperlink
        };

    public static ColorSchemeMapper Load(ColorSchemeMapping? mapping)
    {
        var mapper = new ColorSchemeMapper();
        if (mapping == null)
            return mapper;

        mapper.Override(ThemeColorValues.Background1, mapping.Background1?.Value);
        mapper.Override(ThemeColorValues.Text1, mapping.Text1?.Value);
        mapper.Override(ThemeColorValues.Background2, mapping.Background2?.Value);
        mapper.Override(ThemeColorValues.Text2, mapping.Text2?.Value);
        mapper.Override(ThemeColorValues.Accent1, mapping.Accent1?.Value);
        mapper.Override(ThemeColorValues.Accent2, mapping.Accent2?.Value);
        mapper.Override(ThemeColorValues.Accent3, mapping.Accent3?.Value);
        mapper.Override(ThemeColorValues.Accent4, mapping.Accent4?.Value);
        mapper.Override(ThemeColorValues.Accent5, mapping.Accent5?.Value);
        mapper.Override(ThemeColorValues.Accent6, mapping.Accent6?.Value);
        mapper.Override(ThemeColorValues.Hyperlink, mapping.Hyperlink?.Value);
        mapper.Override(ThemeColorValues.FollowedHyperlink, mapping.FollowedHyperlink?.Value);

        return mapper;
    }

    private void Override(ThemeColorValues key, ColorSchemeIndexValues? mapped)
    {
        if (mapped.HasValue)
        {
            _map[key] = MapIndex(mapped.Value);
        }
    }

    private static ThemeColorValues MapIndex(ColorSchemeIndexValues index)
    {
        if (index == ColorSchemeIndexValues.Light1) return ThemeColorValues.Light1;
        if (index == ColorSchemeIndexValues.Dark1) return ThemeColorValues.Dark1;
        if (index == ColorSchemeIndexValues.Light2) return ThemeColorValues.Light2;
        if (index == ColorSchemeIndexValues.Dark2) return ThemeColorValues.Dark2;
        if (index == ColorSchemeIndexValues.Accent1) return ThemeColorValues.Accent1;
        if (index == ColorSchemeIndexValues.Accent2) return ThemeColorValues.Accent2;
        if (index == ColorSchemeIndexValues.Accent3) return ThemeColorValues.Accent3;
        if (index == ColorSchemeIndexValues.Accent4) return ThemeColorValues.Accent4;
        if (index == ColorSchemeIndexValues.Accent5) return ThemeColorValues.Accent5;
        if (index == ColorSchemeIndexValues.Accent6) return ThemeColorValues.Accent6;
        if (index == ColorSchemeIndexValues.Hyperlink) return ThemeColorValues.Hyperlink;
        if (index == ColorSchemeIndexValues.FollowedHyperlink) return ThemeColorValues.FollowedHyperlink;
        return ThemeColorValues.Dark1;
    }

    public ThemeColorValues Resolve(ThemeColorValues requested) =>
        _map.TryGetValue(requested, out var mapped) ? mapped : requested;
}
