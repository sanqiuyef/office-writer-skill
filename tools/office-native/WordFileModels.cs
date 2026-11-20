using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace OfficeNative;

/// <summary>
/// JSON helpers shared by the Word file-level commands.  The shapes mirror
/// scripts/office_engine/models.py so the CLI JSON contract stays unchanged.
/// </summary>
internal static class WordJson
{
    public static bool TryGet(JsonElement parent, string name, out JsonElement value)
    {
        if (parent.ValueKind == JsonValueKind.Object && parent.TryGetProperty(name, out value))
            return true;
        value = default;
        return false;
    }

    public static string? OptString(JsonElement parent, string name)
    {
        if (!TryGet(parent, name, out var value) || value.ValueKind == JsonValueKind.Null)
            return null;
        return value.ValueKind switch
        {
            JsonValueKind.String => value.GetString(),
            JsonValueKind.True => "True",
            JsonValueKind.False => "False",
            _ => value.ToString(),
        };
    }

    public static bool? OptBool(JsonElement parent, string name)
    {
        if (!TryGet(parent, name, out var value) || value.ValueKind == JsonValueKind.Null)
            return null;
        return value.ValueKind switch
        {
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Number when value.TryGetDouble(out var number) => number != 0,
            JsonValueKind.String => ParseBool(value.GetString()),
            _ => null,
        };
    }

    private static bool? ParseBool(string? text)
    {
        if (text is null)
            return null;
        return text.Trim().ToLowerInvariant() switch
        {
            "true" or "1" or "yes" or "on" => true,
            "false" or "0" or "no" or "off" or "" => false,
            _ => null,
        };
    }

    public static int? OptInt(JsonElement parent, string name)
    {
        if (!TryGet(parent, name, out var value) || value.ValueKind == JsonValueKind.Null)
            return null;
        if (value.ValueKind == JsonValueKind.Number)
        {
            if (value.TryGetInt32(out var number))
                return number;
            if (value.TryGetDouble(out var dbl))
                return (int)Math.Round(dbl);
        }
        if (value.ValueKind == JsonValueKind.String
            && int.TryParse(value.GetString(), NumberStyles.Integer, CultureInfo.InvariantCulture,
                out var parsed))
            return parsed;
        return null;
    }

    public static double? OptDouble(JsonElement parent, string name)
    {
        if (!TryGet(parent, name, out var value) || value.ValueKind == JsonValueKind.Null)
            return null;
        if (value.ValueKind == JsonValueKind.Number && value.TryGetDouble(out var number))
            return number;
        if (value.ValueKind == JsonValueKind.String
            && double.TryParse(value.GetString(), NumberStyles.Float, CultureInfo.InvariantCulture,
                out var parsed))
            return parsed;
        return null;
    }

    public static double? OptSize(JsonElement parent, string name)
    {
        if (!TryGet(parent, name, out var value) || value.ValueKind == JsonValueKind.Null)
            return null;
        return WordFileOoxml.CnSizeToPt(value);
    }

    public static string? OptColor(JsonElement parent, string name)
    {
        var raw = OptString(parent, name);
        if (raw is null)
            return null;
        var text = raw.Trim();
        if (text.StartsWith('#'))
            text = text[1..];
        if (text.Length == 0)
            return null;
        if (!Regex.IsMatch(text, "^[0-9A-Fa-f]{6}$"))
            throw new NativeUsageException($"颜色必须是 6 位十六进制（如 FF0000）: {raw}");
        return text.ToUpperInvariant();
    }

    public static List<string> StringList(JsonElement parent, string name)
    {
        var result = new List<string>();
        if (!TryGet(parent, name, out var value) || value.ValueKind == JsonValueKind.Null)
            return result;
        if (value.ValueKind != JsonValueKind.Array)
            return result;
        foreach (var item in value.EnumerateArray())
            result.Add(CellString(item));
        return result;
    }

    public static List<List<string>> StringRows(JsonElement parent, string name)
    {
        var rows = new List<List<string>>();
        if (!TryGet(parent, name, out var value) || value.ValueKind != JsonValueKind.Array)
            return rows;
        foreach (var row in value.EnumerateArray())
        {
            var cells = new List<string>();
            if (row.ValueKind == JsonValueKind.Array)
            {
                foreach (var cell in row.EnumerateArray())
                    cells.Add(CellString(cell));
            }
            rows.Add(cells);
        }
        return rows;
    }

    public static List<double>? DoubleList(JsonElement parent, string name)
    {
        if (!TryGet(parent, name, out var value) || value.ValueKind == JsonValueKind.Null)
            return null;
        if (value.ValueKind != JsonValueKind.Array)
            return null;
        var result = new List<double>();
        foreach (var item in value.EnumerateArray())
        {
            if (item.ValueKind == JsonValueKind.Number && item.TryGetDouble(out var number))
                result.Add(number);
            else if (item.ValueKind == JsonValueKind.String
                     && double.TryParse(item.GetString(), NumberStyles.Float,
                         CultureInfo.InvariantCulture, out var parsed))
                result.Add(parsed);
        }
        return result;
    }

    public static IEnumerable<JsonElement> ArrayItems(JsonElement parent, string name)
    {
        if (!TryGet(parent, name, out var value) || value.ValueKind == JsonValueKind.Null)
            yield break;
        if (value.ValueKind != JsonValueKind.Array)
            throw new NativeUsageException($"{name} 必须是数组");
        foreach (var item in value.EnumerateArray())
            yield return item;
    }

    /// <summary>Python str() equivalent for single JSON scalars.</summary>
    public static string CellString(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.Null => "",
        JsonValueKind.String => value.GetString() ?? "",
        JsonValueKind.True => "True",
        JsonValueKind.False => "False",
        _ => value.GetRawText(),
    };
}

internal sealed class FontConfig
{
    public string? Name { get; init; }
    public double? SizePt { get; init; }
    public bool? Bold { get; init; }
    public bool? Italic { get; init; }
    public string? Color { get; init; }

    public static FontConfig? ParseOptional(JsonElement parent, string name)
    {
        if (!WordJson.TryGet(parent, name, out var value) || value.ValueKind == JsonValueKind.Null)
            return null;
        if (value.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException($"{name} 必须是对象");
        return Parse(value);
    }

    public static FontConfig Parse(JsonElement value) => new()
    {
        Name = WordJson.OptString(value, "name"),
        SizePt = WordJson.OptSize(value, "size"),
        Bold = WordJson.OptBool(value, "bold"),
        Italic = WordJson.OptBool(value, "italic"),
        Color = WordJson.OptColor(value, "color"),
    };
}

internal sealed class ParagraphStyle
{
    public string? FontName { get; init; }
    public double? SizePt { get; init; }
    public bool? Bold { get; init; }
    public bool? Italic { get; init; }
    public string? Color { get; init; }
    public string? Alignment { get; init; }
    public double? LineSpacing { get; init; }
    public int? FirstLineIndentPt { get; init; }
    public int? SpaceBeforePt { get; init; }
    public int? SpaceAfterPt { get; init; }

    public static ParagraphStyle? ParseOptional(JsonElement parent, string name)
    {
        if (!WordJson.TryGet(parent, name, out var value) || value.ValueKind == JsonValueKind.Null)
            return null;
        if (value.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException($"{name} 必须是对象");
        return Parse(value);
    }

    public static ParagraphStyle Parse(JsonElement value) => new()
    {
        FontName = WordJson.OptString(value, "font_name"),
        SizePt = WordJson.OptSize(value, "font_size"),
        Bold = WordJson.OptBool(value, "bold"),
        Italic = WordJson.OptBool(value, "italic"),
        Color = WordJson.OptColor(value, "color"),
        Alignment = WordJson.OptString(value, "alignment"),
        LineSpacing = WordJson.OptDouble(value, "line_spacing"),
        FirstLineIndentPt = WordJson.OptInt(value, "first_line_indent"),
        SpaceBeforePt = WordJson.OptInt(value, "space_before"),
        SpaceAfterPt = WordJson.OptInt(value, "space_after"),
    };
}

internal sealed class BorderSpec
{
    public string Style { get; init; } = "thin";
    public string? Color { get; init; }

    public static BorderSpec? ParseOptional(JsonElement parent, string name)
    {
        if (!WordJson.TryGet(parent, name, out var value) || value.ValueKind != JsonValueKind.Object)
            return null;
        return new BorderSpec
        {
            Style = WordJson.OptString(value, "style") ?? "thin",
            Color = WordJson.OptColor(value, "color"),
        };
    }
}

internal sealed class CellStyle
{
    public bool? Bold { get; init; }
    public bool? Italic { get; init; }
    public string? FontName { get; init; }
    public double? SizePt { get; init; }
    public string? FontColor { get; init; }
    public string? BgColor { get; init; }
    public string? NumberFormat { get; init; }
    public string? Alignment { get; init; }
    public string? VerticalAlignment { get; init; }
    public bool? WrapText { get; init; }
    public BorderSpec? Border { get; init; }

    public static CellStyle? ParseOptional(JsonElement parent, string name)
    {
        if (!WordJson.TryGet(parent, name, out var value) || value.ValueKind != JsonValueKind.Object)
            return null;
        return new CellStyle
        {
            Bold = WordJson.OptBool(value, "bold"),
            Italic = WordJson.OptBool(value, "italic"),
            FontName = WordJson.OptString(value, "font_name"),
            SizePt = WordJson.OptSize(value, "font_size"),
            FontColor = WordJson.OptColor(value, "font_color"),
            BgColor = WordJson.OptColor(value, "bg_color"),
            NumberFormat = WordJson.OptString(value, "number_format"),
            Alignment = WordJson.OptString(value, "alignment"),
            VerticalAlignment = WordJson.OptString(value, "vertical_alignment"),
            WrapText = WordJson.OptBool(value, "wrap_text"),
            Border = BorderSpec.ParseOptional(value, "border"),
        };
    }
}

internal sealed class PageSetup
{
    public string? Orientation { get; init; }
    public string? PaperSize { get; init; } = "A4";
    public double? MarginTop { get; init; }
    public double? MarginBottom { get; init; }
    public double? MarginLeft { get; init; }
    public double? MarginRight { get; init; }

    public static PageSetup? ParseOptional(JsonElement parent, string name)
    {
        if (!WordJson.TryGet(parent, name, out var value) || value.ValueKind == JsonValueKind.Null)
            return null;
        if (value.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException($"{name} 必须是对象");
        var orientation = WordJson.OptString(value, "orientation");
        if (orientation is not null && orientation is not ("portrait" or "landscape"))
            throw new NativeUsageException("orientation 必须是 portrait/landscape");
        var paper = WordJson.TryGet(value, "paper_size", out var paperElement)
            ? (paperElement.ValueKind == JsonValueKind.Null
                ? null
                : WordJson.OptString(value, "paper_size"))
            : "A4";
        if (paper is not null && paper is not ("A4" or "Letter" or "A3" or "B5"))
            throw new NativeUsageException("paper_size 必须是 A4/Letter/A3/B5");
        return new PageSetup
        {
            Orientation = orientation,
            PaperSize = paper,
            MarginTop = WordJson.OptDouble(value, "margin_top"),
            MarginBottom = WordJson.OptDouble(value, "margin_bottom"),
            MarginLeft = WordJson.OptDouble(value, "margin_left"),
            MarginRight = WordJson.OptDouble(value, "margin_right"),
        };
    }
}

internal sealed class HeaderFooter
{
    public string? HeaderText { get; init; }
    public string? FooterText { get; init; }

    public static HeaderFooter? ParseOptional(JsonElement parent, string name)
    {
        if (!WordJson.TryGet(parent, name, out var value) || value.ValueKind == JsonValueKind.Null)
            return null;
        if (value.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException($"{name} 必须是对象");
        return new HeaderFooter
        {
            HeaderText = WordJson.OptString(value, "header_text"),
            FooterText = WordJson.OptString(value, "footer_text"),
        };
    }
}

internal sealed class CoverConfig
{
    public string? Title { get; init; }
    public string? Subtitle { get; init; }
    public string? Author { get; init; }
    public string? Date { get; init; }
    public string? Organization { get; init; }
    public List<string> ExtraLines { get; init; } = [];
    public string? RecipeId { get; init; }
    public string? DocType { get; init; }

    public static CoverConfig? ParseOptional(JsonElement parent, string name)
    {
        if (!WordJson.TryGet(parent, name, out var value) || value.ValueKind == JsonValueKind.Null)
            return null;
        if (value.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException($"{name} 必须是对象");
        return new CoverConfig
        {
            Title = WordJson.OptString(value, "title"),
            Subtitle = WordJson.OptString(value, "subtitle"),
            Author = WordJson.OptString(value, "author"),
            Date = WordJson.OptString(value, "date"),
            Organization = WordJson.OptString(value, "organization"),
            ExtraLines = WordJson.StringList(value, "extra_lines"),
            RecipeId = WordJson.OptString(value, "recipe_id"),
            DocType = WordJson.OptString(value, "doc_type"),
        };
    }
}

internal sealed class MergeRange
{
    public string Range { get; init; } = "";

    public static MergeRange Parse(JsonElement value)
    {
        if (value.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException("merged_cells 的每一项必须是对象");
        var range = WordJson.OptString(value, "range");
        if (string.IsNullOrWhiteSpace(range))
            throw new NativeUsageException("merged_cells 的每一项必须提供 range");
        return new MergeRange { Range = range };
    }
}

internal sealed class TableDef
{
    public List<string> Headers { get; init; } = [];
    public List<List<string>> Rows { get; init; } = [];
    public List<double>? ColumnWidths { get; init; }
    public CellStyle? HeaderStyle { get; init; }
    public List<MergeRange> MergedCells { get; init; } = [];

    public static TableDef Parse(JsonElement value) => new()
    {
        Headers = WordJson.StringList(value, "headers"),
        Rows = WordJson.StringRows(value, "rows"),
        ColumnWidths = WordJson.DoubleList(value, "column_widths"),
        HeaderStyle = CellStyle.ParseOptional(value, "header_style"),
        MergedCells = WordJson.ArrayItems(value, "merged_cells").Select(MergeRange.Parse).ToList(),
    };
}

internal sealed class ImageDef
{
    public string Path { get; init; } = "";
    public double? Width { get; init; }
    public double? Height { get; init; }

    public static ImageDef? ParseOptional(JsonElement parent, string name)
    {
        if (!WordJson.TryGet(parent, name, out var value) || value.ValueKind != JsonValueKind.Object)
            return null;
        return new ImageDef
        {
            Path = WordJson.OptString(value, "path") ?? "",
            Width = WordJson.OptDouble(value, "width"),
            Height = WordJson.OptDouble(value, "height"),
        };
    }
}

internal sealed class HyperlinkDef
{
    public string Text { get; init; } = "";
    public string Url { get; init; } = "";
    public string? Tooltip { get; init; }

    public static HyperlinkDef? ParseOptional(JsonElement parent, string name)
    {
        if (!WordJson.TryGet(parent, name, out var value) || value.ValueKind != JsonValueKind.Object)
            return null;
        return new HyperlinkDef
        {
            Text = WordJson.OptString(value, "text") ?? "",
            Url = WordJson.OptString(value, "url") ?? "",
            Tooltip = WordJson.OptString(value, "tooltip"),
        };
    }
}

internal sealed class FieldDef
{
    private static readonly HashSet<string> AllowedTypes =
    [
        "PAGE", "NUMPAGES", "SECTIONPAGES", "DATE", "TIME",
        "FILENAME", "AUTHOR", "TITLE", "CUSTOM",
    ];

    public string FieldType { get; init; } = "PAGE";
    public string? Code { get; init; }
    public string? FallbackText { get; init; }
    public string? NumberFormat { get; init; }

    public static FieldDef? ParseOptional(JsonElement parent, string name)
    {
        if (!WordJson.TryGet(parent, name, out var value) || value.ValueKind != JsonValueKind.Object)
            return null;
        var fieldType = WordJson.OptString(value, "field_type") ?? "PAGE";
        if (!AllowedTypes.Contains(fieldType))
            throw new NativeUsageException(
                "field_type 必须是 PAGE/NUMPAGES/SECTIONPAGES/DATE/TIME/FILENAME/AUTHOR/TITLE/CUSTOM");
        var code = WordJson.OptString(value, "code");
        if (fieldType == "CUSTOM" && string.IsNullOrWhiteSpace(code))
            throw new NativeUsageException("field_type=CUSTOM 时必须提供非空 code");
        return new FieldDef
        {
            FieldType = fieldType,
            Code = code,
            FallbackText = WordJson.OptString(value, "fallback_text"),
            NumberFormat = WordJson.OptString(value, "number_format"),
        };
    }
}

internal sealed class SectionItem
{
    public string Type { get; init; } = "";
    public string? Text { get; init; }
    public int? Level { get; init; }
    public string? Language { get; init; }
    public ParagraphStyle? Style { get; init; }
    public List<string> Items { get; init; } = [];
    public TableDef? Table { get; init; }
    public ImageDef? Image { get; init; }
    public string? TocLevels { get; init; }
    public HyperlinkDef? Hyperlink { get; init; }
    public FieldDef? Field { get; init; }

    private static readonly HashSet<string> AllowedTypes =
    [
        "heading", "paragraph", "table", "image", "page_break", "bullet_list",
        "numbered_list", "code", "toc", "section_break", "hyperlink", "field",
    ];

    public static SectionItem Parse(JsonElement value)
    {
        if (value.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException("sections 的每一项必须是对象");
        var type = WordJson.OptString(value, "type");
        if (string.IsNullOrWhiteSpace(type) || !AllowedTypes.Contains(type))
            throw new NativeUsageException(
                "type 必须是 heading/paragraph/table/image/page_break/bullet_list/"
                + "numbered_list/code/toc/section_break/hyperlink/field");
        var hyperlink = HyperlinkDef.ParseOptional(value, "hyperlink");
        var field = FieldDef.ParseOptional(value, "field");
        if (type == "hyperlink" && hyperlink is null)
            throw new NativeUsageException("type=hyperlink 时必须提供 hyperlink 定义");
        if (type == "field" && field is null)
            throw new NativeUsageException("type=field 时必须提供 field 定义");
        return new SectionItem
        {
            Type = type,
            Text = WordJson.OptString(value, "text"),
            Level = WordJson.OptInt(value, "level"),
            Language = WordJson.OptString(value, "language"),
            Style = ParagraphStyle.ParseOptional(value, "style"),
            Items = WordJson.StringList(value, "items"),
            Table = WordJson.TryGet(value, "table", out var tableElement)
                    && tableElement.ValueKind == JsonValueKind.Object
                ? TableDef.Parse(tableElement)
                : null,
            Image = ImageDef.ParseOptional(value, "image"),
            TocLevels = WordJson.OptString(value, "toc_levels"),
            Hyperlink = hyperlink,
            Field = field,
        };
    }
}

internal sealed class WordProperties
{
    private static readonly HashSet<string> KnownFields =
    [
        "title", "subject", "author", "keywords", "comments",
        "category", "last_modified_by", "created", "modified",
    ];

    public string? Title { get; init; }
    public string? Subject { get; init; }
    public string? Author { get; init; }
    public string? Keywords { get; init; }
    public string? Comments { get; init; }
    public string? Category { get; init; }
    public string? LastModifiedBy { get; init; }
    public DateTime? Created { get; init; }
    public DateTime? Modified { get; init; }

    public static WordProperties Parse(JsonElement value, bool rejectUnknown)
    {
        if (value.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException("properties 必须是对象");
        if (rejectUnknown)
        {
            var unsupported = value.EnumerateObject()
                .Select(item => item.Name)
                .Where(name => !KnownFields.Contains(name))
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToList();
            if (unsupported.Count > 0)
                throw new NativeUsageException(
                    "不支持的自定义属性字段：" + string.Join("、", unsupported)
                    + "（仅支持核心属性）");
        }
        return new WordProperties
        {
            Title = WordJson.OptString(value, "title"),
            Subject = WordJson.OptString(value, "subject"),
            Author = WordJson.OptString(value, "author"),
            Keywords = WordJson.OptString(value, "keywords"),
            Comments = WordJson.OptString(value, "comments"),
            Category = WordJson.OptString(value, "category"),
            LastModifiedBy = WordJson.OptString(value, "last_modified_by"),
            Created = ParseIso8601(WordJson.OptString(value, "created"), "created"),
            Modified = ParseIso8601(WordJson.OptString(value, "modified"), "modified"),
        };
    }

    public static DateTime? ParseIso8601(string? raw, string fieldName)
    {
        if (raw is null)
            return null;
        var text = raw.Trim();
        var aware = Regex.IsMatch(text, @"(?:Z|[+-]\d{2}:?\d{2})$", RegexOptions.IgnoreCase);
        if (aware)
        {
            if (!DateTimeOffset.TryParse(text, CultureInfo.InvariantCulture,
                    DateTimeStyles.RoundtripKind, out var parsed))
                throw new NativeUsageException($"{fieldName} 不是合法的 ISO 8601 时间: '{raw}'");
            return DateTime.SpecifyKind(parsed.UtcDateTime, DateTimeKind.Utc);
        }
        if (!DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out var naive))
            throw new NativeUsageException($"{fieldName} 不是合法的 ISO 8601 时间: '{raw}'");
        return DateTime.SpecifyKind(naive, DateTimeKind.Unspecified);
    }

    public List<string> NonNullFields()
    {
        var fields = new List<string>();
        if (Title is not null) fields.Add("title");
        if (Subject is not null) fields.Add("subject");
        if (Author is not null) fields.Add("author");
        if (Keywords is not null) fields.Add("keywords");
        if (Comments is not null) fields.Add("comments");
        if (Category is not null) fields.Add("category");
        if (LastModifiedBy is not null) fields.Add("last_modified_by");
        if (Created is not null) fields.Add("created");
        if (Modified is not null) fields.Add("modified");
        return fields;
    }
}

internal sealed class WordContent
{
    public List<SectionItem> Sections { get; init; } = [];
    public WordProperties? Properties { get; init; }

    public static WordContent Parse(JsonElement value)
    {
        if (value.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException("content 必须是对象");
        if (!WordJson.TryGet(value, "sections", out var sections)
            || sections.ValueKind != JsonValueKind.Array)
            throw new NativeUsageException("content.sections 必须是数组");
        return new WordContent
        {
            Sections = sections.EnumerateArray().Select(SectionItem.Parse).ToList(),
            Properties = WordJson.TryGet(value, "properties", out var properties)
                         && properties.ValueKind == JsonValueKind.Object
                ? WordProperties.Parse(properties, rejectUnknown: false)
                : null,
        };
    }
}

internal sealed class WordCommentDef
{
    public string Text { get; init; } = "";
    public string Author { get; init; } = "office-writer";
    public string Initials { get; init; } = "OW";
    public string? AnchorText { get; init; }
    public int? ParagraphIndex { get; init; }
    public bool MatchCase { get; init; }

    public static List<WordCommentDef> ParseList(JsonElement parent, string name)
    {
        var result = new List<WordCommentDef>();
        var index = 0;
        foreach (var item in WordJson.ArrayItems(parent, name))
        {
            if (item.ValueKind != JsonValueKind.Object)
                throw new NativeUsageException($"{name}[{index}] 必须是对象");
            if (!WordJson.TryGet(item, "text", out var textElement)
                || textElement.ValueKind == JsonValueKind.Null)
                throw new NativeUsageException($"{name}[{index}] 缺少 text");
            var anchor = WordJson.OptString(item, "anchor_text");
            var paragraphIndex = WordJson.OptInt(item, "paragraph_index");
            if (paragraphIndex is < 1)
                throw new NativeUsageException(
                    $"{name}[{index}].paragraph_index 必须 >= 1");
            if (string.IsNullOrEmpty(anchor) && paragraphIndex is null)
                throw new NativeUsageException(
                    "WordCommentDef 必须提供 anchor_text 或 paragraph_index");
            result.Add(new WordCommentDef
            {
                Text = WordJson.CellString(textElement),
                Author = WordJson.OptString(item, "author") ?? "office-writer",
                Initials = WordJson.OptString(item, "initials") ?? "OW",
                AnchorText = string.IsNullOrEmpty(anchor) ? null : anchor,
                ParagraphIndex = paragraphIndex,
                MatchCase = WordJson.OptBool(item, "match_case") ?? false,
            });
            index++;
        }
        return result;
    }
}
