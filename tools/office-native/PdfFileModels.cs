using System.Globalization;
using System.Text.Json;

namespace OfficeNative;

// JSON shapes mirror scripts/office_engine/models.py (PdfContent & friends) so the
// CLI contract of create_pdf_document stays unchanged.

internal sealed class PdfContent
{
    public List<PdfSection> Sections { get; } = [];

    public static PdfContent Parse(JsonElement parent)
    {
        if (!PdfJson.TryGet(parent, "content", out var element)
            || element.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException("缺少参数 content（JSON 对象）");
        if (!PdfJson.TryGet(element, "sections", out var sections)
            || sections.ValueKind != JsonValueKind.Array)
            throw new NativeUsageException("缺少参数 content.sections（JSON 数组）");
        var content = new PdfContent();
        var index = 0;
        foreach (var item in sections.EnumerateArray())
        {
            content.Sections.Add(PdfSection.Parse(item, $"content.sections[{index}]"));
            index++;
        }
        return content;
    }
}

internal sealed class PdfSection
{
    private static readonly HashSet<string> Types =
        ["heading", "paragraph", "table", "image", "page_break", "bullet_list", "numbered_list"];

    public required string Type { get; init; }
    public string? Text { get; init; }
    public int? Level { get; init; }
    public PdfParagraphStyle? Style { get; init; }
    public List<string>? Items { get; init; }
    public PdfTable? Table { get; init; }
    public PdfImage? Image { get; init; }

    public static PdfSection Parse(JsonElement element, string label)
    {
        if (element.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException($"{label} 必须是 JSON 对象");
        var type = PdfJson.RequireString(element, "type", $"{label}.type");
        if (!Types.Contains(type))
            throw new NativeUsageException(
                $"{label}.type 必须是 {string.Join("/", Types)} 之一，收到 {type}");
        return new PdfSection
        {
            Type = type,
            Text = PdfJson.OptString(element, "text"),
            Level = PdfJson.OptInt(element, "level"),
            Style = PdfParagraphStyle.ParseOptional(element, "style", label),
            Items = PdfJson.StringList(element, "items", label),
            Table = PdfTable.ParseOptional(element, "table", label),
            Image = PdfImage.ParseOptional(element, "image", label),
        };
    }
}

internal sealed class PdfParagraphStyle
{
    public string? FontName { get; init; }
    public double? FontSizePt { get; init; }
    public bool? Bold { get; init; }
    public bool? Italic { get; init; }
    public string? Color { get; init; }
    public string? Alignment { get; init; }
    public double? LineSpacing { get; init; }
    public double? FirstLineIndentPt { get; init; }
    public double? SpaceBeforePt { get; init; }
    public double? SpaceAfterPt { get; init; }

    private static readonly HashSet<string> Alignments = ["left", "center", "right", "justify"];

    public static PdfParagraphStyle? ParseOptional(
        JsonElement parent, string name, string label)
    {
        if (!PdfJson.TryGet(parent, name, out var element)
            || element.ValueKind == JsonValueKind.Null)
            return null;
        if (element.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException($"{label}.{name} 必须是 JSON 对象");
        var alignment = PdfJson.OptString(element, "alignment");
        if (alignment is not null && !Alignments.Contains(alignment))
            throw new NativeUsageException(
                $"{label}.{name}.alignment 必须是 left/center/right/justify 之一");
        return new PdfParagraphStyle
        {
            FontName = PdfJson.OptString(element, "font_name"),
            FontSizePt = PdfJson.OptSize(element, "font_size", $"{label}.{name}.font_size"),
            Bold = PdfJson.OptBool(element, "bold"),
            Italic = PdfJson.OptBool(element, "italic"),
            Color = PdfJson.OptString(element, "color"),
            Alignment = alignment,
            LineSpacing = PdfJson.OptDouble(element, "line_spacing"),
            FirstLineIndentPt = PdfJson.OptDouble(element, "first_line_indent"),
            SpaceBeforePt = PdfJson.OptDouble(element, "space_before"),
            SpaceAfterPt = PdfJson.OptDouble(element, "space_after"),
        };
    }
}

internal sealed class PdfTable
{
    public List<string> Headers { get; init; } = [];
    public List<List<string>> Rows { get; init; } = [];
    public List<double>? ColumnWidths { get; init; }

    public static PdfTable? ParseOptional(JsonElement parent, string name, string label)
    {
        if (!PdfJson.TryGet(parent, name, out var element)
            || element.ValueKind == JsonValueKind.Null)
            return null;
        if (element.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException($"{label}.{name} 必须是 JSON 对象");

        var headers = PdfJson.StringList(element, "headers", $"{label}.{name}");
        var rows = new List<List<string>>();
        if (PdfJson.TryGet(element, "rows", out var rowsElement)
            && rowsElement.ValueKind != JsonValueKind.Null)
        {
            if (rowsElement.ValueKind != JsonValueKind.Array)
                throw new NativeUsageException($"{label}.{name}.rows 必须是 JSON 数组");
            var rowIndex = 0;
            foreach (var row in rowsElement.EnumerateArray())
            {
                if (row.ValueKind != JsonValueKind.Array)
                    throw new NativeUsageException(
                        $"{label}.{name}.rows[{rowIndex}] 必须是 JSON 数组");
                var cells = new List<string>();
                foreach (var cell in row.EnumerateArray())
                {
                    cells.Add(cell.ValueKind switch
                    {
                        JsonValueKind.Null => "",
                        JsonValueKind.String => cell.GetString() ?? "",
                        _ => cell.ToString(),
                    });
                }
                rows.Add(cells);
                rowIndex++;
            }
        }

        List<double>? widths = null;
        if (PdfJson.TryGet(element, "column_widths", out var widthsElement)
            && widthsElement.ValueKind != JsonValueKind.Null)
        {
            if (widthsElement.ValueKind != JsonValueKind.Array)
                throw new NativeUsageException($"{label}.{name}.column_widths 必须是 JSON 数组");
            widths = [];
            foreach (var width in widthsElement.EnumerateArray())
            {
                if (width.ValueKind == JsonValueKind.Number && width.TryGetDouble(out var value))
                    widths.Add(value);
                else if (width.ValueKind == JsonValueKind.String
                         && double.TryParse(width.GetString(), NumberStyles.Float,
                             CultureInfo.InvariantCulture, out var parsed))
                    widths.Add(parsed);
                else
                    throw new NativeUsageException(
                        $"{label}.{name}.column_widths 必须是数字数组");
            }
        }

        return new PdfTable
        {
            Headers = headers,
            Rows = rows,
            ColumnWidths = widths,
        };
    }
}

internal sealed class PdfImage
{
    public required string Path { get; init; }
    public double? WidthInches { get; init; }
    public double? HeightInches { get; init; }

    public static PdfImage? ParseOptional(JsonElement parent, string name, string label)
    {
        if (!PdfJson.TryGet(parent, name, out var element)
            || element.ValueKind == JsonValueKind.Null)
            return null;
        if (element.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException($"{label}.{name} 必须是 JSON 对象");
        var path = PdfJson.OptString(element, "path");
        if (string.IsNullOrWhiteSpace(path))
            throw new NativeUsageException($"{label}.{name}.path 不能为空");
        return new PdfImage
        {
            Path = path,
            WidthInches = PdfJson.OptDouble(element, "width"),
            HeightInches = PdfJson.OptDouble(element, "height"),
        };
    }
}

internal sealed class PdfPageSetup
{
    public const double MmToPt = 72.0 / 25.4;

    private static readonly Dictionary<string, (double Width, double Height)> PaperSizes =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["A4"] = (210, 297),
            ["Letter"] = (215.9, 279.4),
            ["A3"] = (297, 420),
            ["B5"] = (176, 250),
        };

    public string PaperSize { get; init; } = "A4";
    public string Orientation { get; init; } = "portrait";
    public double MarginTopMm { get; init; } = 10;
    public double MarginBottomMm { get; init; } = 10;
    public double MarginLeftMm { get; init; } = 10;
    public double MarginRightMm { get; init; } = 10;

    public double WidthMm
    {
        get
        {
            var (width, height) = PaperSizes[PaperSize];
            return Orientation == "landscape" ? height : width;
        }
    }

    public double HeightMm
    {
        get
        {
            var (width, height) = PaperSizes[PaperSize];
            return Orientation == "landscape" ? width : height;
        }
    }

    public double ContentWidthPt =>
        (WidthMm - MarginLeftMm - MarginRightMm) * MmToPt;

    public static PdfPageSetup Parse(JsonElement parent)
    {
        if (!PdfJson.TryGet(parent, "page_setup", out var element)
            || element.ValueKind == JsonValueKind.Null)
            return new PdfPageSetup();
        if (element.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException("page_setup 必须是 JSON 对象");

        var paper = PdfJson.OptString(element, "paper_size") ?? "A4";
        if (!PaperSizes.ContainsKey(paper))
            throw new NativeUsageException("page_setup.paper_size 必须是 A4/Letter/A3/B5 之一");
        var orientation = PdfJson.OptString(element, "orientation") ?? "portrait";
        if (orientation is not ("portrait" or "landscape"))
            throw new NativeUsageException(
                "page_setup.orientation 必须是 portrait/landscape 之一");

        return new PdfPageSetup
        {
            PaperSize = paper,
            Orientation = orientation,
            MarginTopMm = PdfJson.OptDouble(element, "margin_top") ?? 10,
            MarginBottomMm = PdfJson.OptDouble(element, "margin_bottom") ?? 10,
            MarginLeftMm = PdfJson.OptDouble(element, "margin_left") ?? 10,
            MarginRightMm = PdfJson.OptDouble(element, "margin_right") ?? 10,
        };
    }
}

internal sealed class PdfFontConfig
{
    public string CnFont { get; init; } = "";
    public string? CnFontPath { get; init; }
    public string EnFont { get; init; } = "Helvetica";
    public double BodySizePt { get; init; } = 12;
    public double HeadingSizePt { get; init; } = 16;

    public static PdfFontConfig Parse(JsonElement parent)
    {
        if (!PdfJson.TryGet(parent, "font_config", out var element)
            || element.ValueKind == JsonValueKind.Null)
            return new PdfFontConfig();
        if (element.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException("font_config 必须是 JSON 对象");
        return new PdfFontConfig
        {
            CnFont = PdfJson.OptString(element, "cn_font") ?? "",
            CnFontPath = PdfJson.OptString(element, "cn_font_path"),
            EnFont = PdfJson.OptString(element, "en_font") ?? "Helvetica",
            BodySizePt = PdfJson.OptSize(element, "cn_font_size", "font_config.cn_font_size")
                         ?? 12,
            HeadingSizePt = PdfJson.OptSize(
                                element, "heading_font_size", "font_config.heading_font_size")
                            ?? 16,
        };
    }
}

internal static class PdfJson
{
    private static readonly Dictionary<string, double> CnFontSizes =
        new(StringComparer.Ordinal)
        {
            ["初号"] = 42.0,
            ["小初"] = 36.0,
            ["一号"] = 26.0,
            ["小一"] = 24.0,
            ["二号"] = 22.0,
            ["小二"] = 18.0,
            ["三号"] = 16.0,
            ["小三"] = 15.0,
            ["四号"] = 14.0,
            ["小四"] = 12.0,
            ["五号"] = 10.5,
            ["小五"] = 9.0,
            ["六号"] = 7.5,
            ["小六"] = 6.5,
            ["七号"] = 5.5,
            ["八号"] = 5.0,
        };

    public static bool TryGet(JsonElement parent, string name, out JsonElement value)
    {
        if (parent.ValueKind == JsonValueKind.Object
            && parent.TryGetProperty(name, out value))
            return true;
        value = default;
        return false;
    }

    public static string RequireString(JsonElement parent, string name, string label)
    {
        var value = OptString(parent, name);
        if (string.IsNullOrWhiteSpace(value))
            throw new NativeUsageException($"{label} 不能为空");
        return value;
    }

    public static string? OptString(JsonElement parent, string name)
    {
        if (!TryGet(parent, name, out var value)
            || value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;
        return value.ValueKind switch
        {
            JsonValueKind.String => value.GetString(),
            JsonValueKind.True => "True",
            JsonValueKind.False => "False",
            _ => value.ToString(),
        };
    }

    public static int? OptInt(JsonElement parent, string name)
    {
        if (!TryGet(parent, name, out var value)
            || value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;
        if (value.ValueKind == JsonValueKind.Number)
        {
            if (value.TryGetInt32(out var number))
                return number;
            if (value.TryGetDouble(out var dbl))
                return (int)Math.Round(dbl);
        }
        if (value.ValueKind == JsonValueKind.String
            && int.TryParse(value.GetString(), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out var parsed))
            return parsed;
        return null;
    }

    public static double? OptDouble(JsonElement parent, string name)
    {
        if (!TryGet(parent, name, out var value)
            || value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;
        if (value.ValueKind == JsonValueKind.Number && value.TryGetDouble(out var number))
            return number;
        if (value.ValueKind == JsonValueKind.String
            && double.TryParse(value.GetString(), NumberStyles.Float,
                CultureInfo.InvariantCulture, out var parsed))
            return parsed;
        return null;
    }

    public static bool? OptBool(JsonElement parent, string name)
    {
        if (!TryGet(parent, name, out var value)
            || value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;
        return value.ValueKind switch
        {
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Number when value.TryGetDouble(out var number) => number != 0,
            JsonValueKind.String => bool.TryParse(value.GetString(), out var parsed)
                ? parsed
                : null,
            _ => null,
        };
    }

    public static double? OptSize(JsonElement parent, string name, string label)
    {
        if (!TryGet(parent, name, out var value)
            || value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;
        if (value.ValueKind == JsonValueKind.Number && value.TryGetDouble(out var number))
            return number;
        var text = value.ValueKind == JsonValueKind.String
            ? value.GetString()?.Trim()
            : value.ToString().Trim();
        if (string.IsNullOrEmpty(text))
            return null;
        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture,
                out var parsed))
            return parsed;
        if (CnFontSizes.TryGetValue(text, out var pt))
            return pt;
        throw new NativeUsageException(
            $"{label} 未知的中文字号: {text}，可选值: {string.Join(", ", CnFontSizes.Keys)}");
    }

    public static List<string> StringList(JsonElement parent, string name, string label)
    {
        var result = new List<string>();
        if (!TryGet(parent, name, out var value)
            || value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return result;
        if (value.ValueKind != JsonValueKind.Array)
            throw new NativeUsageException($"{label}.{name} 必须是 JSON 数组");
        foreach (var item in value.EnumerateArray())
        {
            result.Add(item.ValueKind switch
            {
                JsonValueKind.String => item.GetString() ?? "",
                JsonValueKind.Null => "",
                _ => item.ToString(),
            });
        }
        return result;
    }

    public static int Clamp(int value, int min, int max) =>
        value < min ? min : value > max ? max : value;
}
