using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace OfficeNative;

// ── 参数模型（与 office_engine/models.py 的 Excel 子模型对齐）──────────

internal sealed class ExcelBorderSpec
{
    public string Style { get; init; } = "thin";
    public string? Color { get; init; }
}

internal sealed class ExcelCellStyle
{
    public bool? Bold { get; init; }
    public bool? Italic { get; init; }
    public string? FontName { get; init; }
    public double? FontSize { get; init; }
    public string? FontColor { get; init; }
    public string? BgColor { get; init; }
    public string? NumberFormat { get; init; }
    public string? Alignment { get; init; }
    public string? VerticalAlignment { get; init; }
    public bool? WrapText { get; init; }
    public ExcelBorderSpec? Border { get; init; }
}

internal sealed class ExcelColumnFormat
{
    public string? NumberFormat { get; init; }
    public string? FontName { get; init; }
    public double? FontSize { get; init; }
    public string? FontColor { get; init; }
    public string? BgColor { get; init; }
    public bool? Bold { get; init; }
    public bool? Italic { get; init; }
    public bool? WrapText { get; init; }
    public double? Width { get; init; }
    public string? Alignment { get; init; }
    public string? VerticalAlignment { get; init; }

    public ExcelCellStyle ToCellStyle() => new()
    {
        NumberFormat = NumberFormat,
        FontName = FontName,
        FontSize = FontSize,
        FontColor = FontColor,
        BgColor = BgColor,
        Bold = Bold,
        Italic = Italic,
        WrapText = WrapText,
        Alignment = Alignment,
        VerticalAlignment = VerticalAlignment,
    };
}

internal sealed class ExcelChartDef
{
    public string ChartType { get; init; } = "column";
    public string? Title { get; init; }
    public string DataRange { get; init; } = "";
    public string? CategoriesRange { get; init; }
    public bool TitlesFromData { get; init; } = true;
    public string Anchor { get; init; } = "E2";
    public double Width { get; init; } = 15;
    public double Height { get; init; } = 8;
    public string Legend { get; init; } = "right";
    public string? XAxisTitle { get; init; }
    public string? YAxisTitle { get; init; }
    public int? Style { get; init; }
}

internal sealed class ExcelConditionalFormatDef
{
    public string Range { get; init; } = "";
    public string Type { get; init; } = "cell_is";
    public string? Operator { get; init; }
    public List<string>? Formula { get; init; }
    public string? FillColor { get; init; }
    public string? FontColor { get; init; }
    public string? MinColor { get; init; }
    public string? MidColor { get; init; }
    public string? MaxColor { get; init; }
    public string? BarColor { get; init; }
    public string? IconStyle { get; init; }
}

internal sealed class ExcelDataValidationDef
{
    public string Range { get; init; } = "";
    public string Type { get; init; } = "list";
    public string Formula1 { get; init; } = "";
    public string? Formula2 { get; init; }
    public bool AllowBlank { get; init; } = true;
    public bool ShowDropdown { get; init; } = true;
    public string? PromptTitle { get; init; }
    public string? Prompt { get; init; }
    public string? ErrorTitle { get; init; }
    public string? Error { get; init; }
}

internal sealed class ExcelTableDef
{
    public string Range { get; init; } = "";
    public string Name { get; init; } = "";
    public string Style { get; init; } = "TableStyleMedium2";
    public bool ShowRowStripes { get; init; } = true;
    public bool ShowTotalsRow { get; init; }
}

internal sealed class ExcelDefinedNameDef
{
    public string Name { get; init; } = "";
    public string RefersTo { get; init; } = "";
    public bool Local { get; init; }
}

internal sealed class ExcelHyperlinkDef
{
    public string Cell { get; init; } = "";
    public string Url { get; init; } = "";
    public string? Tooltip { get; init; }
}

internal sealed class ExcelCommentDef
{
    public string Cell { get; init; } = "";
    public string Text { get; init; } = "";
    public string Author { get; init; } = "office-writer";
}

internal sealed class ExcelImageDef
{
    public string Path { get; init; } = "";
    public string Anchor { get; init; } = "A1";
    public double? Width { get; init; }
    public double? Height { get; init; }
}

internal sealed class ExcelPageSetupDef
{
    public string? Orientation { get; init; }
    public string? PaperSize { get; init; }
    public int? FitToWidth { get; init; }
    public int? FitToHeight { get; init; }
    public string? PrintArea { get; init; }
    public string? PrintTitleRows { get; init; }
    public string? PrintTitleCols { get; init; }
    public double? MarginTop { get; init; }
    public double? MarginBottom { get; init; }
    public double? MarginLeft { get; init; }
    public double? MarginRight { get; init; }
    public bool? CenterHorizontally { get; init; }
    public bool? CenterVertically { get; init; }
}

/// <summary>工作表保护；Allow* 为 true 表示允许该操作（与 OOXML 的禁止位相反）。</summary>
internal sealed class ExcelSheetProtectionDef
{
    public string? Password { get; init; }
    public bool AllowSelectLockedCells { get; init; } = true;
    public bool AllowSelectUnlockedCells { get; init; } = true;
    public bool AllowFormatCells { get; init; }
    public bool AllowFormatColumns { get; init; }
    public bool AllowFormatRows { get; init; }
    public bool AllowInsertRows { get; init; }
    public bool AllowInsertColumns { get; init; }
    public bool AllowDeleteRows { get; init; }
    public bool AllowDeleteColumns { get; init; }
    public bool AllowSort { get; init; }
    public bool AllowAutoFilter { get; init; }
}

internal sealed class ExcelSortKeyDef
{
    public string Column { get; init; } = "";
    public bool Ascending { get; init; } = true;
}

internal sealed class ExcelSortDef
{
    public string? Range { get; init; }
    public List<ExcelSortKeyDef> Keys { get; init; } = [];
    public bool HasHeader { get; init; } = true;
}

internal sealed class ExcelSheetDef
{
    public string Name { get; init; } = "";
    public List<string>? Headers { get; init; }
    public List<List<JsonElement>> Data { get; init; } = [];
    public bool AllowFormulas { get; init; }
    public List<double>? ColumnWidths { get; init; }
    public SortedDictionary<int, double>? RowHeights { get; init; }
    public ExcelCellStyle? HeaderStyle { get; init; }
    public List<ExcelColumnFormat>? ColumnFormats { get; init; }
    public List<string>? MergedCells { get; init; }
    public string? Freeze { get; init; }
    public bool AutoFilter { get; init; }
    public string? TabColor { get; init; }
    public string SheetState { get; init; } = "visible";
    public List<ExcelChartDef>? Charts { get; init; }
    public List<ExcelConditionalFormatDef>? ConditionalFormats { get; init; }
    public List<ExcelDataValidationDef>? DataValidations { get; init; }
    public List<ExcelTableDef>? Tables { get; init; }
    public List<ExcelDefinedNameDef>? DefinedNames { get; init; }
    public List<ExcelHyperlinkDef>? Hyperlinks { get; init; }
    public List<ExcelCommentDef>? Comments { get; init; }
    public List<ExcelImageDef>? Images { get; init; }
    public ExcelPageSetupDef? PageSetup { get; init; }
    public ExcelSheetProtectionDef? Protection { get; init; }
    public ExcelSortDef? Sort { get; init; }
}

// ── A1 引用工具 ──────────────────────────────────────────────────

internal readonly record struct ExcelRange(
    string? Sheet, int MinCol, int MinRow, int MaxCol, int MaxRow);

internal static class ExcelA1
{
    private static readonly Regex RangeRegex = new(
        @"^(?:(?:'(?<quoted>(?:[^']|'')+)'|(?<plain>[^!]+))!)?\$?(?<c1>[A-Za-z]{1,3})\$?(?<r1>\d+)(?::\$?(?<c2>[A-Za-z]{1,3})\$?(?<r2>\d+))?$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex CellRegex = new(
        @"^\$?(?<col>[A-Za-z]{1,3})\$?(?<row>\d+)$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex PlainSheetRegex = new(
        @"^[A-Za-z_\u4e00-\u9fff][A-Za-z0-9_.\u4e00-\u9fff]*$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public static bool TryParseRange(string text, out ExcelRange range)
    {
        range = default;
        var match = RangeRegex.Match(text.Trim());
        if (!match.Success)
            return false;
        var sheet = match.Groups["quoted"].Success
            ? match.Groups["quoted"].Value.Replace("''", "'", StringComparison.Ordinal)
            : match.Groups["plain"].Success
                ? match.Groups["plain"].Value
                : null;
        var minCol = ColumnIndex(match.Groups["c1"].Value);
        var minRow = int.Parse(match.Groups["r1"].Value, CultureInfo.InvariantCulture);
        var maxCol = match.Groups["c2"].Success ? ColumnIndex(match.Groups["c2"].Value) : minCol;
        var maxRow = match.Groups["r2"].Success
            ? int.Parse(match.Groups["r2"].Value, CultureInfo.InvariantCulture)
            : minRow;
        if (minCol < 1 || minRow < 1 || maxCol > 16384 || maxRow > 1048576)
            return false;
        range = new ExcelRange(
            sheet,
            Math.Min(minCol, maxCol),
            Math.Min(minRow, maxRow),
            Math.Max(minCol, maxCol),
            Math.Max(minRow, maxRow));
        return true;
    }

    public static bool TryParseCell(string text, out int col, out int row)
    {
        col = 0;
        row = 0;
        var match = CellRegex.Match(text.Trim());
        if (!match.Success)
            return false;
        col = ColumnIndex(match.Groups["col"].Value);
        row = int.Parse(match.Groups["row"].Value, CultureInfo.InvariantCulture);
        return col is >= 1 and <= 16384 && row is >= 1 and <= 1048576;
    }

    public static int ColumnIndex(string letters)
    {
        var index = 0;
        foreach (var ch in letters.ToUpperInvariant())
        {
            if (ch is < 'A' or > 'Z')
                return 0;
            index = index * 26 + (ch - 'A' + 1);
        }
        return index;
    }

    public static string ColumnName(int index)
    {
        if (index < 1)
            throw new ArgumentOutOfRangeException(nameof(index));
        Span<char> buffer = stackalloc char[4];
        var position = buffer.Length;
        while (index > 0)
        {
            index--;
            buffer[--position] = (char)('A' + index % 26);
            index /= 26;
        }
        return new string(buffer[position..]);
    }

    public static string CellRef(int col, int row) =>
        string.Concat(ColumnName(col), row.ToString(CultureInfo.InvariantCulture));

    public static string FormatRange(int minCol, int minRow, int maxCol, int maxRow) =>
        $"{CellRef(minCol, minRow)}:{CellRef(maxCol, maxRow)}";

    public static string QuoteSheet(string name) =>
        PlainSheetRegex.IsMatch(name) ? name : "'" + name.Replace("'", "''", StringComparison.Ordinal) + "'";

    public static string AbsoluteRangeFormula(
        string sheet, int minCol, int minRow, int maxCol, int maxRow) =>
        $"{QuoteSheet(sheet)}!${ColumnName(minCol)}${minRow}:${ColumnName(maxCol)}${maxRow}";

    public static string AbsoluteCellFormula(string sheet, int col, int row) =>
        $"{QuoteSheet(sheet)}!${ColumnName(col)}${row}";

    public static string AbsoluteRowRangeFormula(string sheet, int minRow, int maxRow) =>
        $"{QuoteSheet(sheet)}!${minRow}:${maxRow}";

    public static string AbsoluteColumnRangeFormula(string sheet, int minCol, int maxCol) =>
        $"{QuoteSheet(sheet)}!${ColumnName(minCol)}:${ColumnName(maxCol)}";
}

// ── 颜色与字号 ──────────────────────────────────────────────────

internal static class ExcelColor
{
    private static readonly Regex HexRegex = new(
        @"^#?(?<hex>[0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public static string Normalize(string value, string label = "颜色")
    {
        var match = HexRegex.Match(value.Trim());
        if (!match.Success)
            throw new NativeUsageException($"{label}必须是 6 位或 8 位 HEX（如 'FF0000'）: {value}");
        var hex = match.Groups["hex"].Value.ToUpperInvariant();
        return hex.Length == 6 ? "FF" + hex : hex;
    }
}

internal static class ExcelFontSize
{
    private static readonly Dictionary<string, double> ChineseSizes = new(StringComparer.Ordinal)
    {
        ["初号"] = 42, ["小初"] = 36, ["一号"] = 26, ["小一"] = 24,
        ["二号"] = 22, ["小二"] = 18, ["三号"] = 16, ["小三"] = 15,
        ["四号"] = 14, ["小四"] = 12, ["五号"] = 10.5, ["小五"] = 9,
        ["六号"] = 7.5, ["小六"] = 6.5, ["七号"] = 5.5, ["八号"] = 5,
    };

    public static double? Parse(string? value, string label = "font_size")
    {
        if (string.IsNullOrWhiteSpace(value))
            return null;
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var size))
            return size;
        if (ChineseSizes.TryGetValue(value.Trim(), out var mapped))
            return mapped;
        throw new NativeUsageException($"{label} 既不是数字也不是中文字号: {value}");
    }
}

// ── JSON 解析 ───────────────────────────────────────────────────

internal static class ExcelModelParser
{
    private static readonly HashSet<string> ChartTypes = new(StringComparer.Ordinal)
    {
        "column", "bar", "line", "pie", "doughnut", "area", "scatter", "radar",
    };

    private static readonly HashSet<string> Legends = new(StringComparer.Ordinal)
    {
        "right", "bottom", "top", "left", "none",
    };

    private static readonly HashSet<string> ConditionalTypes = new(StringComparer.Ordinal)
    {
        "cell_is", "formula", "color_scale", "data_bar", "icon_set",
    };

    private static readonly HashSet<string> ConditionalOperators = new(StringComparer.Ordinal)
    {
        "greaterThan", "lessThan", "between", "equal", "notEqual",
        "greaterThanOrEqual", "lessThanOrEqual",
    };

    private static readonly HashSet<string> ValidationTypes = new(StringComparer.Ordinal)
    {
        "list", "whole", "decimal", "date", "time", "textLength", "custom",
    };

    private static readonly HashSet<string> SheetStates = new(StringComparer.Ordinal)
    {
        "visible", "hidden", "veryHidden",
    };

    private static readonly HashSet<string> Alignments = new(StringComparer.Ordinal)
    {
        "left", "center", "right",
    };

    private static readonly HashSet<string> VerticalAlignments = new(StringComparer.Ordinal)
    {
        "top", "center", "bottom",
    };

    private static readonly HashSet<string> BorderStyles = new(StringComparer.Ordinal)
    {
        "thin", "medium", "thick", "dashed", "dotted", "double", "hair",
    };

    private static readonly HashSet<string> PaperSizes = new(StringComparer.Ordinal)
    {
        "A4", "Letter", "A3", "B5",
    };

    public static List<ExcelSheetDef> ParseSheets(JsonElement p)
    {
        if (p.ValueKind != JsonValueKind.Object
            || !p.TryGetProperty("sheets", out var sheets)
            || sheets.ValueKind != JsonValueKind.Array)
        {
            throw new NativeUsageException("缺少参数 sheets（JSON 数组）");
        }
        var result = new List<ExcelSheetDef>();
        var index = 0;
        foreach (var item in sheets.EnumerateArray())
        {
            result.Add(ParseSheet(item, $"sheets[{index}]"));
            index++;
        }
        return result;
    }

    public static List<ExcelSheetDef> ParseAppendSheets(JsonElement p)
    {
        if (p.ValueKind != JsonValueKind.Object
            || !p.TryGetProperty("append_sheets", out var sheets)
            || sheets.ValueKind != JsonValueKind.Array)
        {
            throw new NativeUsageException("缺少参数 append_sheets（JSON 数组）");
        }
        var result = new List<ExcelSheetDef>();
        var index = 0;
        foreach (var item in sheets.EnumerateArray())
        {
            result.Add(ParseSheet(item, $"append_sheets[{index}]"));
            index++;
        }
        return result;
    }

    public static ExcelSheetDef ParseSheet(JsonElement e, string label)
    {
        if (e.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException($"{label} 必须是 JSON 对象");

        var sheet = new ExcelSheetDef
        {
            Name = NativeCommon.RequireString(e, "name"),
            Headers = ParseStringList(Get(e, "headers"), $"{label}.headers"),
            Data = ParseData(Get(e, "data"), $"{label}.data"),
            AllowFormulas = NativeCommon.OptBool(e, "allow_formulas", false),
            ColumnWidths = ParseDoubleList(Get(e, "column_widths"), $"{label}.column_widths"),
            RowHeights = ParseRowHeights(Get(e, "row_heights"), $"{label}.row_heights"),
            HeaderStyle = Get(e, "header_style") is { } hs ? ParseCellStyle(hs, $"{label}.header_style") : null,
            ColumnFormats = ParseColumnFormats(Get(e, "column_formats"), $"{label}.column_formats"),
            MergedCells = ParseMergeRanges(Get(e, "merged_cells"), $"{label}.merged_cells"),
            Freeze = OptCell(e, "freeze", $"{label}.freeze"),
            AutoFilter = NativeCommon.OptBool(e, "auto_filter", false),
            TabColor = OptColor(e, "tab_color", $"{label}.tab_color"),
            SheetState = OptChoice(e, "sheet_state", "visible", SheetStates, $"{label}.sheet_state"),
            Charts = ParseList(Get(e, "charts"), $"{label}.charts", ParseChart),
            ConditionalFormats = ParseList(Get(e, "conditional_formats"), $"{label}.conditional_formats", ParseConditionalFormat),
            DataValidations = ParseList(Get(e, "data_validations"), $"{label}.data_validations", ParseDataValidation),
            Tables = ParseList(Get(e, "tables"), $"{label}.tables", ParseTable),
            DefinedNames = ParseList(Get(e, "defined_names"), $"{label}.defined_names", ParseDefinedName),
            Hyperlinks = ParseList(Get(e, "hyperlinks"), $"{label}.hyperlinks", ParseHyperlink),
            Comments = ParseList(Get(e, "comments"), $"{label}.comments", ParseComment),
            Images = ParseList(Get(e, "images"), $"{label}.images", ParseImage),
            PageSetup = Get(e, "page_setup") is { } ps ? ParsePageSetup(ps, $"{label}.page_setup") : null,
            Protection = Get(e, "protection") is { } pr
                ? ParseProtection(pr, $"{label}.protection")
                : null,
            Sort = Get(e, "sort") is { } st ? ParseSort(st, $"{label}.sort") : null,
        };
        return sheet;
    }

    private static ExcelSheetProtectionDef ParseProtection(JsonElement e, string label)
    {
        RequireObject(e, label);
        return new ExcelSheetProtectionDef
        {
            Password = NativeCommon.OptString(e, "password"),
            AllowSelectLockedCells = NativeCommon.OptBool(e, "allow_select_locked_cells", true),
            AllowSelectUnlockedCells = NativeCommon.OptBool(e, "allow_select_unlocked_cells", true),
            AllowFormatCells = NativeCommon.OptBool(e, "allow_format_cells", false),
            AllowFormatColumns = NativeCommon.OptBool(e, "allow_format_columns", false),
            AllowFormatRows = NativeCommon.OptBool(e, "allow_format_rows", false),
            AllowInsertRows = NativeCommon.OptBool(e, "allow_insert_rows", false),
            AllowInsertColumns = NativeCommon.OptBool(e, "allow_insert_columns", false),
            AllowDeleteRows = NativeCommon.OptBool(e, "allow_delete_rows", false),
            AllowDeleteColumns = NativeCommon.OptBool(e, "allow_delete_columns", false),
            AllowSort = NativeCommon.OptBool(e, "allow_sort", false),
            AllowAutoFilter = NativeCommon.OptBool(e, "allow_auto_filter", false),
        };
    }

    private static ExcelSortDef ParseSort(JsonElement e, string label)
    {
        RequireObject(e, label);
        // range 与排序列的语义校验统一放到写出阶段（NativeRuntimeException，exit 3），
        // 与 Python 后端在 excel_writer 中抛 ValueError 的行为一致。
        var range = NativeCommon.OptString(e, "range");
        if (string.IsNullOrWhiteSpace(range))
            range = null;
        if (Get(e, "keys") is not { } keysElement || keysElement.ValueKind != JsonValueKind.Array)
            throw new NativeUsageException($"{label}.keys 必须是 JSON 数组");
        var keys = new List<ExcelSortKeyDef>();
        var index = 0;
        foreach (var item in keysElement.EnumerateArray())
        {
            RequireObject(item, $"{label}.keys[{index}]");
            var column = NativeCommon.OptString(item, "column");
            // 缺字段属于参数错误（exit 2）；空字符串留给写出阶段报语义错误（exit 3），
            // 与 Python 后端 Pydantic 通过后由 excel_writer 抛 ValueError 的行为一致。
            if (column is null)
                throw new NativeUsageException($"{label}.keys[{index}].column 不能为空");
            keys.Add(new ExcelSortKeyDef
            {
                Column = column.Trim(),
                Ascending = NativeCommon.OptBool(item, "ascending", true),
            });
            index++;
        }
        if (keys.Count == 0)
            throw new NativeUsageException($"{label}.keys 至少需要 1 个排序键");
        return new ExcelSortDef
        {
            Range = range?.Trim(),
            Keys = keys,
            HasHeader = NativeCommon.OptBool(e, "has_header", true),
        };
    }

    public static ExcelChartDef ParseChart(JsonElement e, string label)
    {
        RequireObject(e, label);
        var chartType = OptChoice(e, "chart_type", "column", ChartTypes, $"{label}.chart_type");
        var dataRange = NativeCommon.RequireString(e, "data_range");
        ValidateRangeText(dataRange, $"{label}.data_range", "图表数据范围");
        var categories = NativeCommon.OptString(e, "categories_range");
        if (categories is not null)
            ValidateRangeText(categories, $"{label}.categories_range", "图表分类范围");
        var anchorText = NativeCommon.OptString(e, "anchor") ?? "E2";
        if (!ExcelA1.TryParseCell(anchorText, out _, out _))
            throw new NativeUsageException($"{label}.anchor 不是有效单元格: {anchorText}");
        var legend = OptChoice(e, "legend", "right", Legends, $"{label}.legend");
        var width = NativeCommon.OptDouble(e, "width", 15);
        var height = NativeCommon.OptDouble(e, "height", 8);
        if (!double.IsFinite(width) || width <= 0 || width > ExcelFileDrawing.MaxDrawingSizeCm)
            throw new NativeUsageException(
                $"{label}.width 必须是 (0, {ExcelFileDrawing.MaxDrawingSizeCm}] 之间的有限数字");
        if (!double.IsFinite(height) || height <= 0 || height > ExcelFileDrawing.MaxDrawingSizeCm)
            throw new NativeUsageException(
                $"{label}.height 必须是 (0, {ExcelFileDrawing.MaxDrawingSizeCm}] 之间的有限数字");
        var style = OptNullableInt(e, "style");
        if (style is < 1 or > 48)
            throw new NativeUsageException($"{label}.style 必须在 1-48 之间");
        return new ExcelChartDef
        {
            ChartType = chartType,
            Title = NativeCommon.OptString(e, "title"),
            DataRange = dataRange,
            CategoriesRange = categories,
            TitlesFromData = NativeCommon.OptBool(e, "titles_from_data", true),
            Anchor = anchorText.Trim().ToUpperInvariant().Replace("$", "", StringComparison.Ordinal),
            Width = width,
            Height = height,
            Legend = legend,
            XAxisTitle = NativeCommon.OptString(e, "x_axis_title"),
            YAxisTitle = NativeCommon.OptString(e, "y_axis_title"),
            Style = style,
        };
    }

    private static ExcelConditionalFormatDef ParseConditionalFormat(JsonElement e, string label)
    {
        RequireObject(e, label);
        var range = NativeCommon.RequireString(e, "range");
        ValidateRangeText(range, $"{label}.range", "条件格式范围", allowSheet: false);
        var type = OptChoice(e, "type", "cell_is", ConditionalTypes, $"{label}.type");
        var op = NativeCommon.OptString(e, "operator");
        if (op is not null && !ConditionalOperators.Contains(op))
            throw new NativeUsageException($"{label}.operator 不受支持: {op}");
        var formula = ParseStringList(Get(e, "formula"), $"{label}.formula");
        if (type is "cell_is" or "formula" && (formula is null || formula.Count == 0))
            throw new NativeUsageException($"{label} 的 {type} 类型必须提供 formula");
        if (type == "cell_is" && op is null)
            throw new NativeUsageException($"{label} 的 cell_is 类型必须提供 operator");
        return new ExcelConditionalFormatDef
        {
            Range = range,
            Type = type,
            Operator = op,
            Formula = formula,
            FillColor = OptColor(e, "fill_color", $"{label}.fill_color"),
            FontColor = OptColor(e, "font_color", $"{label}.font_color"),
            MinColor = OptColor(e, "min_color", $"{label}.min_color"),
            MidColor = OptColor(e, "mid_color", $"{label}.mid_color"),
            MaxColor = OptColor(e, "max_color", $"{label}.max_color"),
            BarColor = OptColor(e, "bar_color", $"{label}.bar_color"),
            IconStyle = NativeCommon.OptString(e, "icon_style"),
        };
    }

    private static ExcelDataValidationDef ParseDataValidation(JsonElement e, string label)
    {
        RequireObject(e, label);
        var range = NativeCommon.RequireString(e, "range");
        ValidateRangeText(range, $"{label}.range", "数据验证范围", allowSheet: false);
        var type = OptChoice(e, "type", "list", ValidationTypes, $"{label}.type");
        var formula1 = NativeCommon.RequireString(e, "formula1");
        return new ExcelDataValidationDef
        {
            Range = range,
            Type = type,
            Formula1 = formula1,
            Formula2 = NativeCommon.OptString(e, "formula2"),
            AllowBlank = NativeCommon.OptBool(e, "allow_blank", true),
            ShowDropdown = NativeCommon.OptBool(e, "show_dropdown", true),
            PromptTitle = NativeCommon.OptString(e, "prompt_title"),
            Prompt = NativeCommon.OptString(e, "prompt"),
            ErrorTitle = NativeCommon.OptString(e, "error_title"),
            Error = NativeCommon.OptString(e, "error"),
        };
    }

    private static ExcelTableDef ParseTable(JsonElement e, string label)
    {
        RequireObject(e, label);
        var range = NativeCommon.RequireString(e, "range");
        ValidateRangeText(range, $"{label}.range", "表格范围", allowSheet: false);
        var name = NativeCommon.RequireString(e, "name");
        return new ExcelTableDef
        {
            Range = range,
            Name = name,
            Style = NativeCommon.OptString(e, "style") ?? "TableStyleMedium2",
            ShowRowStripes = NativeCommon.OptBool(e, "show_row_stripes", true),
            ShowTotalsRow = NativeCommon.OptBool(e, "show_totals_row", false),
        };
    }

    private static ExcelDefinedNameDef ParseDefinedName(JsonElement e, string label)
    {
        RequireObject(e, label);
        return new ExcelDefinedNameDef
        {
            Name = NativeCommon.RequireString(e, "name"),
            RefersTo = NativeCommon.RequireString(e, "refers_to"),
            Local = NativeCommon.OptBool(e, "local", false),
        };
    }

    private static ExcelHyperlinkDef ParseHyperlink(JsonElement e, string label)
    {
        RequireObject(e, label);
        var cell = NativeCommon.RequireString(e, "cell");
        if (!ExcelA1.TryParseCell(cell, out _, out _))
            throw new NativeUsageException($"{label}.cell 不是有效单元格: {cell}");
        return new ExcelHyperlinkDef
        {
            Cell = cell,
            Url = NativeCommon.RequireString(e, "url"),
            Tooltip = NativeCommon.OptString(e, "tooltip"),
        };
    }

    private static ExcelCommentDef ParseComment(JsonElement e, string label)
    {
        RequireObject(e, label);
        var cell = NativeCommon.RequireString(e, "cell");
        if (!ExcelA1.TryParseCell(cell, out _, out _))
            throw new NativeUsageException($"{label}.cell 不是有效单元格: {cell}");
        return new ExcelCommentDef
        {
            Cell = cell,
            Text = NativeCommon.RequireString(e, "text"),
            Author = NativeCommon.OptString(e, "author") ?? "office-writer",
        };
    }

    private static ExcelImageDef ParseImage(JsonElement e, string label)
    {
        RequireObject(e, label);
        var anchor = NativeCommon.OptString(e, "anchor") ?? "A1";
        if (!ExcelA1.TryParseCell(anchor, out _, out _))
            throw new NativeUsageException($"{label}.anchor 不是有效单元格: {anchor}");
        var width = OptNullableDouble(e, "width");
        var height = OptNullableDouble(e, "height");
        if (width is { } widthValue
            && (!double.IsFinite(widthValue) || widthValue <= 0
                || widthValue > ExcelFileDrawing.MaxDrawingSizeCm))
        {
            throw new NativeUsageException(
                $"{label}.width 必须是 (0, {ExcelFileDrawing.MaxDrawingSizeCm}] 之间的有限数字");
        }
        if (height is { } heightValue
            && (!double.IsFinite(heightValue) || heightValue <= 0
                || heightValue > ExcelFileDrawing.MaxDrawingSizeCm))
        {
            throw new NativeUsageException(
                $"{label}.height 必须是 (0, {ExcelFileDrawing.MaxDrawingSizeCm}] 之间的有限数字");
        }
        return new ExcelImageDef
        {
            Path = NativeCommon.RequireString(e, "path"),
            Anchor = anchor,
            Width = width,
            Height = height,
        };
    }

    private static ExcelPageSetupDef ParsePageSetup(JsonElement e, string label)
    {
        RequireObject(e, label);
        var orientation = NativeCommon.OptString(e, "orientation");
        if (orientation is not null && orientation is not ("portrait" or "landscape"))
            throw new NativeUsageException($"{label}.orientation 必须是 portrait/landscape");
        var paper = NativeCommon.OptString(e, "paper_size");
        if (paper is not null && !PaperSizes.Contains(paper))
            throw new NativeUsageException($"{label}.paper_size 必须是 A4/Letter/A3/B5");
        var fitWidth = OptNullableInt(e, "fit_to_width");
        var fitHeight = OptNullableInt(e, "fit_to_height");
        if (fitWidth is < 0 || fitHeight is < 0)
            throw new NativeUsageException($"{label}.fit_to_width/fit_to_height 不能为负");
        return new ExcelPageSetupDef
        {
            Orientation = orientation,
            PaperSize = paper,
            FitToWidth = fitWidth,
            FitToHeight = fitHeight,
            PrintArea = OptRangeText(e, "print_area", $"{label}.print_area"),
            PrintTitleRows = OptRowColText(e, "print_title_rows", $"{label}.print_title_rows", true),
            PrintTitleCols = OptRowColText(e, "print_title_cols", $"{label}.print_title_cols", false),
            MarginTop = OptNullableDouble(e, "margin_top"),
            MarginBottom = OptNullableDouble(e, "margin_bottom"),
            MarginLeft = OptNullableDouble(e, "margin_left"),
            MarginRight = OptNullableDouble(e, "margin_right"),
            CenterHorizontally = OptNullableBool(e, "center_horizontally"),
            CenterVertically = OptNullableBool(e, "center_vertically"),
        };
    }

    public static ExcelCellStyle ParseCellStyle(JsonElement e, string label)
    {
        RequireObject(e, label);
        return new ExcelCellStyle
        {
            Bold = OptNullableBool(e, "bold"),
            Italic = OptNullableBool(e, "italic"),
            FontName = NativeCommon.OptString(e, "font_name"),
            FontSize = ParseFontSize(e, "font_size", $"{label}.font_size"),
            FontColor = OptColor(e, "font_color", $"{label}.font_color"),
            BgColor = OptColor(e, "bg_color", $"{label}.bg_color"),
            NumberFormat = NativeCommon.OptString(e, "number_format"),
            Alignment = OptChoiceOrNull(e, "alignment", Alignments, $"{label}.alignment"),
            VerticalAlignment = OptChoiceOrNull(e, "vertical_alignment", VerticalAlignments, $"{label}.vertical_alignment"),
            WrapText = OptNullableBool(e, "wrap_text"),
            Border = Get(e, "border") is { } b ? ParseBorder(b, $"{label}.border") : null,
        };
    }

    private static ExcelBorderSpec ParseBorder(JsonElement e, string label)
    {
        RequireObject(e, label);
        var style = OptChoice(e, "style", "thin", BorderStyles, $"{label}.style");
        return new ExcelBorderSpec
        {
            Style = style,
            Color = OptColor(e, "color", $"{label}.color"),
        };
    }

    private static ExcelColumnFormat ParseColumnFormat(JsonElement e, string label)
    {
        RequireObject(e, label);
        var width = OptNullableDouble(e, "width");
        if (width is < 0)
            throw new NativeUsageException($"{label}.width 不能为负");
        return new ExcelColumnFormat
        {
            NumberFormat = NativeCommon.OptString(e, "number_format"),
            FontName = NativeCommon.OptString(e, "font_name"),
            FontSize = ParseFontSize(e, "font_size", $"{label}.font_size"),
            FontColor = OptColor(e, "font_color", $"{label}.font_color"),
            BgColor = OptColor(e, "bg_color", $"{label}.bg_color"),
            Bold = OptNullableBool(e, "bold"),
            Italic = OptNullableBool(e, "italic"),
            WrapText = OptNullableBool(e, "wrap_text"),
            Width = width,
            Alignment = OptChoiceOrNull(e, "alignment", Alignments, $"{label}.alignment"),
            VerticalAlignment = OptChoiceOrNull(e, "vertical_alignment", VerticalAlignments, $"{label}.vertical_alignment"),
        };
    }

    // ── 基础读取 ─────────────────────────────────────────────────

    public static JsonElement? Get(JsonElement e, string name)
    {
        if (e.ValueKind != JsonValueKind.Object || !e.TryGetProperty(name, out var value)
            || value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;
        return value;
    }

    private static void RequireObject(JsonElement e, string label)
    {
        if (e.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException($"{label} 必须是 JSON 对象");
    }

    private static string OptChoice(
        JsonElement e, string name, string fallback, HashSet<string> allowed, string label)
    {
        var value = NativeCommon.OptString(e, name) ?? fallback;
        if (!allowed.Contains(value))
            throw new NativeUsageException($"{label} 必须是 {string.Join("/", allowed)} 之一，收到 {value}");
        return value;
    }

    private static string? OptChoiceOrNull(
        JsonElement e, string name, HashSet<string> allowed, string label)
    {
        var value = NativeCommon.OptString(e, name);
        if (value is null)
            return null;
        if (!allowed.Contains(value))
            throw new NativeUsageException($"{label} 必须是 {string.Join("/", allowed)} 之一，收到 {value}");
        return value;
    }

    private static string? OptColor(JsonElement e, string name, string label)
    {
        var value = NativeCommon.OptString(e, name);
        return string.IsNullOrWhiteSpace(value) ? null : ExcelColor.Normalize(value, label + " ");
    }

    private static bool? OptNullableBool(JsonElement e, string name)
    {
        if (e.ValueKind != JsonValueKind.Object || !e.TryGetProperty(name, out var value)
            || value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;
        if (value.ValueKind is JsonValueKind.True or JsonValueKind.False)
            return value.GetBoolean();
        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var number)
            && number is 0 or 1)
            return number == 1;
        if (value.ValueKind == JsonValueKind.String)
        {
            var text = value.GetString()?.Trim();
            if (text is "0" or "1")
                return text == "1";
            if (bool.TryParse(text, out var parsed))
                return parsed;
        }
        throw new NativeUsageException($"参数 {name} 必须是布尔值");
    }

    private static int? OptNullableInt(JsonElement e, string name)
    {
        if (e.ValueKind != JsonValueKind.Object || !e.TryGetProperty(name, out var value)
            || value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;
        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var number))
            return number;
        if (value.ValueKind == JsonValueKind.String && int.TryParse(value.GetString(), out var parsed))
            return parsed;
        throw new NativeUsageException($"参数 {name} 必须是整数");
    }

    private static double? OptNullableDouble(JsonElement e, string name)
    {
        if (e.ValueKind != JsonValueKind.Object || !e.TryGetProperty(name, out var value)
            || value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;
        if (value.ValueKind == JsonValueKind.Number && value.TryGetDouble(out var number))
            return number;
        if (value.ValueKind == JsonValueKind.String
            && double.TryParse(value.GetString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed))
            return parsed;
        throw new NativeUsageException($"参数 {name} 必须是数字");
    }

    private static double? ParseFontSize(JsonElement e, string name, string label)
    {
        if (e.ValueKind != JsonValueKind.Object || !e.TryGetProperty(name, out var value)
            || value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;
        return value.ValueKind switch
        {
            JsonValueKind.Number => value.GetDouble(),
            JsonValueKind.String => ExcelFontSize.Parse(value.GetString(), label),
            _ => throw new NativeUsageException($"{label} 必须是数字或中文字号"),
        };
    }

    private static List<string>? ParseStringList(JsonElement? element, string label)
    {
        if (element is null || element.Value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;
        if (element.Value.ValueKind != JsonValueKind.Array)
            throw new NativeUsageException($"{label} 必须是 JSON 数组");
        var result = new List<string>();
        foreach (var item in element.Value.EnumerateArray())
        {
            result.Add(item.ValueKind switch
            {
                JsonValueKind.String => item.GetString()!,
                JsonValueKind.Null or JsonValueKind.Undefined => "",
                JsonValueKind.Number => item.GetRawText(),
                JsonValueKind.True => "True",
                JsonValueKind.False => "False",
                _ => throw new NativeUsageException($"{label} 的列表项必须是标量"),
            });
        }
        return result;
    }

    private static List<double>? ParseDoubleList(JsonElement? element, string label)
    {
        if (element is null || element.Value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;
        if (element.Value.ValueKind != JsonValueKind.Array)
            throw new NativeUsageException($"{label} 必须是 JSON 数组");
        var result = new List<double>();
        foreach (var item in element.Value.EnumerateArray())
        {
            if (item.ValueKind == JsonValueKind.Number && item.TryGetDouble(out var number))
                result.Add(number);
            else if (item.ValueKind == JsonValueKind.String
                     && double.TryParse(item.GetString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed))
                result.Add(parsed);
            else
                throw new NativeUsageException($"{label} 的列表项必须是数字");
        }
        return result;
    }

    private static List<List<JsonElement>> ParseData(JsonElement? element, string label)
    {
        var result = new List<List<JsonElement>>();
        if (element is null || element.Value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return result;
        if (element.Value.ValueKind != JsonValueKind.Array)
            throw new NativeUsageException($"{label} 必须是二维 JSON 数组");
        foreach (var row in element.Value.EnumerateArray())
        {
            if (row.ValueKind != JsonValueKind.Array)
                throw new NativeUsageException($"{label} 的每一行必须是 JSON 数组");
            result.Add(row.EnumerateArray().ToList());
        }
        return result;
    }

    private static SortedDictionary<int, double>? ParseRowHeights(JsonElement? element, string label)
    {
        if (element is null || element.Value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;
        if (element.Value.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException($"{label} 必须是 JSON 对象（行号→磅值）");
        var result = new SortedDictionary<int, double>();
        foreach (var property in element.Value.EnumerateObject())
        {
            if (!int.TryParse(property.Name, NumberStyles.Integer, CultureInfo.InvariantCulture, out var row)
                || row < 1)
                throw new NativeUsageException($"{label} 的行号必须是从 1 开始的整数: {property.Name}");
            double height;
            if (property.Value.ValueKind == JsonValueKind.Number && property.Value.TryGetDouble(out var number))
                height = number;
            else if (property.Value.ValueKind == JsonValueKind.String
                     && double.TryParse(property.Value.GetString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed))
                height = parsed;
            else
                throw new NativeUsageException($"{label}[{property.Name}] 必须是数字");
            if (height < 0)
                throw new NativeUsageException($"{label}[{property.Name}] 不能为负");
            result[row] = height;
        }
        return result;
    }

    private static List<string>? ParseMergeRanges(JsonElement? element, string label)
    {
        if (element is null || element.Value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;
        if (element.Value.ValueKind != JsonValueKind.Array)
            throw new NativeUsageException($"{label} 必须是 JSON 数组");
        var result = new List<string>();
        foreach (var item in element.Value.EnumerateArray())
        {
            string? range = item.ValueKind switch
            {
                JsonValueKind.String => item.GetString(),
                JsonValueKind.Object => NativeCommon.OptString(item, "range"),
                _ => null,
            };
            if (string.IsNullOrWhiteSpace(range))
                throw new NativeUsageException($"{label} 的列表项必须提供 range");
            ValidateRangeText(range, label, "合并单元格", allowSheet: false);
            if (!ExcelA1.TryParseRange(range, out var parsed))
                throw new NativeUsageException($"{label} 的范围无效: {range}");
            result.Add(ExcelA1.FormatRange(parsed.MinCol, parsed.MinRow, parsed.MaxCol, parsed.MaxRow));
        }
        return result;
    }

    private static List<T>? ParseList<T>(
        JsonElement? element, string label, Func<JsonElement, string, T> parser)
    {
        if (element is null || element.Value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;
        if (element.Value.ValueKind != JsonValueKind.Array)
            throw new NativeUsageException($"{label} 必须是 JSON 数组");
        var result = new List<T>();
        var index = 0;
        foreach (var item in element.Value.EnumerateArray())
        {
            result.Add(parser(item, $"{label}[{index}]"));
            index++;
        }
        return result;
    }

    private static List<T>? ParseList<T>(
        JsonElement? element, string label, Func<JsonElement, T> parser)
    {
        return ParseList(element, label, (item, _) => parser(item));
    }

    private static List<ExcelColumnFormat>? ParseColumnFormats(JsonElement? element, string label)
    {
        if (element is null || element.Value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;
        if (element.Value.ValueKind != JsonValueKind.Array)
            throw new NativeUsageException($"{label} 必须是 JSON 数组");
        var result = new List<ExcelColumnFormat>();
        var index = 0;
        foreach (var item in element.Value.EnumerateArray())
        {
            result.Add(ParseColumnFormat(item, $"{label}[{index}]"));
            index++;
        }
        return result;
    }

    private static string? OptCell(JsonElement e, string name, string label)
    {
        var value = NativeCommon.OptString(e, name);
        if (string.IsNullOrWhiteSpace(value))
            return null;
        if (!ExcelA1.TryParseCell(value, out _, out _))
            throw new NativeUsageException($"{label} 不是有效单元格: {value}");
        return value.Trim().ToUpperInvariant().Replace("$", "", StringComparison.Ordinal);
    }

    private static string? OptRangeText(JsonElement e, string name, string label)
    {
        var value = NativeCommon.OptString(e, name);
        if (string.IsNullOrWhiteSpace(value))
            return null;
        foreach (var piece in value.Split(','))
            ValidateRangeText(piece, label, "打印区域");
        return value.Trim();
    }

    private static string? OptRowColText(JsonElement e, string name, string label, bool rows)
    {
        var value = NativeCommon.OptString(e, name);
        if (string.IsNullOrWhiteSpace(value))
            return null;
        var regex = rows ? @"^\d+(?::\d+)?$" : @"^[A-Za-z]{1,3}(?::[A-Za-z]{1,3})?$";
        if (!Regex.IsMatch(value.Trim(), regex, RegexOptions.CultureInvariant))
            throw new NativeUsageException($"{label} 格式无效: {value}");
        return value.Trim();
    }

    public static void ValidateRangeText(
        string value, string label, string friendly, bool allowSheet = true)
    {
        if (!ExcelA1.TryParseRange(value, out var range))
            throw new NativeUsageException(
                $"{label} 必须是含具体行列的 A1 范围（如 A1:C10），不支持整列/整行或空值: {value}");
        if (!allowSheet && range.Sheet is not null)
            throw new NativeUsageException(
                $"{label} 不能带工作表前缀（{friendly}只能作用于当前工作表）: {value}");
    }
}
