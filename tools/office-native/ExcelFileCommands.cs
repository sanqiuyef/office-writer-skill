using System.Text.Json;

namespace OfficeNative;

/// <summary>Excel 文件级能力入口；由 CommandRouter 统一注册。</summary>
internal static class ExcelFileCommands
{
    public static Dictionary<string, object?> RunCreateWorkbook(JsonElement p)
    {
        var outputPath = NativeCommon.RequireString(p, "output_path");
        var template = NativeCommon.OptString(p, "template");
        // 与 Python 的 `if template:` 一致：空字符串视为不使用模板。
        if (string.IsNullOrEmpty(template))
            template = null;
        var sheets = ExcelModelParser.ParseSheets(p);
        var ifExists = NativeCommon.OptString(p, "if_exists") ?? "error";
        var confirmOverwrite = NativeCommon.OptBool(p, "confirm_overwrite", false);
        return ExcelFileOoxml.CreateWorkbook(
            outputPath, template, sheets, ifExists, confirmOverwrite);
    }

    public static Dictionary<string, object?> RunEditWorkbook(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        var appendSheets = ExcelModelParser.ParseAppendSheets(p);
        // 文件已在活动 Excel/WPS 实例中打开时优先走 COM，避免用内存副本覆盖用户编辑。
        var active = OpenEditCommon.FindExcel(filePath, out var workbook, out _);
        if (active is not null)
        {
            NativeCommon.ReleaseCom((object?)workbook);
            NativeCommon.ReleaseCom((object?)active);
            // 活动文档 COM 追加路径没有实现工作表保护/排序；显式失败，不静默丢弃。
            foreach (var sheet in appendSheets)
            {
                if (sheet.Protection is not null || sheet.Sort is not null)
                {
                    throw new NativeRuntimeException(
                        $"工作表 '{sheet.Name}' 使用了 protection/sort 字段，活动 Excel 追加暂不支持；"
                        + "请关闭文件后使用文件级 edit_excel_workbook");
                }
            }
            return OpenEditExcelCommands.RunEditOpenExcelWorkbook(p);
        }
        return ExcelFileOoxml.EditWorkbook(filePath, appendSheets);
    }

    public static Dictionary<string, object?> RunReadWorkbook(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        var maxRows = NativeCommon.OptInt(p, "max_rows_per_sheet", 1000);
        var maxChars = NativeCommon.OptInt(p, "max_chars", 50000);
        return ExcelFileOoxml.ReadWorkbook(filePath, maxRows, maxChars);
    }

    public static Dictionary<string, object?> RunAddChart(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        var sheet = NativeCommon.OptString(p, "sheet");
        if (ExcelModelParser.Get(p, "chart") is not { } chartElement)
            throw new NativeUsageException("缺少参数 chart（JSON 对象）");
        var chart = ExcelModelParser.ParseChart(chartElement, "chart");
        return ExcelFileOoxml.AddChart(filePath, sheet, chart);
    }

    public static Dictionary<string, object?> RunMergeTemplate(JsonElement p)
    {
        var templatePath = NativeCommon.RequireString(p, "template_path");
        var outputPath = NativeCommon.RequireString(p, "output_path");
        if (ExcelModelParser.Get(p, "data") is not { } data
            || data.ValueKind != JsonValueKind.Object)
        {
            throw new NativeUsageException(
                "缺少参数 data（JSON 对象，用于填充 {{key}} 占位符）");
        }
        var ifExists = NativeCommon.OptString(p, "if_exists") ?? "error";
        var confirmOverwrite = NativeCommon.OptBool(p, "confirm_overwrite", false);
        return ExcelFileOoxml.MergeTemplate(
            templatePath, outputPath, data, ifExists, confirmOverwrite);
    }
}
