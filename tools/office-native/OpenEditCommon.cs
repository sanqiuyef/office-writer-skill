using System.Globalization;
using System.Text.Json;

namespace OfficeNative;

/// <summary>
/// 活动文档 COM 编辑的共享辅助：定位已打开实例、路径比较、颜色转换与结果构造。
/// 只连接已运行的 Office/WPS 实例，绝不启动或退出用户的应用。
/// </summary>
internal static class OpenEditCommon
{
    public static readonly string[] WordProgIds = ["Word.Application", "kwps.Application"];

    public static readonly string[] ExcelProgIds =
        ["Excel.Application", "KET.Application", "et.Application"];

    public static readonly string[] PptProgIds = ["PowerPoint.Application", "KWPP.Application"];

    public static JsonElement RequireArray(JsonElement p, string name)
    {
        if (p.ValueKind != JsonValueKind.Object
            || !p.TryGetProperty(name, out var value)
            || value.ValueKind != JsonValueKind.Array)
            throw new NativeUsageException($"缺少参数 {name}（JSON 数组）");
        return value;
    }

    public static bool PathsEqual(string? left, string? right)
    {
        if (string.IsNullOrEmpty(left) || string.IsNullOrEmpty(right))
            return false;
        try
        {
            return string.Equals(
                Path.GetFullPath(left).Replace('/', '\\'),
                Path.GetFullPath(right).Replace('/', '\\'),
                StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
        }
    }

    /// <summary>HEX（RRGGBB / AARRGGBB，可带 #）转 Office COM COLORREF（BGR）。</summary>
    public static int HexToBgr(string? value, string label = "颜色")
    {
        var text = (value ?? "").Trim().TrimStart('#');
        if (text.Length == 8)
            text = text[2..];
        if (text.Length != 6
            || !uint.TryParse(text, NumberStyles.HexNumber, CultureInfo.InvariantCulture,
                out var rgb))
            throw new NativeRuntimeException($"无效的{label}: '{value}'");
        var red = (int)((rgb >> 16) & 0xFF);
        var green = (int)((rgb >> 8) & 0xFF);
        var blue = (int)(rgb & 0xFF);
        return red + (green << 8) + (blue << 16);
    }

    public static Dictionary<string, object?> EditResult(
        string filePath, int operationCount, string message) => new()
    {
        ["success"] = true,
        ["file_path"] = filePath,
        ["backup_path"] = null,
        ["operation_count"] = operationCount,
        ["message"] = message,
    };

    /// <summary>拒绝修改以只读方式打开的演示文稿（与 Word/Excel 活动文档一致）。</summary>
    public static void EnsureWritablePresentation(dynamic presentation)
    {
        try
        {
            if (Convert.ToInt32(presentation.ReadOnly) != 0)
                throw new NativeRuntimeException(
                    "PowerPoint 演示文稿以只读方式打开，无法修改；"
                    + "请取消只读并重新打开后再试");
        }
        catch (NativeRuntimeException)
        {
            throw;
        }
        catch
        {
            // 个别 WPS 版本不暴露 ReadOnly；无法判断时继续由 Save 失败兜底。
        }
    }

    /// <summary>保存 PPT；Save 失败时尝试 SaveAs 同路径，仍失败则提示用户手工保存。</summary>
    public static void SavePresentation(dynamic presentation)    {
        try
        {
            presentation.Save();
        }
        catch (Exception saveError)
        {
            try
            {
                presentation.SaveAs(presentation.FullName);
            }
            catch (Exception saveAsError)
            {
                throw new NativeRuntimeException(
                    "PowerPoint 保存被拒绝（文件可能被占用），请在 PowerPoint 中按 Ctrl+S 手工保存: "
                    + $"{saveError.Message}; SaveAs: {saveAsError.Message}");
            }
        }
    }

    public static dynamic? FindWord(
        string filePath, out dynamic? document, out string progId)
    {
        foreach (var candidate in WordProgIds)
        {
            var app = GetActive(candidate);
            if (app is null)
                continue;
            try
            {
                dynamic documents = app.Documents;
                var count = (int)documents.Count;
                for (var index = 1; index <= count; index++)
                {
                    dynamic current = documents.Item(index);
                    if (PathsEqual(SafeFullName(current), filePath))
                    {
                        document = current;
                        progId = candidate;
                        return app;
                    }
                    NativeCommon.ReleaseCom(current);
                }
            }
            catch { /* 应用可能正在关闭 */ }
            NativeCommon.ReleaseCom(app);
        }
        document = null;
        progId = string.Empty;
        return null;
    }

    public static dynamic? FindExcel(
        string filePath, out dynamic? workbook, out string progId)
    {
        foreach (var candidate in ExcelProgIds)
        {
            var app = GetActive(candidate);
            if (app is null)
                continue;
            try
            {
                dynamic workbooks = app.Workbooks;
                var count = (int)workbooks.Count;
                for (var index = 1; index <= count; index++)
                {
                    dynamic current = workbooks.Item(index);
                    if (PathsEqual(SafeFullName(current), filePath))
                    {
                        workbook = current;
                        progId = candidate;
                        return app;
                    }
                    NativeCommon.ReleaseCom(current);
                }
            }
            catch { /* 应用可能正在关闭 */ }
            NativeCommon.ReleaseCom(app);
        }
        workbook = null;
        progId = string.Empty;
        return null;
    }

    public static dynamic? FindPowerPoint(
        string filePath, out dynamic? presentation, out string progId)
    {
        foreach (var candidate in PptProgIds)
        {
            var app = GetActive(candidate);
            if (app is null)
                continue;
            try
            {
                dynamic presentations = app.Presentations;
                var count = (int)presentations.Count;
                for (var index = 1; index <= count; index++)
                {
                    dynamic current = presentations.Item(index);
                    if (PathsEqual(SafeFullName(current), filePath))
                    {
                        presentation = current;
                        progId = candidate;
                        return app;
                    }
                    NativeCommon.ReleaseCom(current);
                }
            }
            catch { /* 应用可能正在关闭 */ }
            NativeCommon.ReleaseCom(app);
        }
        presentation = null;
        progId = string.Empty;
        return null;
    }

    private static dynamic? GetActive(string progId)
    {
        try { return GenericOpenCommands.GetActive(progId); }
        catch { return null; }
    }

    private static string? SafeFullName(dynamic item)
    {
        try { return (string)item.FullName; }
        catch { return null; }
    }
}
