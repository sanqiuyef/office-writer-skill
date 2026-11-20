using System.Text.Json;

namespace OfficeNative;

/// <summary>
/// edit-open-excel-workbook：向已在 Excel/WPS 中打开的工作簿追加工作表并保存。
/// 对应 Python office_com.try_edit_open_excel_workbook。
/// </summary>
internal static class OpenEditExcelCommands
{
    private static readonly char[] InvalidNameChars = ['[', ']', ':', '*', '?', '/', '\\'];

    public static Dictionary<string, object?> RunEditOpenExcelWorkbook(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        var sheets = ExcelModelParser.ParseAppendSheets(p);
        NativeCommon.EnsureFileLight(filePath);

        var app = OpenEditCommon.FindExcel(filePath, out var workbook, out var progId)
            ?? throw new NativeRuntimeException(
                $"目标文件未在活动 Excel/WPS 实例中打开: {filePath}；请先打开文件");
        var created = new List<dynamic>();
        try
        {
            if (workbook is null)
                throw new NativeRuntimeException(
                    $"目标文件未在活动 Excel/WPS 实例中打开: {filePath}");
            if ((bool)workbook.ReadOnly)
                throw new NativeRuntimeException("活动 Excel 工作簿为只读，无法保存");

            ValidateSheets(workbook, sheets);

            try
            {
                foreach (var sheet in sheets)
                    WriteSheet(workbook, sheet, created);
                workbook.Save();
            }
            catch (Exception error)
            {
                try
                {
                    RollbackSheets(workbook, created);
                }
                catch (Exception rollbackError)
                {
                    throw new NativeRuntimeException(
                        $"活动 Excel 工作簿编辑失败: {error.Message}；"
                        + $"回滚失败，请勿保存该工作簿: {rollbackError.Message}");
                }
                throw new NativeRuntimeException(
                    $"活动 Excel 工作簿编辑失败，已回滚本次新增工作表: {error.Message}");
            }

            return OpenEditCommon.EditResult(filePath, sheets.Count,
                $"已通过 {progId} 修改并保存活动 Excel 工作簿");
        }
        finally
        {
            foreach (var sheet in created)
                NativeCommon.ReleaseCom(sheet);
            NativeCommon.ReleaseCom(workbook);
            NativeCommon.ReleaseCom(app);
        }
    }

    // ── 校验 ────────────────────────────────────────────────────

    private static void ValidateSheets(dynamic workbook, List<ExcelSheetDef> sheets)
    {
        var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        dynamic worksheets = workbook.Worksheets;
        var count = (int)worksheets.Count;
        for (var index = 1; index <= count; index++)
        {
            dynamic worksheet = worksheets.Item(index);
            try { existing.Add((string)worksheet.Name); }
            finally { NativeCommon.ReleaseCom(worksheet); }
        }

        var planned = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var sheet in sheets)
        {
            var name = TruncateName(sheet.Name);
            if (name.Length == 0 || name.IndexOfAny(InvalidNameChars) >= 0)
                throw new NativeRuntimeException($"无效的工作表名称: '{sheet.Name}'");
            if (existing.Contains(name) || !planned.Add(name))
                throw new NativeRuntimeException($"工作表名称已存在: '{name}'");
            if (HasComUnsupportedFields(sheet))
                throw new NativeRuntimeException(
                    $"工作表 '{name}' 使用了活动 Excel 追加暂不支持的扩展字段"
                    + "（图表/条件格式/数据验证/超级表/名称/批注/图片/页面设置/行高/隐藏/边框等）；"
                    + "请关闭文件后使用文件级 edit_excel_workbook");
        }
    }

    private static bool HasComUnsupportedFields(ExcelSheetDef sheet)
    {
        if (NonEmpty(sheet.RowHeights)
            || NonEmpty(sheet.Charts)
            || NonEmpty(sheet.ConditionalFormats)
            || NonEmpty(sheet.DataValidations)
            || NonEmpty(sheet.Tables)
            || NonEmpty(sheet.DefinedNames)
            || NonEmpty(sheet.Hyperlinks)
            || NonEmpty(sheet.Comments)
            || NonEmpty(sheet.Images)
            || sheet.PageSetup is not null)
            return true;
        if (!string.Equals(sheet.SheetState, "visible", StringComparison.Ordinal))
            return true;
        if (sheet.HeaderStyle is { } header
            && (header.Italic is not null
                || header.VerticalAlignment is not null
                || header.Border is not null))
            return true;
        return sheet.ColumnFormats?.Any(column =>
            column.Width is not null
            || column.FontSize is not null
            || column.Bold is not null
            || column.Italic is not null
            || column.WrapText is not null
            || column.VerticalAlignment is not null) == true;
    }

    private static bool NonEmpty<T>(IReadOnlyCollection<T>? values) =>
        values is { Count: > 0 };

    private static string TruncateName(string name) =>
        name.Length <= 31 ? name : name[..31];

    // ── 写入 ────────────────────────────────────────────────────

    private static void WriteSheet(
        dynamic workbook, ExcelSheetDef sheet, List<dynamic> created)
    {
        dynamic worksheets = workbook.Worksheets;
        dynamic worksheet = worksheets.Add(
            Type.Missing,
            worksheets.Item(worksheets.Count),
            Type.Missing,
            Type.Missing);
        created.Add(worksheet);
        worksheet.Name = TruncateName(sheet.Name);

        try
        {
            var hasHeaders = sheet.Headers is { Count: > 0 };
            if (hasHeaders)
            {
                for (var columnIndex = 1; columnIndex <= sheet.Headers!.Count; columnIndex++)
                {
                    dynamic cell = worksheet.Cells(1, columnIndex);
                    try
                    {
                        cell.Value = sheet.Headers[columnIndex - 1];
                        ApplyExcelStyle(cell, sheet.HeaderStyle, defaultBold: true);
                    }
                    finally
                    {
                        NativeCommon.ReleaseCom(cell);
                    }
                }
            }

            var startRow = hasHeaders ? 2 : 1;
            for (var rowOffset = 0; rowOffset < sheet.Data.Count; rowOffset++)
            {
                var row = sheet.Data[rowOffset];
                for (var columnIndex = 1; columnIndex <= row.Count; columnIndex++)
                {
                    dynamic cell = worksheet.Cells(startRow + rowOffset, columnIndex);
                    try
                    {
                        SetCellValue(cell, row[columnIndex - 1], sheet.AllowFormulas);
                        if (sheet.ColumnFormats is { } formats
                            && columnIndex <= formats.Count)
                            ApplyExcelStyle(cell, formats[columnIndex - 1].ToCellStyle());
                    }
                    finally
                    {
                        NativeCommon.ReleaseCom(cell);
                    }
                }
            }

            if (sheet.ColumnWidths is { } widths)
            {
                for (var columnIndex = 1; columnIndex <= widths.Count; columnIndex++)
                {
                    dynamic column = worksheet.Columns(columnIndex);
                    try { column.ColumnWidth = widths[columnIndex - 1]; }
                    finally { NativeCommon.ReleaseCom(column); }
                }
            }

            if (sheet.MergedCells is { Count: > 0 })
            {
                foreach (var merged in sheet.MergedCells)
                {
                    dynamic range = worksheet.Range(merged);
                    try { range.Merge(); }
                    finally { NativeCommon.ReleaseCom(range); }
                }
            }

            if (sheet.AutoFilter && hasHeaders && sheet.Data.Count > 0)
            {
                var lastColumn = sheet.Headers!.Count;
                foreach (var row in sheet.Data)
                    lastColumn = Math.Max(lastColumn, row.Count);
                dynamic range = worksheet.Range(
                    worksheet.Cells(1, 1),
                    worksheet.Cells(sheet.Data.Count + 1, lastColumn));
                try { range.AutoFilter(); }
                finally { NativeCommon.ReleaseCom(range); }
            }

            if (!string.IsNullOrEmpty(sheet.TabColor))
                worksheet.Tab.Color = OpenEditCommon.HexToBgr(sheet.TabColor, "工作表标签颜色");

            if (!string.IsNullOrEmpty(sheet.Freeze))
                ApplyFreeze(workbook, worksheet, sheet.Freeze!);
        }
        catch
        {
            // 已登记到 created，由上层统一回滚
            throw;
        }
    }

    private static void ApplyFreeze(dynamic workbook, dynamic worksheet, string freezeRange)
    {
        dynamic application = workbook.Application;
        dynamic previousSheet = application.ActiveSheet;
        dynamic previousSelection = application.Selection;
        try
        {
            worksheet.Activate();
            application.ActiveWindow.FreezePanes = false;
            dynamic range = worksheet.Range(freezeRange);
            try { range.Select(); }
            finally { NativeCommon.ReleaseCom(range); }
            application.ActiveWindow.FreezePanes = true;
        }
        finally
        {
            try
            {
                previousSheet.Activate();
                previousSelection.Select();
            }
            catch { /* 恢复原视图失败不影响写入结果 */ }
            NativeCommon.ReleaseCom(previousSelection);
            NativeCommon.ReleaseCom(previousSheet);
            NativeCommon.ReleaseCom(application);
        }
    }

    private static void SetCellValue(dynamic cell, JsonElement value, bool allowFormulas)
    {
        switch (value.ValueKind)
        {
            case JsonValueKind.String:
                var text = value.GetString() ?? "";
                cell.Value = text.StartsWith('=') && !allowFormulas ? "'" + text : text;
                break;
            case JsonValueKind.Number:
                if (value.TryGetInt32(out var integer))
                    cell.Value = integer;
                else if (value.TryGetDouble(out var number))
                    cell.Value = number;
                break;
            case JsonValueKind.True:
                cell.Value = true;
                break;
            case JsonValueKind.False:
                cell.Value = false;
                break;
            case JsonValueKind.Null:
            case JsonValueKind.Undefined:
                break;
            default:
                cell.Value = value.GetRawText();
                break;
        }
    }

    private static void ApplyExcelStyle(
        dynamic cell, ExcelCellStyle? style, bool defaultBold = false)
    {
        if (style is null)
        {
            if (defaultBold)
                cell.Font.Bold = true;
            return;
        }
        if (!string.IsNullOrEmpty(style.FontName))
            cell.Font.Name = style.FontName;
        if (style.FontSize is not null)
            cell.Font.Size = style.FontSize.Value;
        if (style.Bold is not null)
            cell.Font.Bold = style.Bold.Value;
        else if (defaultBold)
            cell.Font.Bold = true;
        if (!string.IsNullOrEmpty(style.FontColor))
            cell.Font.Color = OpenEditCommon.HexToBgr(style.FontColor, "字体颜色");
        if (!string.IsNullOrEmpty(style.BgColor))
            cell.Interior.Color = OpenEditCommon.HexToBgr(style.BgColor, "背景颜色");
        if (!string.IsNullOrEmpty(style.NumberFormat))
            cell.NumberFormat = style.NumberFormat;
        if (!string.IsNullOrEmpty(style.Alignment))
        {
            cell.HorizontalAlignment = style.Alignment switch
            {
                "left" => -4131,
                "center" => -4108,
                "right" => -4152,
                _ => -4131,
            };
        }
        if (style.WrapText is not null)
            cell.WrapText = style.WrapText.Value;
    }

    private static void RollbackSheets(dynamic workbook, List<dynamic> worksheets)
    {
        if (worksheets.Count == 0)
            return;
        dynamic application = workbook.Application;
        var previousAlerts = application.DisplayAlerts;
        try
        {
            application.DisplayAlerts = false;
            for (var index = worksheets.Count - 1; index >= 0; index--)
                worksheets[index].Delete();
        }
        finally
        {
            application.DisplayAlerts = previousAlerts;
            NativeCommon.ReleaseCom(application);
        }
    }
}
