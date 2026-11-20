using System.Diagnostics;
using System.Text.Json;

namespace OfficeNative;

internal static class ExcelCommands
{
    private static readonly Dictionary<string, int> PivotFunctions = new(StringComparer.OrdinalIgnoreCase)
    {
        ["sum"] = -4157,
        ["count"] = -4112,
        ["average"] = -4106,
        ["max"] = -4136,
        ["min"] = -4139,
        ["product"] = -4149,
        ["stdev"] = -4155,
        ["var"] = -4164,
    };

    public static Dictionary<string, object?> RunPivots(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        var action = NativeCommon.OptString(p, "action") ?? "list";
        NativeCommon.EnsureFile(filePath, ".xlsx", ".xlsm");
        var backupPath = action is "create" or "refresh"
            ? NativeCommon.Backup(filePath)
            : null;
        var (app, workbook, before) = OpenWorkbook(filePath);
        try
        {
            var data = action switch
            {
                "create" => CreatePivot(workbook, p),
                "list" => ListPivots(workbook, p),
                "refresh" => RefreshPivots(workbook, p),
                _ => throw new NativeUsageException("action 必须是 create/list/refresh"),
            };
            if (action is "create" or "refresh")
            {
                workbook.Save();
                data["backup_path"] = backupPath;
            }
            data["file_path"] = Path.GetFullPath(filePath);
            return data;
        }
        finally
        {
            CloseWorkbook(app, workbook, before);
        }
    }

    private static Dictionary<string, object?> CreatePivot(dynamic workbook, JsonElement p)
    {
        var sourceSheet = NativeCommon.RequireString(p, "source_sheet");
        var sourceRange = NativeCommon.RequireString(p, "source_range");
        var targetSheet = NativeCommon.RequireString(p, "target_sheet");
        var targetCell = NativeCommon.OptString(p, "target_cell") ?? "A3";
        var name = NativeCommon.OptString(p, "name")
                   ?? $"Pivot_{DateTime.Now:HHmmss}";
        var rowField = NativeCommon.OptString(p, "row_field");
        var columnField = NativeCommon.OptString(p, "column_field");
        var dataField = NativeCommon.OptString(p, "data_field");
        var functionName = NativeCommon.OptString(p, "function") ?? "sum";
        if (!PivotFunctions.TryGetValue(functionName, out var function))
            throw new NativeUsageException(
                $"未知聚合函数: {functionName}（sum/count/average/max/min/product/stdev/var）");

        dynamic source = GetWorksheet(workbook, sourceSheet);
        dynamic target = GetWorksheet(workbook, targetSheet);
        dynamic cache = workbook.PivotCaches().Create(1, source.Range(sourceRange));
        dynamic pivot = cache.CreatePivotTable(target.Range(targetCell), name);
        if (rowField is not null)
            pivot.PivotFields(rowField).Orientation = 1; // xlRowField
        if (columnField is not null)
            pivot.PivotFields(columnField).Orientation = 2; // xlColumnField
        if (dataField is not null)
        {
            dynamic field = pivot.PivotFields(dataField);
            field.Orientation = 4; // xlDataField
            field.Function = function;
        }
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["action"] = "create",
            ["name"] = (string)pivot.Name,
            ["table_range"] = (string)pivot.TableRange1.Address,
            ["source_sheet"] = sourceSheet,
            ["source_range"] = sourceRange,
            ["target_sheet"] = targetSheet,
            ["target_cell"] = targetCell,
            ["row_field"] = rowField,
            ["column_field"] = columnField,
            ["data_field"] = dataField,
            ["function"] = functionName,
        };
    }

    private static Dictionary<string, object?> ListPivots(dynamic workbook, JsonElement p)
    {
        var sheetName = NativeCommon.OptString(p, "sheet");
        var results = new List<Dictionary<string, object?>>();
        if (sheetName is not null)
        {
            CollectPivots(GetWorksheet(workbook, sheetName), results);
        }
        else
        {
            foreach (dynamic sheet in workbook.Worksheets)
                CollectPivots(sheet, results);
        }
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["action"] = "list",
            ["count"] = results.Count,
            ["pivots"] = results,
        };
    }

    private static void CollectPivots(dynamic sheet, List<Dictionary<string, object?>> results)
    {
        dynamic pivots = sheet.PivotTables();
        for (var i = 1; i <= (int)pivots.Count; i++)
        {
            dynamic pivot = pivots.Item(i);
            results.Add(new Dictionary<string, object?>
            {
                ["sheet"] = (string)sheet.Name,
                ["name"] = (string)pivot.Name,
                ["table_range"] = (string)pivot.TableRange1.Address,
            });
        }
    }

    private static Dictionary<string, object?> RefreshPivots(dynamic workbook, JsonElement p)
    {
        var sheetName = NativeCommon.OptString(p, "sheet");
        var pivotName = NativeCommon.OptString(p, "name");
        var refreshed = new List<string>();
        if (sheetName is not null)
        {
            dynamic pivots = GetWorksheet(workbook, sheetName).PivotTables();
            for (var i = 1; i <= (int)pivots.Count; i++)
            {
                dynamic pivot = pivots.Item(i);
                if (pivotName is null || (string)pivot.Name == pivotName)
                {
                    pivot.RefreshTable();
                    refreshed.Add((string)pivot.Name);
                }
            }
        }
        else
        {
            workbook.RefreshAll();
            foreach (dynamic sheet in workbook.Worksheets)
            {
                dynamic pivots = sheet.PivotTables();
                for (var i = 1; i <= (int)pivots.Count; i++)
                    refreshed.Add((string)pivots.Item(i).Name);
            }
        }
        if (pivotName is not null && refreshed.Count == 0)
            throw new NativeRuntimeException($"找不到透视表: {pivotName}");
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["action"] = "refresh",
            ["refreshed"] = refreshed,
        };
    }

    public static Dictionary<string, object?> RunPrint(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        var action = NativeCommon.OptString(p, "action") ?? "pdf";
        NativeCommon.EnsureFile(filePath, ".xlsx", ".xlsm");
        var output = NativeCommon.OptString(p, "output");
        if (action == "pdf")
        {
            if (string.IsNullOrWhiteSpace(output))
                throw new NativeUsageException("action=pdf 必须提供 output");
            NativeCommon.EnsureOutput(
                output,
                NativeCommon.OptBool(p, "overwrite", false),
                NativeCommon.OptBool(p, "confirm_overwrite", false));
        }
        if (action == "print"
            && !NativeCommon.OptBool(p, "confirm_print", false))
        {
            throw new NativeUsageException(
                "action=print 是外部副作用操作，必须显式传 confirm_print=true");
        }

        var (app, workbook, before) = OpenWorkbook(filePath);
        try
        {
            var sheetName = NativeCommon.OptString(p, "sheet");
            dynamic target = sheetName is not null ? GetWorksheet(workbook, sheetName) : workbook;
            Dictionary<string, object?> data;
            if (action == "pdf")
            {
                Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(output!))!);
                target.ExportAsFixedFormat(0, Path.GetFullPath(output!));
                data = new Dictionary<string, object?>
                {
                    ["success"] = true,
                    ["action"] = "pdf",
                    ["output"] = Path.GetFullPath(output!),
                    ["bytes"] = new FileInfo(output!).Length,
                    ["sheet"] = sheetName,
                };
            }
            else if (action == "print")
            {
                var printer = NativeCommon.OptString(p, "printer");
                var copies = NativeCommon.OptInt(p, "copies", 1);
                if (printer is null)
                    workbook.PrintOut(Type.Missing, Type.Missing, copies, false);
                else
                    workbook.PrintOut(Type.Missing, Type.Missing, copies, false, printer);
                data = new Dictionary<string, object?>
                {
                    ["success"] = true,
                    ["action"] = "print",
                    ["copies"] = copies,
                    ["printer"] = printer ?? "(default)",
                    ["sheet"] = sheetName,
                };
            }
            else
            {
                throw new NativeUsageException("action 必须是 pdf/print");
            }
            data["file_path"] = Path.GetFullPath(filePath);
            return data;
        }
        finally
        {
            CloseWorkbook(app, workbook, before);
        }
    }

    private static dynamic GetWorksheet(dynamic workbook, string name)
    {
        try
        {
            return workbook.Worksheets.Item(name);
        }
        catch
        {
            throw new NativeRuntimeException($"工作表不存在: {name}");
        }
    }

    private static (dynamic App, dynamic Workbook, HashSet<int> Before) OpenWorkbook(string filePath)
    {
        var before = Process.GetProcessesByName("EXCEL")
            .Select(process => process.Id).ToHashSet();
        var type = Type.GetTypeFromProgID("Excel.Application")
                   ?? throw new NativeRuntimeException("Excel.Application 未注册");
        dynamic app = Activator.CreateInstance(type)!;
        try
        {
            app.Visible = false;
            app.DisplayAlerts = false;
            dynamic workbook = app.Workbooks.Open(Path.GetFullPath(filePath));
            return (app, workbook, before);
        }
        catch
        {
            try { app.Quit(); } catch { }
            NativeCommon.ReleaseCom(app);
            throw;
        }
    }

    private static void CloseWorkbook(
        dynamic app, dynamic workbook, HashSet<int> before)
    {
        try { workbook.Close(false); } catch { }
        try { app.Quit(); } catch { }
        NativeCommon.ReleaseCom(workbook);
        NativeCommon.ReleaseCom(app);
        // Excel 的 Quit 也可能滞后；只清理本次启动前不存在的进程。
        Thread.Sleep(800);
        foreach (var process in Process.GetProcessesByName("EXCEL"))
        {
            if (before.Contains(process.Id))
                continue;
            try { process.Kill(); } catch { /* already exited */ }
        }
    }
}
