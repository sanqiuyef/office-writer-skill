using System.Diagnostics;
using System.Text.Json;

namespace OfficeNative;

internal static class RecalcCommands
{
    public static Dictionary<string, object?> RunRecalculate(JsonElement p)
    {
        var backend = (Environment.GetEnvironmentVariable("OFFICE_CONVERT_BACKEND") ?? "auto")
            .Trim().ToLowerInvariant();
        if (backend == "none")
            throw new NativeRuntimeException(NativeCommon.BackendUnavailableHint);
        if (backend is not ("auto" or "dsoffice" or "com"))
            throw new NativeRuntimeException(
                "OFFICE_CONVERT_BACKEND 必须是 auto/dsoffice/com/none");
        var inputPath = Path.GetFullPath(NativeCommon.RequireString(p, "input_path"));
        NativeCommon.EnsureFile(inputPath, ".xlsx");
        var outputPath = NativeCommon.OptString(p, "output_path")
                         ?? Path.Combine(Path.GetDirectoryName(inputPath)!,
                             $"{Path.GetFileNameWithoutExtension(inputPath)}_recalc{Path.GetExtension(inputPath)}");
        outputPath = Path.GetFullPath(outputPath);
        NativeCommon.EnsureOutput(
            outputPath,
            NativeCommon.OptBool(p, "overwrite", false),
            NativeCommon.OptBool(p, "confirm_overwrite", false));
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
        if (!string.Equals(inputPath, outputPath, StringComparison.OrdinalIgnoreCase))
            File.Copy(inputPath, outputPath, overwrite: true);

        var before = Process.GetProcessesByName("EXCEL").Select(process => process.Id).ToHashSet();
        var started = before.Count == 0;
        dynamic? app = null;
        dynamic? workbook = null;
        try
        {
            var type = Type.GetTypeFromProgID("Excel.Application")
                       ?? throw new NativeRuntimeException(NativeCommon.BackendUnavailableHint);
            app = Activator.CreateInstance(type)
                  ?? throw new NativeRuntimeException("无法创建 Excel 实例");
            try { app.Visible = false; } catch { }
            try { app.DisplayAlerts = false; } catch { }
            workbook = app.Workbooks.Open(outputPath);
            string calculation;
            try
            {
                app.CalculateFullRebuild();
                calculation = "CalculateFullRebuild";
            }
            catch
            {
                app.Calculate();
                calculation = "Calculate";
            }
            workbook.Save();
            workbook.Close(false);
            workbook = null;
            return new Dictionary<string, object?>
            {
                ["success"] = true,
                ["file_path"] = outputPath,
                ["backend"] = "com",
                ["calculation"] = calculation,
                ["message"] = $"公式重算完成（COM: {calculation}）: {outputPath}",
            };
        }
        finally
        {
            if (workbook is not null) { try { workbook.Close(false); } catch { } }
            if (app is not null)
            {
                if (started) { try { app.Quit(); } catch { } }
                NativeCommon.ReleaseCom(workbook);
                NativeCommon.ReleaseCom(app);
                if (started)
                {
                    Thread.Sleep(800);
                    foreach (var process in Process.GetProcessesByName("EXCEL"))
                    {
                        if (before.Contains(process.Id))
                            continue;
                        try { process.Kill(); } catch { }
                    }
                }
            }
        }
    }
}
