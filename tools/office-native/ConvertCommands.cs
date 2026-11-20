using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace OfficeNative;

internal static class ConvertCommands
{
    private static readonly HashSet<string> InputExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".docx", ".doc", ".rtf", ".xlsx", ".xls", ".pptx", ".ppt",
    };

    public static Dictionary<string, object?> RunConvert(JsonElement p)
    {
        var inputPath = Path.GetFullPath(NativeCommon.RequireString(p, "input_path"));
        NativeCommon.EnsureFile(inputPath, InputExtensions.ToArray());
        var outputPath = NativeCommon.OptString(p, "output_path")
                         ?? Path.ChangeExtension(inputPath, ".pdf");
        outputPath = Path.GetFullPath(outputPath);
        if (!string.Equals(Path.GetExtension(outputPath), ".pdf", StringComparison.OrdinalIgnoreCase))
            throw new NativeUsageException("输出必须是 .pdf");
        outputPath = ResolveOutput(outputPath, p);
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

        var mode = (Environment.GetEnvironmentVariable("OFFICE_CONVERT_BACKEND") ?? "auto")
            .Trim().ToLowerInvariant();
        if (mode is not ("auto" or "dsoffice" or "com" or "none"))
            throw new NativeRuntimeException(
                "OFFICE_CONVERT_BACKEND 必须是 auto/dsoffice/com/none");
        if (mode == "none")
            throw new NativeRuntimeException(NativeCommon.BackendUnavailableHint);

        var fallback = new List<Dictionary<string, object?>>();
        if (mode is "auto" or "dsoffice")
        {
            var (cli, cliError) = FindDsoffice();
            if (cli is not null)
            {
                try
                {
                    return ConvertWithDsoffice(
                        cli, inputPath, outputPath, NativeCommon.OptString(p, "sheet"));
                }
                catch (Exception ex) when (mode == "auto")
                {
                    fallback.Add(new Dictionary<string, object?>
                    {
                        ["backend"] = "dsoffice",
                        ["error"] = ex.Message,
                    });
                }
            }
            else
            {
                var error = cliError
                    ?? "未找到 dsoffice CLI；设置 MCP_OFFICE_DSOFFICE 或安装 dsoffice";
                if (mode == "dsoffice")
                    throw new NativeRuntimeException(NativeCommon.DsofficeInstallHint);
                fallback.Add(new Dictionary<string, object?>
                {
                    ["backend"] = "dsoffice",
                    ["error"] = error,
                });
            }
        }

        var result = ConvertWithCom(inputPath, outputPath, NativeCommon.OptString(p, "sheet"));
        if (fallback.Count > 0)
            result["fallback_from"] = fallback;
        return result;
    }

    private static string ResolveOutput(string outputPath, JsonElement p)
    {
        var ifExists = NativeCommon.OptString(p, "if_exists");
        var overwrite = NativeCommon.OptBool(p, "overwrite", false)
                        || string.Equals(ifExists, "overwrite", StringComparison.OrdinalIgnoreCase);
        var confirm = NativeCommon.OptBool(p, "confirm_overwrite", false);
        var rename = NativeCommon.OptBool(p, "rename", false)
                     || string.Equals(ifExists, "rename", StringComparison.OrdinalIgnoreCase);
        if (rename && File.Exists(outputPath) && new FileInfo(outputPath).Length > 0)
        {
            var directory = Path.GetDirectoryName(outputPath)!;
            var stem = Path.GetFileNameWithoutExtension(outputPath);
            var extension = Path.GetExtension(outputPath);
            var counter = 1;
            string candidate;
            do
            {
                candidate = Path.Combine(directory, $"{stem} ({counter++}){extension}");
            } while (File.Exists(candidate));
            return candidate;
        }
        NativeCommon.EnsureOutput(outputPath, overwrite, confirm);
        return outputPath;
    }

    private static (string? Path, string? Error) FindDsoffice()
    {
        foreach (var name in new[] { "MCP_OFFICE_DSOFFICE", "MODULE_OFFICE_DSOFFICE" })
        {
            var value = Environment.GetEnvironmentVariable(name);
            if (!string.IsNullOrWhiteSpace(value))
                // 显式指定但路径无效时不再回退 PATH，避免误用全局 dsoffice。
                return File.Exists(value)
                    ? (value, null)
                    : (null, $"{name} 指定的 dsoffice 不存在: {value}");
        }
        var path = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        foreach (var directory in path.Split(Path.PathSeparator))
        {
            if (string.IsNullOrWhiteSpace(directory))
                continue;
            foreach (var file in new[] { "dsoffice.exe", "dsoffice.cmd", "dsoffice" })
            {
                var candidate = Path.Combine(directory, file);
                if (File.Exists(candidate))
                    return (candidate, null);
            }
        }
        return (null, null);
    }

    private static Dictionary<string, object?> ConvertWithDsoffice(
        string cli, string inputPath, string outputPath, string? sheet)
    {
        var temp = TempPath(outputPath);
        try
        {
            var start = new ProcessStartInfo
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
            };
            // npm installs dsoffice as a .cmd shim on Windows; CreateProcess cannot
            // start it directly, so route it through cmd.exe.
            if (cli.EndsWith(".cmd", StringComparison.OrdinalIgnoreCase)
                || cli.EndsWith(".bat", StringComparison.OrdinalIgnoreCase))
            {
                start.FileName =
                    Environment.GetEnvironmentVariable("COMSPEC") ?? "cmd.exe";
                start.ArgumentList.Add("/d");
                start.ArgumentList.Add("/c");
                start.ArgumentList.Add(cli);
            }
            else
            {
                start.FileName = cli;
            }
            start.ArgumentList.Add("convert");
            start.ArgumentList.Add("--input");
            start.ArgumentList.Add(inputPath);
            start.ArgumentList.Add("--output");
            start.ArgumentList.Add(temp);
            if (!string.IsNullOrWhiteSpace(sheet))
            {
                start.ArgumentList.Add("--sheet");
                start.ArgumentList.Add(sheet);
            }
            using var process = Process.Start(start)
                ?? throw new NativeRuntimeException("dsoffice 无法启动");
            var stdoutTask = process.StandardOutput.ReadToEndAsync();
            var stderrTask = process.StandardError.ReadToEndAsync();
            if (!process.WaitForExit(300_000))
            {
                try { process.Kill(true); } catch { }
                throw new NativeRuntimeException("dsoffice 转换超时（>300s）");
            }
            var stdout = stdoutTask.GetAwaiter().GetResult();
            var stderr = stderrTask.GetAwaiter().GetResult();
            if (process.ExitCode != 0 || !File.Exists(temp) || new FileInfo(temp).Length == 0)
                throw new NativeRuntimeException(
                    $"dsoffice 转换失败（exit={process.ExitCode}）: {Trim(stdout + stderr)}");
            var payload = ParseDsofficePayload(stdout);
            var data = new Dictionary<string, object?>
            {
                ["success"] = true,
                ["file_path"] = outputPath,
                ["backend"] = payload.GetValueOrDefault("backend") ?? "dsoffice",
                ["backend_kind"] = "dsoffice",
                ["missing_fonts"] = payload.GetValueOrDefault("missingFonts") ?? new List<string>(),
                ["font_substitutions"] = new List<object>(),
                ["message"] = "转换成功（dsoffice）",
            };
            Publish(temp, outputPath, data);
            return data;
        }
        finally
        {
            if (File.Exists(temp))
            {
                try { File.Delete(temp); } catch { }
            }
        }
    }

    private static Dictionary<string, object?> ParseDsofficePayload(string stdout)
    {
        try
        {
            using var document = JsonDocument.Parse(stdout);
            if (document.RootElement.ValueKind != JsonValueKind.Object)
                return new Dictionary<string, object?>();
            var result = new Dictionary<string, object?>();
            if (document.RootElement.TryGetProperty("backend", out var backend))
                result["backend"] = backend.ToString();
            if (document.RootElement.TryGetProperty("missingFonts", out var fonts))
                result["missingFonts"] = fonts.EnumerateArray()
                    .Select(item => item.ToString()).ToList();
            return result;
        }
        catch (JsonException)
        {
            return new Dictionary<string, object?>();
        }
    }

    private static Dictionary<string, object?> ConvertWithCom(
        string inputPath, string outputPath, string? sheet)
    {
        var extension = Path.GetExtension(inputPath).ToLowerInvariant();
        return extension switch
        {
            ".docx" or ".doc" or ".rtf" => ConvertWord(inputPath, outputPath),
            ".xlsx" or ".xls" => ConvertExcel(inputPath, outputPath, sheet),
            ".pptx" or ".ppt" => ConvertPowerPoint(inputPath, outputPath),
            _ => throw new NativeRuntimeException($"不支持的输入格式: {extension}"),
        };
    }

    private static Dictionary<string, object?> ConvertWord(string inputPath, string outputPath)
    {
        var temp = TempPath(outputPath);
        var before = Snapshot("WINWORD");
        dynamic? app = null;
        dynamic? document = null;
        HashSet<int> owned = [];
        try
        {
            app = IsolatedCom.Create("Word.Application", "Word 未注册");
            owned = NewPids("WINWORD", before);
            TrySet(app, "Visible", false);
            TrySet(app, "DisplayAlerts", 0);
            document = app.Documents.Open(inputPath, false, true, false);
            document.ExportAsFixedFormat(temp, 17); // wdExportFormatPDF
            document.Close(false);
            document = null;
            var data = ComResult("Word.Application");
            Publish(temp, outputPath, data);
            return data;
        }
        finally
        {
            Cleanup(document, app, "WINWORD", owned);
            DeleteIfExists(temp);
        }
    }

    private static Dictionary<string, object?> ConvertExcel(
        string inputPath, string outputPath, string? sheet)
    {
        var temp = TempPath(outputPath);
        var before = Snapshot("EXCEL");
        dynamic? app = null;
        dynamic? workbook = null;
        HashSet<int> owned = [];
        try
        {
            app = IsolatedCom.Create("Excel.Application", "Excel 未注册");
            owned = NewPids("EXCEL", before);
            TrySet(app, "Visible", false);
            TrySet(app, "DisplayAlerts", false);
            workbook = app.Workbooks.Open(inputPath, 0, true); // UpdateLinks=0, ReadOnly=true
            if (string.IsNullOrWhiteSpace(sheet))
                workbook.ExportAsFixedFormat(0, temp);
            else
                workbook.Worksheets.Item(sheet).ExportAsFixedFormat(0, temp);
            workbook.Close(false);
            workbook = null;
            var data = ComResult("Excel.Application");
            Publish(temp, outputPath, data);
            return data;
        }
        finally
        {
            Cleanup(workbook, app, "EXCEL", owned);
            DeleteIfExists(temp);
        }
    }

    private static Dictionary<string, object?> ConvertPowerPoint(
        string inputPath, string outputPath)
    {
        var temp = TempPath(outputPath);
        var before = Snapshot("POWERPNT");
        dynamic? app = null;
        dynamic? presentation = null;
        HashSet<int> owned = [];
        try
        {
            app = IsolatedCom.Create("PowerPoint.Application", "PowerPoint 未注册");
            owned = NewPids("POWERPNT", before);
            TrySet(app, "Visible", -1);
            presentation = app.Presentations.Open(inputPath, -1, 0, 0); // ReadOnly=true
            presentation.SaveAs(temp, 32); // ppSaveAsPDF
            presentation.Close();
            presentation = null;
            var data = ComResult("PowerPoint.Application");
            Publish(temp, outputPath, data);
            return data;
        }
        finally
        {
            Cleanup(presentation, app, "POWERPNT", owned);
            DeleteIfExists(temp);
        }
    }

    private static Dictionary<string, object?> ComResult(string progId) => new()
    {
        ["success"] = true,
        ["file_path"] = string.Empty,
        ["backend"] = "com",
        ["backend_kind"] = "com",
        ["com_engine"] = "office",
        ["com_prog_id"] = progId,
        ["com_application"] = progId,
        ["missing_fonts"] = Array.Empty<string>(),
        ["font_substitutions"] = Array.Empty<object>(),
        ["message"] = $"转换成功（COM: {progId}）",
    };

    private static void Publish(string temp, string outputPath, Dictionary<string, object?> data)
    {
        if (!File.Exists(temp) || new FileInfo(temp).Length == 0)
            throw new NativeRuntimeException($"转换未生成有效 PDF: {temp}");
        string? backup = null;
        if (File.Exists(outputPath) && new FileInfo(outputPath).Length > 0)
            backup = NativeCommon.Backup(outputPath);
        File.Move(temp, outputPath, overwrite: true);
        data["file_path"] = outputPath;
        if (backup is not null)
            data["backup_path"] = backup;
    }

    private static string TempPath(string outputPath) =>
        Path.Combine(Path.GetDirectoryName(outputPath)!,
            $".convert-{Guid.NewGuid():N}.pdf");

    private static void DeleteIfExists(string path)
    {
        if (File.Exists(path))
        {
            try { File.Delete(path); } catch { }
        }
    }

    private static void Cleanup(dynamic? document, dynamic? app, string processName,
        HashSet<int> owned)
    {
        if (document is not null) { try { document.Close(false); } catch { } }
        if (app is not null)
        {
            try { app.Quit(); } catch { }
            NativeCommon.ReleaseCom(document);
            NativeCommon.ReleaseCom(app);
        }
        foreach (var pid in owned)
        {
            try
            {
                using var process = Process.GetProcessById(pid);
                if (!process.HasExited)
                    process.Kill();
            }
            catch { /* already exited */ }
        }
    }

    private static void TrySet(dynamic app, string property, object value)
    {
        try
        {
            var target = (object)app;
            target.GetType().InvokeMember(property,
                System.Reflection.BindingFlags.SetProperty, null, target, new[] { value });
        }
        catch { }
    }

    private static HashSet<int> Snapshot(string processName) =>
        Process.GetProcessesByName(processName).Select(process => process.Id).ToHashSet();

    private static HashSet<int> NewPids(string processName, HashSet<int> before) =>
        Process.GetProcessesByName(processName)
            .Select(process => process.Id)
            .Where(pid => !before.Contains(pid))
            .ToHashSet();

    private static string Trim(string value) =>
        value.Trim()[..Math.Min(300, value.Trim().Length)];

    /// <summary>DispatchEx 语义：通过 IClassFactory 新建独立 Office 实例，不附着用户实例。</summary>
    private static class IsolatedCom
    {
        private const uint ClsCtxLocalServer = 4;
        private static readonly Guid IidClassFactory = new("00000001-0000-0000-C000-000000000046");
        private static readonly Guid IidDispatch = new("00020400-0000-0000-C000-000000000046");

        public static dynamic Create(string progId, string message)
        {
            if (CLSIDFromProgID(progId, out var clsid) != 0)
                throw new NativeRuntimeException(message);
            var factoryIid = IidClassFactory;
            if (CoGetClassObject(ref clsid, ClsCtxLocalServer, IntPtr.Zero,
                    ref factoryIid, out var factoryPointer) != 0)
                throw new NativeRuntimeException($"无法创建隔离的 {progId} 实例");
            object? factory = null;
            try
            {
                factory = Marshal.GetObjectForIUnknown(factoryPointer);
                if (factory is not IClassFactory classFactory)
                    throw new NativeRuntimeException($"无法获取 {progId} 的 IClassFactory");
                var dispatchIid = IidDispatch;
                if (classFactory.CreateInstance(null, ref dispatchIid, out var instance) != 0)
                    throw new NativeRuntimeException($"无法创建隔离的 {progId} 实例");
                return instance;
            }
            finally
            {
                if (factory is not null)
                    Marshal.ReleaseComObject(factory);
                Marshal.Release(factoryPointer);
            }
        }

        [ComImport]
        [Guid("00000001-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IClassFactory
        {
            [PreserveSig]
            int CreateInstance(
                [MarshalAs(UnmanagedType.IUnknown)] object? outer,
                ref Guid riid,
                [MarshalAs(UnmanagedType.IUnknown)] out object instance);

            [PreserveSig]
            int LockServer(bool lockServer);
        }

        [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
        private static extern int CLSIDFromProgID(string progId, out Guid clsid);

        [DllImport("ole32.dll")]
        private static extern int CoGetClassObject(
            ref Guid clsid, uint context, IntPtr reserved,
            ref Guid riid, out IntPtr classObject);
    }
}
