using System.Buffers.Binary;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using PDFtoImage;

namespace OfficeNative;

/// <summary>
/// render-document-pages: PDF pages are rasterized directly with PDFium
/// (PDFtoImage, MIT); Office files are first converted through the existing
/// ConvertCommands backend (dsoffice or isolated Office COM) and then rendered.
/// </summary>
internal static class PdfFileRenderCommands
{
    private static readonly HashSet<string> InputExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt", ".rtf", ".pdf",
    };

    public static Dictionary<string, object?> RunRenderDocumentPages(JsonElement p)
    {
        var inputPath = Path.GetFullPath(NativeCommon.RequireString(p, "input_path"));
        NativeCommon.EnsureFile(inputPath, InputExtensions.ToArray());

        var pages = (NativeCommon.OptString(p, "pages") ?? "all").Trim();
        if (pages.Length == 0)
            pages = "all";
        if (!pages.Equals("all", StringComparison.OrdinalIgnoreCase)
            && !Regex.IsMatch(pages, @"^\d+(?:\s*,\s*\d+)*$"))
            throw new NativeUsageException("pages 必须是 'all' 或 '1,3,5' 样式");

        var dpi = ParseDpi(p);
        var sheet = NativeCommon.OptString(p, "sheet");
        var cellRange = NativeCommon.OptString(p, "cell_range");
        if (!string.IsNullOrWhiteSpace(cellRange) && string.IsNullOrWhiteSpace(sheet))
            throw new NativeUsageException("cell_range 需要与 sheet 同时指定");

        var extension = Path.GetExtension(inputPath).ToLowerInvariant();
        var backendMode = (Environment.GetEnvironmentVariable("OFFICE_CONVERT_BACKEND") ?? "auto")
            .Trim().ToLowerInvariant();
        if (backendMode is not ("auto" or "dsoffice" or "com" or "none"))
            throw new NativeRuntimeException(
                "OFFICE_CONVERT_BACKEND 必须是 auto/dsoffice/com/none");
        if (backendMode == "none" && extension != ".pdf")
            throw new NativeRuntimeException(NativeCommon.BackendUnavailableHint);
        var isExcel = extension is ".xlsx" or ".xls";
        if (isExcel
            && (string.IsNullOrWhiteSpace(sheet) || string.IsNullOrWhiteSpace(cellRange)))
            throw new NativeUsageException(
                "Excel 渲染需要同时指定 --sheet 与 --cell-range（如 --sheet 数据 --cell-range A1:D20）");

        var requestedDirectory = NativeCommon.OptString(p, "output_dir");
        if (string.IsNullOrWhiteSpace(requestedDirectory))
            requestedDirectory = Path.Combine(
                Path.GetDirectoryName(inputPath)!,
                $"{Path.GetFileNameWithoutExtension(inputPath)}_pages");
        var outputDirectory = Path.GetFullPath(requestedDirectory);

        var mode = NativeCommon.OptString(p, "if_exists") ?? "error";
        if (mode is not ("error" or "overwrite"))
            throw new NativeUsageException(
                $"if_exists 必须是 error/overwrite 之一，收到 {mode}");
        var confirm = NativeCommon.OptBool(p, "confirm_overwrite", false);
        GuardDirectoryTarget(outputDirectory, mode, confirm);

        var parent = Path.GetDirectoryName(outputDirectory)!;
        Directory.CreateDirectory(parent);
        var workRoot = Path.Combine(
            parent,
            $".{Path.GetFileName(outputDirectory)}-{Guid.NewGuid():N}");
        Directory.CreateDirectory(workRoot);
        try
        {
            var pdfPath = inputPath;
            var backend = "pdfium";
            var backendKind = "pdfium";
            Dictionary<string, object?>? conversionInfo = null;
            if (!extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            {
                var converted = Path.Combine(workRoot, "converted.pdf");
                conversionInfo = ConvertOfficeToPdf(
                    inputPath, converted, sheet, cellRange, isExcel);
                pdfPath = converted;
                backendKind = conversionInfo.GetValueOrDefault("backend_kind") as string
                              ?? "com";
                backend = backendKind == "dsoffice"
                    ? "dsoffice+pdfium"
                    : "com+pdfium";
            }

            var bytes = File.ReadAllBytes(pdfPath);
            var reportedPages = Conversion.GetPageCount(bytes, null);
            var selectedPages = ParsePageNumbers(pages, reportedPages);

            var pageDirectory = Path.Combine(workRoot, "pages");
            Directory.CreateDirectory(pageDirectory);
            var rendered = new List<(int Index, int Page, int Width, int Height,
                string TempPath)>();
            var index = 1;
            foreach (var pageNumber in selectedPages)
            {
                var pagePath = Path.Combine(
                    pageDirectory, $"page-{pageNumber:0000}.png");
                Conversion.SavePng(
                    pagePath, bytes, new Index(pageNumber - 1), null,
                    BuildRenderOptions(bytes, pageNumber, dpi));
                if (!File.Exists(pagePath) || new FileInfo(pagePath).Length == 0)
                    throw new NativeRuntimeException(
                        $"PDF 第 {pageNumber} 页渲染未生成 PNG");
                var (width, height) = ReadPngSize(pagePath);
                rendered.Add((index++, pageNumber, width, height, pagePath));
            }

            PublishDirectory(pageDirectory, outputDirectory, mode, confirm);
            var images = rendered.Select(item => new Dictionary<string, object?>
            {
                ["index"] = item.Index,
                ["page"] = item.Page,
                ["width"] = item.Width,
                ["height"] = item.Height,
                ["path"] = Path.Combine(
                    outputDirectory, Path.GetFileName(item.TempPath)),
            }).ToList();
            var actualCount = images.Count;
            var message = actualCount == reportedPages
                ? $"渲染完成，共 {actualCount} 页"
                : $"渲染完成，共 {actualCount} 页"
                  + $"（上游元数据报告 {reportedPages} 页，已按实际产物修正）";

            var result = new Dictionary<string, object?>
            {
                ["success"] = true,
                ["output_dir"] = outputDirectory,
                ["page_count"] = actualCount,
                ["page_count_reported"] = reportedPages,
                ["images"] = images,
                ["missing_fonts"] = Array.Empty<string>(),
                ["backend"] = backend,
                ["backend_kind"] = backendKind,
                ["message"] = message,
            };
            if (conversionInfo is not null)
            {
                foreach (var key in new[]
                         {
                             "com_prog_id", "com_engine", "com_application",
                             "fallback_from",
                         })
                {
                    if (conversionInfo.TryGetValue(key, out var value) && value is not null)
                        result[key] = value;
                }
            }
            return result;
        }
        finally
        {
            if (Directory.Exists(workRoot))
            {
                try
                {
                    Directory.Delete(workRoot, recursive: true);
                }
                catch
                {
                    // Temporary work products can be cleaned by the OS later.
                }
            }
        }
    }

    private static int ParseDpi(JsonElement p)
    {
        if (!PdfJson.TryGet(p, "dpi", out var value)
            || value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return 144;
        int dpi;
        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var number))
            dpi = number;
        else if (value.ValueKind == JsonValueKind.String
                 && int.TryParse(value.GetString(), out var parsed))
            dpi = parsed;
        else
            throw new NativeUsageException("dpi 必须是 72-600 之间的整数");
        return Math.Clamp(dpi, 72, 600);
    }

    /// <summary>
    /// PDFium truncates fractional pixels; PyMuPDF rounds them up. Compute the
    /// target bitmap size explicitly so rendered dimensions match the Python
    /// backend (e.g. A4 at 72 dpi → 596×842 instead of 595×841).
    /// </summary>
    private static RenderOptions BuildRenderOptions(
        byte[] pdfBytes, int pageNumber, int dpi)
    {
        try
        {
            var size = Conversion.GetPageSize(
                pdfBytes, new Index(pageNumber - 1), null);
            var width = Math.Max(
                1, (int)Math.Ceiling(size.Width * dpi / 72.0));
            var height = Math.Max(
                1, (int)Math.Ceiling(size.Height * dpi / 72.0));
            return new RenderOptions
            {
                Dpi = dpi,
                Width = width,
                Height = height,
                WithAspectRatio = false,
            };
        }
        catch
        {
            return new RenderOptions { Dpi = dpi };
        }
    }

    private static List<int> ParsePageNumbers(string pages, int total)
    {
        if (pages.Equals("all", StringComparison.OrdinalIgnoreCase))
            return Enumerable.Range(1, total).ToList();
        var selected = new List<int>();
        foreach (var item in pages.Split(','))
        {
            var number = int.Parse(item.Trim());
            if (number < 1 || number > total)
                throw new NativeRuntimeException($"页码 {number} 超出范围（共 {total} 页）");
            if (!selected.Contains(number))
                selected.Add(number);
        }
        return selected;
    }

    private static void GuardDirectoryTarget(string target, string mode, bool confirm)
    {
        var fileExists = File.Exists(target);
        var directoryExists = Directory.Exists(target);
        if (!fileExists && !directoryExists)
            return;
        var nonEmpty = fileExists
            ? new FileInfo(target).Length > 0
            : Directory.EnumerateFileSystemEntries(target).Any();
        if (!nonEmpty)
            return;
        if (mode == "overwrite" && confirm)
            return;
        throw new NativeUsageException(
            $"拒绝复写：目标已存在且非空 {target}；确需覆盖请传 "
            + "if_exists=overwrite 与 confirm_overwrite=true");
    }

    private static void PublishDirectory(
        string tempDirectory, string target, string mode, bool confirm)
    {
        GuardDirectoryTarget(target, mode, confirm);
        string? aside = null;
        if (Directory.Exists(target) || File.Exists(target))
        {
            aside = AsidePath(target);
            if (Directory.Exists(target))
                Directory.Move(target, aside);
            else
                File.Move(target, aside);
        }
        try
        {
            Directory.Move(tempDirectory, target);
        }
        catch
        {
            if (!Directory.Exists(target) && aside is not null
                && (Directory.Exists(aside) || File.Exists(aside)))
            {
                if (Directory.Exists(aside))
                    Directory.Move(aside, target);
                else
                    File.Move(aside, target);
            }
            throw;
        }
    }

    private static string AsidePath(string target)
    {
        var stamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
        var candidate = $"{target}.bak-{stamp}";
        var counter = 1;
        while (Directory.Exists(candidate) || File.Exists(candidate))
            candidate = $"{target}.bak-{stamp}-{counter++}";
        return candidate;
    }

    private static (int Width, int Height) ReadPngSize(string path)
    {
        using var stream = File.OpenRead(path);
        Span<byte> header = stackalloc byte[24];
        try
        {
            stream.ReadExactly(header);
        }
        catch (EndOfStreamException)
        {
            throw new NativeRuntimeException($"PNG 文件无效: {path}");
        }
        if (header[0] != 0x89 || header[1] != 0x50 || header[2] != 0x4E
            || header[3] != 0x47)
            throw new NativeRuntimeException($"PNG 文件无效: {path}");
        var width = checked((int)BinaryPrimitives.ReadUInt32BigEndian(header[16..20]));
        var height = checked((int)BinaryPrimitives.ReadUInt32BigEndian(header[20..24]));
        return (width, height);
    }

    // ── Office → PDF ──────────────────────────────────────────────

    private static Dictionary<string, object?> ConvertOfficeToPdf(
        string inputPath, string outputPdf, string? sheet, string? cellRange,
        bool isExcel)
    {
        if (isExcel && !string.IsNullOrWhiteSpace(sheet))
            return ConvertExcelWithCom(inputPath, outputPdf, sheet, cellRange);

        var parameters = new Dictionary<string, object?>
        {
            ["input_path"] = inputPath,
            ["output_path"] = outputPdf,
            ["if_exists"] = "overwrite",
            ["confirm_overwrite"] = true,
        };
        if (!string.IsNullOrWhiteSpace(sheet))
            parameters["sheet"] = sheet;
        using var document = JsonDocument.Parse(
            JsonSerializer.Serialize(parameters));

        // Keep the same auto chain as Python: dsoffice first, then Office COM;
        // ConvertCommands records the failed dsoffice attempt in fallback_from.
        var data = ConvertCommands.RunConvert(document.RootElement);
        if (!File.Exists(outputPdf) || new FileInfo(outputPdf).Length == 0)
            throw new NativeRuntimeException("Office 文档转 PDF 未生成有效文件");
        return data;
    }

    /// <summary>Excel sheet/print-area export; mirrors office_convert._export_with_com.</summary>
    private static Dictionary<string, object?> ConvertExcelWithCom(
        string inputPath, string outputPdf, string? sheet, string? cellRange)
    {
        var errors = new List<string>();
        foreach (var progId in new[] { "Excel.Application", "KET.Application" })
        {
            var temporary = Path.Combine(
                Path.GetDirectoryName(outputPdf)!,
                $".com-{Guid.NewGuid():N}.pdf");
            var processNames = progId.StartsWith("KET", StringComparison.OrdinalIgnoreCase)
                ? new[] { "et", "wps" }
                : new[] { "EXCEL" };
            var before = Snapshot(processNames);
            dynamic? app = null;
            dynamic? workbook = null;
            try
            {
                app = IsolatedCom.Create(progId, "Excel 未注册");
                TrySet(app, "Visible", false);
                TrySet(app, "DisplayAlerts", false);
                workbook = app.Workbooks.Open(inputPath, 0, true);
                if (!string.IsNullOrWhiteSpace(sheet))
                {
                    dynamic worksheet = workbook.Worksheets.Item(sheet);
                    if (!string.IsNullOrWhiteSpace(cellRange))
                        worksheet.PageSetup.PrintArea = cellRange.Replace("$", "");
                    worksheet.ExportAsFixedFormat(0, temporary);
                }
                else
                {
                    workbook.ExportAsFixedFormat(0, temporary);
                }
                workbook.Close(false);
                workbook = null;
                if (!File.Exists(temporary) || new FileInfo(temporary).Length == 0)
                    throw new NativeRuntimeException($"{progId} 未生成有效 PDF");
                File.Move(temporary, outputPdf, overwrite: true);
                return new Dictionary<string, object?>
                {
                    ["success"] = true,
                    ["file_path"] = outputPdf,
                    ["backend"] = "com",
                    ["backend_kind"] = "com",
                    ["com_prog_id"] = progId,
                    ["com_engine"] = progId.StartsWith(
                        "KET", StringComparison.OrdinalIgnoreCase) ? "wps" : "office",
                    ["com_application"] = progId,
                    ["missing_fonts"] = Array.Empty<string>(),
                };
            }
            catch (Exception ex)
            {
                errors.Add($"{progId}: {ex.Message}");
            }
            finally
            {
                if (workbook is not null)
                {
                    try { workbook.Close(false); } catch { }
                }
                if (app is not null)
                {
                    try { app.Quit(); } catch { }
                    NativeCommon.ReleaseCom(workbook);
                    NativeCommon.ReleaseCom(app);
                }
                KillNewProcesses(processNames, before);
                if (File.Exists(temporary))
                {
                    try { File.Delete(temporary); } catch { }
                }
            }
        }
        throw new NativeRuntimeException(
            "COM 转换失败（已尝试 Microsoft Excel 与 WPS）：" + string.Join("；", errors));
    }

    private static HashSet<int> Snapshot(IEnumerable<string> processNames) =>
        processNames.SelectMany(Process.GetProcessesByName)
            .Select(process => process.Id).ToHashSet();

    private static void KillNewProcesses(
        IEnumerable<string> processNames, HashSet<int> before)
    {
        foreach (var process in processNames.SelectMany(Process.GetProcessesByName))
        {
            if (before.Contains(process.Id))
                continue;
            try
            {
                if (!process.HasExited)
                    process.Kill();
            }
            catch
            {
                // Process exited between enumeration and cleanup.
            }
        }
    }

    private static void TrySet(dynamic target, string property, object value)
    {
        try
        {
            var instance = (object)target;
            instance.GetType().InvokeMember(
                property, System.Reflection.BindingFlags.SetProperty, null,
                instance, new[] { value });
        }
        catch
        {
            // Optional COM properties are best-effort.
        }
    }

    /// <summary>DispatchEx equivalent: create an isolated Office instance.</summary>
    private static class IsolatedCom
    {
        private const uint ClsCtxLocalServer = 4;
        private static readonly Guid IidClassFactory =
            new("00000001-0000-0000-C000-000000000046");
        private static readonly Guid IidDispatch =
            new("00020400-0000-0000-C000-000000000046");

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
                    throw new NativeRuntimeException(
                        $"无法获取 {progId} 的 IClassFactory");
                var dispatchIid = IidDispatch;
                if (classFactory.CreateInstance(null, ref dispatchIid, out var instance) != 0)
                    throw new NativeRuntimeException(
                        $"无法创建隔离的 {progId} 实例");
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
