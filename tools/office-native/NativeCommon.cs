using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;

namespace OfficeNative;

internal sealed class NativeUsageException : Exception
{
    public NativeUsageException(string message) : base(message) { }
}

internal sealed class NativeRuntimeException : Exception
{
    public NativeRuntimeException(string message) : base(message) { }
}

internal static class NativeCommon
{
    public const string DsofficeInstallHint =
        "未找到 dsoffice CLI。请先安装：npm i -g dsoffice"
        + "（要求 Node.js >= 22.19；Windows x64 需 MSVC v14 运行库），"
        + "检测方法：在终端运行 which dsoffice（PowerShell 用 Get-Command dsoffice），"
        + "或通过环境变量 MCP_OFFICE_DSOFFICE（兼容 MODULE_OFFICE_DSOFFICE）指定 dsoffice 可执行文件路径";

    public const string BackendUnavailableHint =
        "未找到可用的 Office 转换后端。可任选一种方式启用：\n"
        + "1) dsoffice：npm i -g dsoffice（Node.js >= 22.19；Windows x64 需 MSVC v14 运行库），"
        + "安装后运行 which dsoffice 检测，或设 MCP_OFFICE_DSOFFICE 指向可执行文件；\n"
        + "2) Microsoft Office COM：Windows 安装 Microsoft Office（Word/Excel/PowerPoint）与 pywin32，"
        + "并设 OFFICE_CONVERT_BACKEND=com（或 auto）；\n"
        + "3) WPS Office COM：Windows 安装 WPS Office（KWps.Application/KET.Application/"
        + "KWPP.Application）与 pywin32，并设 OFFICE_CONVERT_BACKEND=com（或 auto）。";

    public static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public static string RequireString(JsonElement p, string name)
    {
        var value = OptString(p, name);
        if (string.IsNullOrWhiteSpace(value))
            throw new NativeUsageException($"缺少参数 {name}");
        return value;
    }

    public static string? OptString(JsonElement p, string name)
    {
        if (p.ValueKind != JsonValueKind.Object || !p.TryGetProperty(name, out var v))
            return null;
        return v.ValueKind switch
        {
            JsonValueKind.String => v.GetString(),
            JsonValueKind.Null => null,
            _ => v.ToString(),
        };
    }

    public static int OptInt(JsonElement p, string name, int fallback)
    {
        if (p.ValueKind != JsonValueKind.Object || !p.TryGetProperty(name, out var v))
            return fallback;
        if (v.ValueKind == JsonValueKind.Number && v.TryGetInt32(out var number))
            return number;
        if (v.ValueKind == JsonValueKind.String && int.TryParse(v.GetString(), out var parsed))
            return parsed;
        return fallback;
    }

    public static double OptDouble(JsonElement p, string name, double fallback)
    {
        if (p.ValueKind != JsonValueKind.Object || !p.TryGetProperty(name, out var v))
            return fallback;
        if (v.ValueKind == JsonValueKind.Number && v.TryGetDouble(out var number))
            return number;
        if (v.ValueKind == JsonValueKind.String && double.TryParse(v.GetString(), out var parsed))
            return parsed;
        return fallback;
    }

    public static bool OptBool(JsonElement p, string name, bool fallback)
    {
        if (p.ValueKind != JsonValueKind.Object || !p.TryGetProperty(name, out var v))
            return fallback;
        if (v.ValueKind is JsonValueKind.True or JsonValueKind.False)
            return v.GetBoolean();
        if (v.ValueKind == JsonValueKind.Number && v.TryGetInt32(out var number))
            return number switch { 0 => false, 1 => true, _ => fallback };
        if (v.ValueKind == JsonValueKind.String)
        {
            var text = v.GetString()?.Trim();
            if (text is "0" or "1")
                return text == "1";
            if (bool.TryParse(text, out var parsed))
                return parsed;
        }
        return fallback;
    }

    public static List<string> OptStringList(JsonElement p, string name)
    {
        if (p.ValueKind != JsonValueKind.Object || !p.TryGetProperty(name, out var v)
            || v.ValueKind != JsonValueKind.Array)
            return [];
        return v.EnumerateArray()
            .Where(item => item.ValueKind == JsonValueKind.String)
            .Select(item => item.GetString()!)
            .ToList();
    }

    private const long MaxInputBytes = 100L * 1024 * 1024;
    private const long MaxUncompressedBytes = 500L * 1024 * 1024;
    private static readonly HashSet<string> OoxmlExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".docx", ".dotx", ".xlsx", ".xlsm", ".xltx", ".xltm", ".pptx", ".potx",
    };

    public static void EnsureFile(string path, params string[] extensions)
    {
        EnsureFileLight(path, extensions);
        if (OoxmlExtensions.Contains(Path.GetExtension(path)))
        {
            try
            {
                using var archive = System.IO.Compression.ZipFile.OpenRead(path);
                if (archive.Entries.Sum(entry => entry.Length) > MaxUncompressedBytes)
                    throw new NativeRuntimeException("Office 文件解压后超过 500 MB 安全上限");
            }
            catch (InvalidDataException ex)
            {
                throw new NativeRuntimeException($"Office 文件损坏或不是有效的 OOXML 压缩包: {ex.Message}");
            }
        }
    }

    /// <summary>活动文档已被 Office 锁定，不能解压检查，只做存在/扩展名/大小校验。</summary>
    public static void EnsureFileLight(string path, params string[] extensions)
    {
        if (!File.Exists(path))
            throw new NativeRuntimeException($"文件不存在: {path}");
        if (extensions.Length > 0)
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            if (!extensions.Contains(ext, StringComparer.OrdinalIgnoreCase))
                throw new NativeRuntimeException(
                    $"不支持的格式: {ext}（支持 {string.Join("/", extensions)}）");
        }
        if (new FileInfo(path).Length > MaxInputBytes)
            throw new NativeRuntimeException("文件超过 100 MB 安全读取上限");
    }

    public static void EnsureOutput(string path, bool overwrite, bool confirm)
    {
        if (!File.Exists(path))
            return;
        if (new FileInfo(path).Length == 0)
            return;
        if (overwrite && confirm)
            return;
        throw new NativeUsageException(
            $"拒绝复写：目标已存在且非空 {path}；确需覆盖请传 overwrite=true 与 confirm_overwrite=true");
    }

    /// <summary>写回前复制时间戳备份，返回备份路径。</summary>
    public static string Backup(string path)
    {
        var stamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
        var backup = $"{path}.bak-{stamp}";
        var counter = 1;
        while (File.Exists(backup))
            backup = $"{path}.bak-{stamp}-{counter++}";
        File.Copy(path, backup, overwrite: false);
        return backup;
    }

    public static bool ProcessRunning(string name) =>
        Process.GetProcessesByName(name).Length > 0;

    public static void ReleaseCom(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
        {
            try { Marshal.ReleaseComObject(value); } catch { /* COM already released */ }
        }
    }

    public static string? JsonString(JsonElement p, string name) => OptString(p, name);
}
