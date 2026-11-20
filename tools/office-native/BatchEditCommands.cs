using System.Text.Json;

namespace OfficeNative;

/// <summary>
/// batch-edit-office-document：内存中应用全部操作，成功后临时文件 + 原子替换写回。
/// </summary>
internal static class BatchEditCommands
{
    private static readonly string[] AllowedExtensions =
        [".docx", ".xlsx", ".xlsm", ".pptx"];

    public static Dictionary<string, object?> RunBatchEdit(JsonElement p)
    {
        try
        {
            return Run(p);
        }
        catch (BatchEditException ex)
        {
            throw new NativeRuntimeException(ex.ToDisplay());
        }
        catch (Exception ex) when (ex is not NativeUsageException
                                   and not NativeRuntimeException)
        {
            throw new NativeRuntimeException($"批量编辑失败: {ex.Message}");
        }
    }

    private static Dictionary<string, object?> Run(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        var outputPath = NativeCommon.OptString(p, "output_path");
        if (string.IsNullOrWhiteSpace(outputPath))
            outputPath = null;
        var noBackup = NativeCommon.OptBool(p, "no_backup", false);
        var format = DetectFormat(filePath);
        NativeCommon.EnsureFile(filePath, ExtensionsFor(format));

        string targetPath;
        if (outputPath is not null)
        {
            ValidateOutputPath(filePath, outputPath, format);
            targetPath = ApplyOutputPolicy(p, outputPath);
        }
        else
        {
            targetPath = filePath;
        }
        var operations = BatchEditOperations.Parse(p);

        var fullSource = Path.GetFullPath(filePath);
        var fullTarget = Path.GetFullPath(targetPath);
        var targetDirectory = Path.GetDirectoryName(fullTarget)
                              ?? throw new NativeRuntimeException(
                                  $"无法解析输出目录: {targetPath}");
        Directory.CreateDirectory(targetDirectory);
        var workPath = Path.Combine(targetDirectory,
            $".{Path.GetFileNameWithoutExtension(fullTarget)}-{Guid.NewGuid():N}.tmp{Path.GetExtension(fullTarget)}");

        string? backupPath = null;
        try
        {
            File.Copy(fullSource, workPath, overwrite: false);
            var attributes = File.GetAttributes(workPath);
            if ((attributes & FileAttributes.ReadOnly) != 0)
                File.SetAttributes(workPath, attributes & ~FileAttributes.ReadOnly);
            var messages = format switch
            {
                "docx" => BatchEditWord.Apply(workPath, operations),
                "xlsx" => BatchEditExcel.Apply(workPath, operations),
                _ => BatchEditPpt.Apply(workPath, operations),
            };

            if (File.Exists(fullTarget) && new FileInfo(fullTarget).Length > 0)
            {
                // no_backup 只作用于“省略 output_path 的原地写回”；
                // 显式输出路径覆盖已有目标时仍保留备份（与 Python atomic_save 一致）。
                if (!(outputPath is null && noBackup))
                    backupPath = NativeCommon.Backup(fullTarget);
            }

            ReplaceFile(workPath, fullTarget);

            var results = new List<Dictionary<string, object?>>(operations.Count);
            for (var index = 0; index < operations.Count; index++)
            {
                results.Add(new Dictionary<string, object?>
                {
                    ["index"] = index,
                    ["success"] = true,
                    ["op"] = operations[index].Op,
                    ["path"] = operations[index].Path,
                    ["message"] = messages[index],
                });
            }
            var data = new Dictionary<string, object?>
            {
                ["success"] = true,
                ["file_path"] = filePath,
                ["output_path"] = targetPath,
                ["saved"] = true,
                ["file_size"] = new FileInfo(fullTarget).Length,
                ["operation_count"] = operations.Count,
                ["completed_count"] = operations.Count,
                ["results"] = results,
                ["message"] = $"已原子执行 {operations.Count} 个操作",
            };
            // 与 Python 入口一致：只有真正生成备份时才返回 backup_path。
            if (backupPath is not null)
                data["backup_path"] = backupPath;
            return data;
        }
        finally
        {
            TryDelete(workPath);
        }
    }

    private static string DetectFormat(string filePath)
    {
        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        return extension switch
        {
            ".docx" => "docx",
            ".xlsx" or ".xlsm" => "xlsx",
            ".pptx" => "pptx",
            _ => throw new BatchEditException(
                "unsupported_type",
                $"不支持的文件格式: {(extension.Length == 0 ? "(无扩展名)" : extension)}",
                filePath,
                validValues: AllowedExtensions),
        };
    }

    private static string[] ExtensionsFor(string format) => format switch
    {
        "docx" => [".docx"],
        "xlsx" => [".xlsx", ".xlsm"],
        _ => [".pptx"],
    };

    private static void ValidateOutputPath(
        string filePath, string outputPath, string format)
    {
        if (Directory.Exists(outputPath))
            throw new NativeUsageException($"输出路径是目录，不能作为文件: {outputPath}");
        if (outputPath.Any(character => "<>\"|?*".Contains(character)))
            throw new NativeUsageException($"路径包含非法字符: {outputPath}");
        var sourceExtension = Path.GetExtension(filePath).ToLowerInvariant();
        var outputExtension = Path.GetExtension(outputPath).ToLowerInvariant();
        if (outputExtension != sourceExtension)
            throw new BatchEditException(
                "extension_mismatch", "output_path 必须与源文件使用相同扩展名",
                outputPath);
        var allowed = ExtensionsFor(format);
        if (!allowed.Contains(outputExtension, StringComparer.OrdinalIgnoreCase))
            throw new BatchEditException(
                "unsupported_type",
                $"output_path 必须使用以下扩展名之一: {string.Join("、", allowed)}",
                outputPath, validValues: allowed);
    }

    private static string ApplyOutputPolicy(JsonElement p, string targetPath)
    {
        var mode = NativeCommon.OptString(p, "if_exists");
        if (string.IsNullOrEmpty(mode))
        {
            if (NativeCommon.OptBool(p, "rename", false))
                mode = "rename";
            else if (NativeCommon.OptBool(p, "overwrite", false))
                mode = "overwrite";
            else
                mode = "error";
        }
        if (mode is not ("error" or "overwrite" or "rename"))
            throw new NativeUsageException(
                $"if_exists 必须是 error/overwrite/rename 之一，收到 {mode}");
        if (mode == "rename")
            return MakeRenamedPath(targetPath);
        NativeCommon.EnsureOutput(targetPath, mode == "overwrite",
            NativeCommon.OptBool(p, "confirm_overwrite", false));
        return targetPath;
    }

    /// <summary>rename 模式返回第一个不存在的 “名称 (N).扩展名”；空白目标无需改名。</summary>
    private static string MakeRenamedPath(string targetPath)
    {
        var full = Path.GetFullPath(targetPath);
        if (!File.Exists(full) || new FileInfo(full).Length == 0)
            return targetPath;
        var directory = Path.GetDirectoryName(full)!;
        var stem = Path.GetFileNameWithoutExtension(full);
        var extension = Path.GetExtension(full);
        for (var counter = 1; ; counter++)
        {
            var candidate = Path.Combine(directory, $"{stem} ({counter}){extension}");
            if (!File.Exists(candidate))
                return candidate;
        }
    }

    private static void ReplaceFile(string workPath, string targetPath)
    {
        var fullTarget = Path.GetFullPath(targetPath);
        if (!File.Exists(fullTarget))
        {
            File.Move(workPath, fullTarget);
            return;
        }
        try
        {
            File.Replace(workPath, fullTarget, null,
                ignoreMetadataErrors: true);
        }
        catch (PlatformNotSupportedException)
        {
            // 仅当平台明确不支持 Replace 时才退回覆盖移动；失败仍按占用处理。
            try
            {
                File.Move(workPath, fullTarget, overwrite: true);
            }
            catch (IOException ex)
            {
                throw new NativeRuntimeException(
                    "无法替换文件（文件可能被其他程序占用，请关闭占用该文件的程序后重试）: "
                    + ex.Message);
            }
        }
        catch (IOException ex)
        {
            // 不使用 File.Copy 原地覆盖：覆盖中途失败会破坏原文件。
            throw new NativeRuntimeException(
                "无法替换文件（文件可能被其他程序占用，请关闭占用该文件的程序后重试）: "
                + ex.Message);
        }
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        catch
        {
            // 临时文件清理失败不影响已提交结果。
        }
    }
}
