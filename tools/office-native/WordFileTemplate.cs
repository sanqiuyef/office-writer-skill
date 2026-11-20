using System.Globalization;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeNative;

/// <summary>docx branch of merge-office-template: {{key}} replacement.</summary>
internal static class WordFileTemplate
{
    private static readonly Regex PlaceholderPattern =
        new(@"\{\{\s*([^{}]+?)\s*\}\}", RegexOptions.Compiled);

    private static readonly JsonSerializerOptions RelaxedJsonOptions = new()
    {
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        WriteIndented = false,
    };

    public static Dictionary<string, object?> RunMergeOfficeTemplate(JsonElement p)
    {
        var templatePath = NativeCommon.RequireString(p, "template_path");
        var outputPath = NativeCommon.RequireString(p, "output_path");
        if (!WordJson.TryGet(p, "data", out var dataElement)
            || dataElement.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException(
                "缺少参数 data（通过 --params 或 --params-file 传入 JSON 对象）");
        var mode = NativeCommon.OptString(p, "if_exists");
        var confirm = NativeCommon.OptBool(p, "confirm_overwrite", false);
        try
        {
            var suffix = Path.GetExtension(templatePath).ToLowerInvariant();
            if (suffix != ".docx")
            {
                throw new NativeRuntimeException(
                    $"不支持的模板格式: {(suffix.Length == 0 ? "(无扩展名)" : suffix)}"
                    + "（当前 C# 后端仅实现 docx 分支）");
            }
            if (!File.Exists(templatePath))
                throw new NativeRuntimeException($"文件不存在: {templatePath}");
            WordFileOoxml.ValidateOfficeInput(templatePath);
            if (!Path.GetExtension(outputPath).Equals(".docx", StringComparison.OrdinalIgnoreCase))
                throw new NativeRuntimeException("输出文件必须使用以下扩展名之一: .docx");
            var resolvedPath = WordFileCommands.ResolveOutputPath(outputPath, mode, confirm);
            var values = Flatten(dataElement);
            var used = new HashSet<string>(StringComparer.Ordinal);
            var unresolved = new HashSet<string>(StringComparer.Ordinal);
            var replacedCount = 0;
            var tempPath = WordFileOoxml.TempSiblingPath(resolvedPath);
            try
            {
                File.Copy(Path.GetFullPath(templatePath), tempPath, overwrite: true);
                using (var document = WordprocessingDocument.Open(tempPath, true))
                {
                    var main = document.MainDocumentPart
                               ?? throw new NativeRuntimeException("模板缺少主文档部件");
                    if (main.Document?.Body is null)
                        throw new NativeRuntimeException("模板没有正文");
                    var paragraphs = EnumerateParagraphs(main).ToList();
                    foreach (var paragraph in paragraphs)
                        replacedCount += ReplaceParagraph(paragraph, values, used);
                    foreach (var paragraph in paragraphs)
                    {
                        foreach (Match match in PlaceholderPattern.Matches(
                                     WordFileOoxml.ParagraphText(paragraph)))
                            unresolved.Add(match.Groups[1].Value.Trim());
                    }
                    main.Document.Save();
                }
                if (File.Exists(resolvedPath) && new FileInfo(resolvedPath).Length > 0)
                    NativeCommon.Backup(resolvedPath);
                WordFileOoxml.ReplaceFile(tempPath, resolvedPath);
            }
            finally
            {
                if (File.Exists(tempPath))
                {
                    try { File.Delete(tempPath); } catch { /* best effort */ }
                }
            }
            var fileSize = new FileInfo(resolvedPath).Length;
            return new Dictionary<string, object?>
            {
                ["success"] = true,
                ["template_path"] = templatePath,
                ["file_path"] = resolvedPath,
                ["file_size"] = fileSize,
                ["replaced_count"] = replacedCount,
                ["used_keys"] = used.OrderBy(key => key, StringComparer.Ordinal).ToList(),
                ["unresolved_placeholders"] =
                    unresolved.OrderBy(key => key, StringComparer.Ordinal).ToList(),
                ["message"] = $"模板合并完成，替换 {replacedCount} 处，"
                              + $"占位符未解析 {unresolved.Count} 个",
            };
        }
        catch (NativeUsageException)
        {
            throw;
        }
        catch (NativeRuntimeException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new NativeRuntimeException($"模板合并失败: {ex.Message}");
        }
    }

    // ── data flattening ─────────────────────────────────────────

    private static Dictionary<string, string> Flatten(JsonElement data)
    {
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var property in data.EnumerateObject())
            values[property.Name] = Stringify(property.Value);
        foreach (var property in data.EnumerateObject())
            Visit(property.Value, property.Name);
        return values;

        void Visit(JsonElement value, string prefix)
        {
            if (value.ValueKind == JsonValueKind.Object)
            {
                foreach (var child in value.EnumerateObject())
                {
                    var path = prefix + "." + child.Name;
                    if (!values.ContainsKey(path))
                        values[path] = Stringify(child.Value);
                    Visit(child.Value, path);
                }
            }
            else if (value.ValueKind == JsonValueKind.Array)
            {
                var index = 0;
                foreach (var child in value.EnumerateArray())
                {
                    var path = prefix + "[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                    if (!values.ContainsKey(path))
                        values[path] = Stringify(child);
                    Visit(child, path);
                    index++;
                }
            }
        }
    }

    private static string Stringify(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.Null => "",
        JsonValueKind.String => value.GetString() ?? "",
        JsonValueKind.True => "True",
        JsonValueKind.False => "False",
        JsonValueKind.Object or JsonValueKind.Array => JsonDump(value),
        _ => value.GetRawText(),
    };

    /// <summary>Python json.dumps(value, ensure_ascii=False) equivalent.</summary>
    private static string JsonDump(JsonElement value)
    {
        switch (value.ValueKind)
        {
            case JsonValueKind.Object:
                var properties = value.EnumerateObject()
                    .Select(property => JsonSerializer.Serialize(property.Name, RelaxedJsonOptions)
                                        + ": " + JsonDump(property.Value));
                return "{" + string.Join(", ", properties) + "}";
            case JsonValueKind.Array:
                var items = value.EnumerateArray().Select(JsonDump);
                return "[" + string.Join(", ", items) + "]";
            case JsonValueKind.String:
                return JsonSerializer.Serialize(value.GetString(), RelaxedJsonOptions);
            case JsonValueKind.True:
                return "true";
            case JsonValueKind.False:
                return "false";
            case JsonValueKind.Null:
                return "null";
            default:
                return value.GetRawText();
        }
    }

    // ── placeholder replacement ─────────────────────────────────

    private static int ReplaceParagraph(
        Paragraph paragraph,
        Dictionary<string, string> values,
        HashSet<string> used)
    {
        var runs = paragraph.Descendants<Run>().ToList();
        var fullText = string.Concat(runs.Select(WordFileOoxml.RunText));
        var (changed, count) = ReplacePlaceholders(fullText, values, used);
        if (count == 0)
            return 0;
        if (runs.Count > 0)
        {
            WordFileOoxml.SetRunText(runs[0], changed);
            foreach (var run in runs.Skip(1))
            {
                foreach (var text in run.Elements<Text>())
                    text.Text = "";
            }
        }
        else
        {
            paragraph.Append(new Run(new Text(changed)
            {
                Space = SpaceProcessingModeValues.Preserve,
            }));
        }
        return count;
    }

    private static (string Text, int Count) ReplacePlaceholders(
        string text, Dictionary<string, string> values, HashSet<string> used)
    {
        var count = 0;
        var result = PlaceholderPattern.Replace(text, match =>
        {
            var key = match.Groups[1].Value.Trim();
            if (!values.TryGetValue(key, out var replacement))
                return match.Value;
            used.Add(key);
            count++;
            return replacement;
        });
        return (result, count);
    }

    // ── paragraph enumeration ───────────────────────────────────

    private static IEnumerable<Paragraph> EnumerateParagraphs(MainDocumentPart main)
    {
        var body = main.Document?.Body;
        if (body is not null)
        {
            foreach (var paragraph in body.Elements<Paragraph>())
                yield return paragraph;
            foreach (var table in body.Elements<Table>())
            {
                foreach (var paragraph in TableParagraphs(table))
                    yield return paragraph;
            }
        }
        foreach (var part in main.HeaderParts)
        {
            if (part.Header is null)
                continue;
            foreach (var paragraph in ContainerParagraphs(part.Header))
                yield return paragraph;
        }
        foreach (var part in main.FooterParts)
        {
            if (part.Footer is null)
                continue;
            foreach (var paragraph in ContainerParagraphs(part.Footer))
                yield return paragraph;
        }
    }

    private static IEnumerable<Paragraph> ContainerParagraphs(OpenXmlElement container)
    {
        foreach (var paragraph in container.Elements<Paragraph>())
            yield return paragraph;
        foreach (var table in container.Elements<Table>())
        {
            foreach (var paragraph in TableParagraphs(table))
                yield return paragraph;
        }
    }

    private static IEnumerable<Paragraph> TableParagraphs(Table table)
    {
        foreach (var row in table.Elements<TableRow>())
        {
            foreach (var cell in row.Elements<TableCell>())
            {
                foreach (var paragraph in cell.Elements<Paragraph>())
                    yield return paragraph;
                foreach (var nested in cell.Elements<Table>())
                {
                    foreach (var paragraph in TableParagraphs(nested))
                        yield return paragraph;
                }
            }
        }
    }
}
