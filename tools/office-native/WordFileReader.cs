using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeNative;

/// <summary>Reads .docx body content into the Python reader's Markdown/statistics shape.</summary>
internal static class WordFileReader
{
    private const string TruncationMarker = "_(内容已截断)_";

    private static readonly Regex HeadingStylePattern =
        new(@"(?:Heading|heading)\s*(\d+)", RegexOptions.Compiled);

    public static Dictionary<string, object?> RunReadWordDocument(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        var maxChars = NativeCommon.OptInt(p, "max_chars", 50000);
        if (maxChars < 1)
            throw new NativeRuntimeException("max_chars 必须大于 0");
        if (!File.Exists(filePath))
            throw new NativeRuntimeException($"文件不存在: {filePath}");
        if (!Path.GetExtension(filePath).Equals(".docx", StringComparison.OrdinalIgnoreCase))
            throw new NativeRuntimeException("不支持的文件格式，仅支持 .docx");
        WordFileOoxml.ValidateOfficeInput(filePath);

        try
        {
            using var document = WordprocessingDocument.Open(filePath, false);
            var main = document.MainDocumentPart
                       ?? throw new NativeRuntimeException("文档没有主部件");
            var body = main.Document?.Body
                       ?? throw new NativeRuntimeException("文档没有正文");
            var styleNames = LoadStyleNames(main);
            var defaultStyle = DefaultStyleName(main, styleNames);
            var parts = new List<string>();
            var paragraphCount = 0;

            foreach (var paragraph in body.Elements<Paragraph>())
            {
                var text = WordFileOoxml.ParagraphText(paragraph).Trim();
                if (text.Length == 0)
                    continue;
                paragraphCount++;
                var styleName = StyleNameOf(paragraph, styleNames, defaultStyle);
                var heading = HeadingStylePattern.Match(styleName);
                if (heading.Success)
                {
                    var level = Math.Min(int.Parse(heading.Groups[1].Value), 6);
                    parts.Add(new string('#', level) + " " + text);
                }
                else if (styleName.Contains("List", StringComparison.Ordinal))
                {
                    var prefix = styleName.Contains("Bullet", StringComparison.Ordinal)
                        ? "- "
                        : "1. ";
                    parts.Add(prefix + text);
                }
                else
                {
                    parts.Add(text);
                }
            }

            var tableCount = body.Elements<Table>().Count();
            var tableIndex = 0;
            foreach (var table in body.Elements<Table>())
            {
                tableIndex++;
                var rows = table.Elements<TableRow>().ToList();
                if (rows.Count == 0)
                    continue;
                parts.Add($"\n[Table {tableIndex}]");
                var headerCells = RowCells(rows[0]);
                if (headerCells.Any(cell => cell.Length > 0))
                {
                    parts.Add("| " + string.Join(" | ", headerCells) + " |");
                    parts.Add("| " + string.Join(" | ", headerCells.Select(_ => "---")) + " |");
                }
                foreach (var row in rows.Skip(1))
                    parts.Add("| " + string.Join(" | ", RowCells(row)) + " |");
            }

            var (content, truncated) = TruncateAtBoundary(parts, maxChars);
            var message = $"成功读取文档，共 {paragraphCount} 段落、{tableCount} 个表格";
            if (truncated)
                message += $"（已截断至 {maxChars} 字符）";
            var result = new Dictionary<string, object?>
            {
                ["success"] = true,
                ["file_path"] = filePath,
                ["content"] = content,
                ["paragraphs_count"] = paragraphCount,
                ["tables_count"] = tableCount,
                ["truncated"] = truncated,
                ["message"] = message,
            };
            return new Dictionary<string, object?>
            {
                ["markdown"] = BuildMarkdownView(result),
                ["result"] = result,
            };
        }
        catch (NativeRuntimeException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new NativeRuntimeException($"读取 Word 文档失败: {ex.Message}");
        }
    }

    private static Dictionary<string, string> LoadStyleNames(MainDocumentPart main)
    {
        var names = new Dictionary<string, string>(StringComparer.Ordinal);
        var styles = main.StyleDefinitionsPart?.Styles;
        if (styles is null)
            return names;
        foreach (var style in styles.Elements<Style>())
        {
            var styleId = style.StyleId?.Value;
            if (string.IsNullOrEmpty(styleId))
                continue;
            names[styleId] = style.StyleName?.Val?.Value ?? styleId;
        }
        return names;
    }

    private static string DefaultStyleName(
        MainDocumentPart main, Dictionary<string, string> styleNames)
    {
        var styles = main.StyleDefinitionsPart?.Styles;
        var defaultStyle = styles?.Elements<Style>().FirstOrDefault(style =>
            style.Default?.Value == true && style.Type?.Value == StyleValues.Paragraph);
        var styleId = defaultStyle?.StyleId?.Value;
        if (styleId is not null && styleNames.TryGetValue(styleId, out var name))
            return name;
        return styleNames.TryGetValue("Normal", out var normal) ? normal : "Normal";
    }

    private static string StyleNameOf(
        Paragraph paragraph,
        Dictionary<string, string> styleNames,
        string defaultStyle)
    {
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (string.IsNullOrEmpty(styleId))
            return defaultStyle;
        return styleNames.TryGetValue(styleId, out var name) ? name : styleId;
    }

    private static List<string> RowCells(TableRow row)
    {
        var cells = new List<string>();
        foreach (var cell in row.Elements<TableCell>())
        {
            var text = string.Join("\n",
                cell.Elements<Paragraph>().Select(WordFileOoxml.ParagraphText)).Trim();
            var span = Math.Max(1, cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1);
            for (var index = 0; index < span; index++)
                cells.Add(text);
        }
        return cells;
    }

    private static (string Content, bool Truncated) TruncateAtBoundary(
        List<string> parts, int maxChars)
    {
        var content = string.Join("\n\n", parts);
        if (content.Length <= maxChars)
            return (content, false);
        var length = 0;
        for (var index = 0; index < parts.Count; index++)
        {
            var separator = index > 0 ? 2 : 0;
            if (length + separator + parts[index].Length > maxChars)
                return (ContentWithTruncationMarker(parts.Take(index), maxChars), true);
            length += separator + parts[index].Length;
        }
        return (content, false);
    }

    private static string ContentWithTruncationMarker(
        IEnumerable<string> parts, int maxChars)
    {
        var content = string.Join("\n\n", parts);
        if (content.Length == 0)
            return TruncationMarker[..Math.Min(maxChars, TruncationMarker.Length)];
        var candidate = content + "\n\n" + TruncationMarker;
        return candidate.Length <= maxChars ? candidate : content;
    }

    private static string BuildMarkdownView(Dictionary<string, object?> result)
    {
        var parts = new List<string>();
        if (result.TryGetValue("content", out var contentValue)
            && contentValue is string content && content.Length > 0)
            parts.Add(content);
        var stats = new List<string>();
        foreach (var pair in result)
        {
            if (pair.Key is "success" or "file_path" or "content" or "message" or "truncated")
                continue;
            if (pair.Value is null)
                continue;
            stats.Add($"{pair.Key}: {pair.Value}");
        }
        if (stats.Count > 0)
            parts.Add("> " + string.Join("；", stats));
        if (result.TryGetValue("truncated", out var truncated) && truncated is true)
            parts.Add("> 内容过长已截断");
        return string.Join("\n\n", parts.Where(part => part.Length > 0));
    }
}
