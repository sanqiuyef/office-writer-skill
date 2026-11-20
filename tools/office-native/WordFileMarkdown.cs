using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace OfficeNative;

/// <summary>Markdown file → WordContent conversion (mirrors office_engine/markdown_writer.py).</summary>
internal static class WordFileMarkdown
{
    private static readonly Regex HeadingPattern =
        new(@"^(#{1,6})\s+(.+?)\s*$", RegexOptions.Compiled);

    private static readonly Regex BulletPattern =
        new(@"^\s*[-*+]\s+(.+?)\s*$", RegexOptions.Compiled);

    private static readonly Regex NumberedPattern =
        new(@"^\s*\d+[.)]\s+(.+?)\s*$", RegexOptions.Compiled);

    private static readonly Regex HorizontalRulePattern =
        new(@"^\s*(?:---+|\*\*\*+|___+)\s*$", RegexOptions.Compiled);

    private static readonly Regex TableDividerPattern =
        new(@"^\s*\|?(?:\s*:?-{3,}:?\s*\|)+\s*$", RegexOptions.Compiled);

    private static readonly Regex ImagePattern =
        new(@"^\s*!\[([^\]]*)\]\(([^\s)]+)(?:\s+[""'][^""']*[""'])?\)\s*$",
            RegexOptions.Compiled);

    private static readonly string[] SupportedImageSuffixes =
        [".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff"];

    public static Dictionary<string, object?> RunCreateWordFromMarkdown(JsonElement p)
    {
        var sourcePath = NativeCommon.RequireString(p, "source_path");
        var outputPath = NativeCommon.RequireString(p, "output_path");
        var template = NativeCommon.OptString(p, "template");
        // Python CLI defaults this entry point to scene="report".
        var scene = NativeCommon.OptString(p, "scene") ?? "report";
        try
        {
            var markdown = ReadMarkdownFile(sourcePath);
            var sourceDirectory = Path.GetDirectoryName(Path.GetFullPath(sourcePath));
            var sections = MarkdownToSections(markdown, sourceDirectory);
            var content = new WordContent { Sections = sections };
            var result = WordFileCommands.CreateDocumentResult(
                p, outputPath, content, template, null, null, null, null, scene);
            if (result.TryGetValue("success", out var success) && success is true)
            {
                var resolvedPath = result["file_path"];
                result["message"] =
                    $"Markdown 转 Word 成功，共转换 {sections.Count} 个内容块：{resolvedPath}";
            }
            return result;
        }
        catch (NativeUsageException)
        {
            throw;
        }
        catch (NativeRuntimeException ex)
        {
            throw new NativeRuntimeException($"Markdown 转 Word 失败: {ex.Message}");
        }
        catch (Exception ex)
        {
            throw new NativeRuntimeException($"Markdown 转 Word 失败: {ex.Message}");
        }
    }

    private static string ReadMarkdownFile(string filePath)
    {
        var path = Path.GetFullPath(filePath);
        if (!File.Exists(path))
            throw new NativeRuntimeException($"文件不存在: {filePath}");
        var suffix = Path.GetExtension(path).ToLowerInvariant();
        if (suffix is not (".md" or ".markdown" or ".txt"))
            throw new NativeRuntimeException("不支持的文件格式，仅支持: .md、.markdown、.txt");
        if (new FileInfo(path).Length > 100L * 1024 * 1024)
            throw new NativeRuntimeException("文件超过 100 MB 安全读取上限");
        var data = File.ReadAllBytes(path);

        if (data.Length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF)
            return new UTF8Encoding(false, true).GetString(data, 3, data.Length - 3);
        if (data.Length >= 2 && data[0] == 0xFF && data[1] == 0xFE)
            return new UnicodeEncoding(false, false, true)
                .GetString(data, 2, data.Length - 2).TrimStart('\uFEFF');
        if (data.Length >= 2 && data[0] == 0xFE && data[1] == 0xFF)
            return new UnicodeEncoding(true, false, true)
                .GetString(data, 2, data.Length - 2).TrimStart('\uFEFF');

        try
        {
            return new UTF8Encoding(false, true).GetString(data).TrimStart('\uFEFF');
        }
        catch (DecoderFallbackException)
        {
            // Not UTF-8; try UTF-16 without BOM, then GB18030.
        }
        if (LooksLikeUtf16(data, out var bigEndian))
        {
            try
            {
                var encoding = new UnicodeEncoding(bigEndian, false, true);
                return encoding.GetString(data).TrimStart('\uFEFF');
            }
            catch (DecoderFallbackException)
            {
                // fall through to GB18030
            }
        }
        try
        {
            return DecodeGb18030(data);
        }
        catch (Exception)
        {
            throw new NativeRuntimeException(
                "无法解码 Markdown 文件；仅支持 UTF-8、GB18030 或 UTF-16");
        }
    }

    private static bool LooksLikeUtf16(byte[] data, out bool bigEndian)
    {
        bigEndian = false;
        if (data.Length < 2 || data.Length % 2 != 0)
            return false;
        var limit = Math.Min(data.Length, 4096);
        var evenZeros = 0;
        var oddZeros = 0;
        for (var index = 0; index < limit; index++)
        {
            if (data[index] != 0)
                continue;
            if (index % 2 == 0)
                evenZeros++;
            else
                oddZeros++;
        }
        var threshold = Math.Max(2, limit / 8);
        if (oddZeros < threshold && evenZeros < threshold)
            return false;
        bigEndian = evenZeros > oddZeros;
        return true;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern int MultiByteToWideChar(
        uint codePage, uint flags, byte[] multiByteString, int multiByteLength,
        [Out] char[]? wideCharString, int wideCharLength);

    private static string DecodeGb18030(byte[] data)
    {
        const uint Gb18030CodePage = 54936;
        var needed = MultiByteToWideChar(
            Gb18030CodePage, 0, data, data.Length, null, 0);
        if (needed <= 0)
            throw new InvalidOperationException("GB18030 解码失败");
        var buffer = new char[needed];
        var written = MultiByteToWideChar(
            Gb18030CodePage, 0, data, data.Length, buffer, needed);
        if (written <= 0)
            throw new InvalidOperationException("GB18030 解码失败");
        return new string(buffer, 0, written);
    }

    private static List<SectionItem> MarkdownToSections(string markdown, string? sourceDirectory)
    {
        var sections = new List<SectionItem>();
        var paragraphLines = new List<string>();
        var listItems = new List<string>();
        string? listType = null;
        var tableLines = new List<string>();
        var codeLines = new List<string>();
        var inCodeBlock = false;

        void FlushParagraph()
        {
            var text = string.Join(" ", paragraphLines
                .Select(PlainText)
                .Where(value => !string.IsNullOrEmpty(value)));
            if (text.Length > 0)
                sections.Add(new SectionItem { Type = "paragraph", Text = text });
            paragraphLines.Clear();
        }

        void FlushList()
        {
            if (listItems.Count > 0 && listType is not null)
            {
                sections.Add(new SectionItem
                {
                    Type = listType,
                    Items = [.. listItems],
                });
            }
            listItems.Clear();
            listType = null;
        }

        void FlushTable()
        {
            var rows = new List<List<string>>();
            foreach (var line in tableLines)
            {
                if (TableDividerPattern.IsMatch(line))
                    continue;
                rows.Add(TableCells(line));
            }
            if (rows.Count > 0)
            {
                sections.Add(new SectionItem
                {
                    Type = "table",
                    Table = new TableDef
                    {
                        Headers = rows[0],
                        Rows = rows.Skip(1).ToList(),
                    },
                });
            }
            tableLines.Clear();
        }

        void FlushCode()
        {
            if (codeLines.Count > 0)
            {
                sections.Add(new SectionItem
                {
                    Type = "code",
                    Text = string.Join("\n", codeLines),
                });
            }
            codeLines.Clear();
        }

        var normalized = markdown.Replace("\r\n", "\n").Replace('\r', '\n');
        foreach (var rawLine in normalized.Split('\n'))
        {
            var line = rawLine.TrimEnd();
            if (line.StartsWith("```", StringComparison.Ordinal))
            {
                FlushParagraph();
                FlushList();
                FlushTable();
                if (inCodeBlock)
                    FlushCode();
                inCodeBlock = !inCodeBlock;
                continue;
            }
            if (inCodeBlock)
            {
                codeLines.Add(rawLine);
                continue;
            }
            if (line.TrimStart().StartsWith('|'))
            {
                FlushParagraph();
                FlushList();
                tableLines.Add(line);
                continue;
            }
            if (tableLines.Count > 0)
                FlushTable();

            var image = ImagePattern.Match(line);
            if (image.Success)
            {
                FlushParagraph();
                FlushList();
                var altText = image.Groups[1].Value;
                var target = image.Groups[2].Value;
                var imagePath = target;
                var imagePathValid = true;
                if (!string.IsNullOrEmpty(sourceDirectory))
                {
                    try
                    {
                        imagePath = Path.GetFullPath(Path.Combine(sourceDirectory, target));
                    }
                    catch (Exception ex) when (ex is ArgumentException or NotSupportedException
                                                   or IOException or UnauthorizedAccessException
                                                   or System.Security.SecurityException)
                    {
                        imagePathValid = false;
                    }
                }
                var suffix = Path.GetExtension(imagePath).ToLowerInvariant();
                if (imagePathValid
                    && File.Exists(imagePath)
                    && SupportedImageSuffixes.Contains(suffix, StringComparer.OrdinalIgnoreCase))
                {
                    sections.Add(new SectionItem
                    {
                        Type = "image",
                        Image = new ImageDef { Path = imagePath },
                    });
                }
                else
                {
                    var label = string.IsNullOrEmpty(altText) ? target : altText;
                    sections.Add(new SectionItem
                    {
                        Type = "paragraph",
                        Text = $"[图片未嵌入: {label}]",
                    });
                }
                continue;
            }

            var heading = HeadingPattern.Match(line);
            if (heading.Success)
            {
                FlushParagraph();
                FlushList();
                sections.Add(new SectionItem
                {
                    Type = "heading",
                    Level = Math.Min(heading.Groups[1].Value.Length, 6),
                    Text = PlainText(heading.Groups[2].Value),
                });
                continue;
            }
            if (HorizontalRulePattern.IsMatch(line))
            {
                FlushParagraph();
                FlushList();
                continue;
            }

            var quote = line.TrimStart();
            if (quote.StartsWith('>'))
                line = quote[1..].TrimStart();

            var bullet = BulletPattern.Match(line);
            var numbered = NumberedPattern.Match(line);
            if (bullet.Success || numbered.Success)
            {
                var itemType = bullet.Success ? "bullet_list" : "numbered_list";
                if (listType is not null && listType != itemType)
                    FlushList();
                FlushParagraph();
                listType = itemType;
                var match = bullet.Success ? bullet : numbered;
                listItems.Add(PlainText(match.Groups[1].Value));
                continue;
            }
            if (string.IsNullOrWhiteSpace(line))
            {
                FlushParagraph();
                FlushList();
                continue;
            }
            FlushList();
            paragraphLines.Add(line);
        }

        if (inCodeBlock)
            FlushCode();
        FlushTable();
        FlushParagraph();
        FlushList();
        return sections;
    }

    private static string PlainText(string value)
    {
        value = Regex.Replace(value, @"!\[([^\]]*)\]\([^)]*\)", "$1");
        value = Regex.Replace(value, @"\[([^\]]+)\]\([^)]*\)", "$1");
        value = Regex.Replace(value, @"\\text(?:bf|it)\{([^{}]*)\}", "$1");
        value = Regex.Replace(value, @"(```|`|\*\*|__|~~)", "");
        return value.Trim();
    }

    private static List<string> TableCells(string line) =>
        line.Trim().Trim('|').Split('|').Select(PlainText).ToList();
}
