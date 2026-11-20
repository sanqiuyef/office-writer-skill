using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.Json;
using PDFtoImage;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Actions;
using UglyToad.PdfPig.Annotations;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;
using UglyToad.PdfPig.Graphics.Colors;
using UglyToad.PdfPig.Outline;
using UglyToad.PdfPig.Outline.Destinations;

namespace OfficeNative;

/// <summary>
/// read-pdf-document / read-pdf-structured backed by PdfPig (Apache-2.0).
/// Field shapes follow scripts/office_engine/reader.py. Scanned pages fall back
/// to Tesseract OCR when the CLI is available (OFFICE_TESSERACT_CMD /
/// TESSERACT_CMD / PATH); otherwise the text layer is returned with a note.
/// </summary>
internal static class PdfFileReader
{
    private const string TruncationMarker = "_(内容已截断)_";
    private const string OcrUnavailable =
        "OCR 不可用（未找到 Tesseract），已返回 PDF 文本层";

    public static Dictionary<string, object?> RunReadPdfDocument(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        var maxChars = NativeCommon.OptInt(p, "max_chars", 50000);
        if (maxChars < 1)
            throw new NativeUsageException("max_chars 必须大于 0");
        ValidatePdfInput(filePath);
        var result = ReadDocument(filePath, maxChars);
        return new Dictionary<string, object?>
        {
            ["markdown"] = BuildMarkdown(result),
            ["result"] = result,
        };
    }

    public static Dictionary<string, object?> RunReadPdfStructured(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        var maxChars = NativeCommon.OptInt(p, "max_chars", 50000);
        if (maxChars < 1)
            throw new NativeUsageException("max_chars 必须大于 0");
        var includeBlocks = NativeCommon.OptBool(p, "include_blocks", false);
        var extractImages = NativeCommon.OptBool(p, "extract_images", false);
        ValidatePdfInput(filePath);
        return ReadStructured(filePath, maxChars, includeBlocks, extractImages);
    }

    private static void ValidatePdfInput(string filePath)
    {
        NativeCommon.EnsureFile(filePath);
        if (!Path.GetExtension(filePath).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            throw new NativeRuntimeException("不支持的文件格式，仅支持 .pdf");
    }

    // ── read-pdf-document ──────────────────────────────────────────

    private static Dictionary<string, object?> ReadDocument(string filePath, int maxChars)
    {
        try
        {
            using var document = PdfDocument.Open(filePath);
            var totalPages = document.NumberOfPages;
            var texts = ExtractPageTexts(document, out var rawTexts);
            var header = BuildHeader(document);
            string? note = null;
            if (NeedsOcr(string.Join("\n", rawTexts)))
            {
                if (TryOcrDocument(
                        filePath, totalPages, out var ocrTexts,
                        out var ocrMessage))
                {
                    texts = ocrTexts;
                    note = ocrMessage;
                }
                else
                {
                    note = ocrMessage;
                }
            }
            var pageParts = new List<string>();
            var truncated = false;

            for (var index = 0; index < texts.Count; index++)
            {
                var text = texts[index];
                var pageText = text.Length > 0
                    ? $"## Page {index + 1}\n\n{text}"
                    : $"## Page {index + 1}\n\n_(空页)_";
                var currentContent = string.Join("\n\n", pageParts);
                var separator = pageParts.Count > 0 ? "\n\n" : "";
                if (header.Length + currentContent.Length + separator.Length
                    + pageText.Length > maxChars)
                {
                    truncated = true;
                    break;
                }
                pageParts.Add(pageText);
            }

            var body = string.Join("\n\n", pageParts);
            var content = header + body;
            if (truncated)
            {
                var markerSuffix = $"\n\n{TruncationMarker}";
                if (content.Length + markerSuffix.Length <= maxChars)
                {
                    content += markerSuffix;
                }
                else if (pageParts.Count > 0)
                {
                    content = header + ContentWithTruncationMarker(
                        pageParts, maxChars - header.Length);
                }
                else
                {
                    content = header[..Math.Min(maxChars, header.Length)];
                }
            }

            var message = $"成功读取 PDF，共 {totalPages} 页";
            if (!string.IsNullOrEmpty(note))
                message += $"（{note}）";
            if (truncated)
                message += $"（已截断至 {maxChars} 字符）";
            return new Dictionary<string, object?>
            {
                ["success"] = true,
                ["file_path"] = filePath,
                ["content"] = content,
                ["pages_count"] = totalPages,
                ["truncated"] = truncated,
                ["message"] = message,
            };
        }
        catch (NativeRuntimeException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new NativeRuntimeException($"读取 PDF 失败: {ex.Message}");
        }
    }

    private static string BuildHeader(PdfDocument document)
    {
        var parts = new List<string>();
        var metadataLines = new List<string>();
        try
        {
            var info = document.Information;
            foreach (var (value, label) in new[]
                     {
                         (info.Title, "标题"),
                         (info.Author, "作者"),
                         (info.Subject, "主题"),
                         (info.Keywords, "关键词"),
                         (info.Creator, "创建工具"),
                         (info.Producer, "生成工具"),
                     })
            {
                if (!string.IsNullOrWhiteSpace(value))
                    metadataLines.Add($"{label}: {value.Trim()}");
            }
        }
        catch
        {
            // Malformed metadata must not abort text extraction.
        }
        if (metadataLines.Count > 0)
            parts.Add("--- 文档元数据 ---\n" + string.Join("\n", metadataLines));

        try
        {
            if (document.TryGetBookmarks(out var bookmarks))
            {
                var lines = new List<string> { "--- 书签 ---" };
                foreach (var node in bookmarks.GetNodes())
                {
                    var indent = new string(' ', Math.Max(0, node.Level) * 2);
                    var page = node is DocumentBookmarkNode destination
                        ? destination.PageNumber
                        : -1;
                    lines.Add($"{indent}- [{node.Title}](page {page})");
                }
                if (lines.Count > 1)
                    parts.Add(string.Join("\n", lines));
            }
        }
        catch
        {
            // Broken outline trees are ignored like PyMuPDF's tolerant get_toc().
        }

        return parts.Count > 0 ? string.Join("\n\n", parts) + "\n\n" : "";
    }

    private static string ContentWithTruncationMarker(List<string> parts, int maxChars)
    {
        if (maxChars < 1)
            return "";
        var content = string.Join("\n\n", parts);
        if (content.Length == 0)
            return TruncationMarker[..Math.Min(maxChars, TruncationMarker.Length)];
        var candidate = $"{content}\n\n{TruncationMarker}";
        return candidate.Length <= maxChars ? candidate : content;
    }

    private static string BuildMarkdown(Dictionary<string, object?> result)
    {
        if (result.TryGetValue("success", out var success) && success is false)
            return $"读取失败: {result.GetValueOrDefault("message") ?? "未知错误"}";

        var parts = new List<string>();
        if (result.GetValueOrDefault("content") is string content && content.Length > 0)
            parts.Add(content);
        var stats = new List<string>();
        foreach (var (key, value) in result)
        {
            if (key is "success" or "file_path" or "content" or "message" or "truncated")
                continue;
            if (value is null)
                continue;
            stats.Add($"{key}: {FormatStat(value)}");
        }
        if (stats.Count > 0)
            parts.Add("> " + string.Join("；", stats));
        if (result.GetValueOrDefault("truncated") is true)
            parts.Add("> 内容过长已截断");
        return string.Join("\n\n", parts.Where(part => part.Length > 0));
    }

    private static string FormatStat(object value) => value switch
    {
        bool boolean => boolean ? "True" : "False",
        IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture),
        _ => value.ToString() ?? "",
    };

    // ── read-pdf-structured ────────────────────────────────────────

    private static Dictionary<string, object?> ReadStructured(
        string filePath, int maxChars, bool includeBlocks, bool extractImages)
    {
        try
        {
            using var document = PdfDocument.Open(filePath);
            var totalPages = document.NumberOfPages;
            var metadata = ReadMetadata(document);
            var bookmarks = ReadBookmarks(document);
            var pages = new List<Dictionary<string, object?>>();
            var links = new List<Dictionary<string, object?>>();
            for (var index = 0; index < totalPages; index++)
            {
                try
                {
                    links.AddRange(ReadLinks(document.GetPage(index + 1), index + 1));
                }
                catch
                {
                    // A broken page's annotations must not hide the other pages.
                }
            }

            var imageDirectory = extractImages
                ? EnsureImageDirectory(filePath)
                : null;
            var textLength = 0;
            var truncated = false;
            var extractedImages = 0;

            for (var index = 0; index < totalPages; index++)
            {
                var page = document.GetPage(index + 1);
                var pageText = SafePageText(page);
                if (pageText.Length == 0)
                    pageText = "_(空页)_";
                textLength += pageText.Length;

                var entry = new Dictionary<string, object?>
                {
                    ["page"] = index + 1,
                    ["width"] = Round(page.Width),
                    ["height"] = Round(page.Height),
                    ["text"] = pageText,
                };

                if (includeBlocks)
                {
                    entry["blocks"] = ReadBlocks(
                        page, index + 1, page.Height, extractImages,
                        imageDirectory, ref extractedImages);
                }
                else if (extractImages)
                {
                    var imageBlocks = ExtractImageBlocks(
                        page, index + 1, page.Height, imageDirectory,
                        ref extractedImages);
                    if (imageBlocks.Count > 0)
                        entry["blocks"] = imageBlocks;
                }

                if (textLength > maxChars)
                {
                    truncated = true;
                    break;
                }
                pages.Add(entry);
            }

            var rawText = string.Join("\n", pages
                .Select(page => page.GetValueOrDefault("text") as string ?? ""));
            string? ocrNote = null;
            if (!truncated && NeedsOcr(rawText))
            {
                if (TryOcrDocument(
                        filePath, totalPages, out var ocrTexts,
                        out var ocrMessage))
                {
                    for (var index = 0; index < pages.Count && index < ocrTexts.Count; index++)
                    {
                        var pageText = ocrTexts[index].Trim();
                        pages[index]["text"] = pageText.Length > 0 ? pageText : "_(空页)_";
                    }
                    ocrNote = ocrMessage;
                }
                else
                {
                    ocrNote = ocrMessage;
                }
            }

            var message = $"成功读取 PDF，共 {totalPages} 页";
            if (extractedImages > 0)
                message += $"，提取 {extractedImages} 张图片到 {imageDirectory}";
            if (truncated)
                message += $"（已截断至 {maxChars} 字符）";
            if (includeBlocks)
                message += "，含文本块详细信息";
            if (!string.IsNullOrEmpty(ocrNote))
                message += $"（{ocrNote}）";

            return new Dictionary<string, object?>
            {
                ["success"] = true,
                ["file_path"] = filePath,
                ["pages_count"] = totalPages,
                ["message"] = message,
                ["metadata"] = metadata,
                ["bookmarks"] = bookmarks,
                ["links"] = links,
                ["pages"] = pages,
                ["truncated"] = truncated,
            };
        }
        catch (NativeRuntimeException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new NativeRuntimeException($"读取 PDF 失败: {ex.Message}");
        }
    }

    private static Dictionary<string, string> ReadMetadata(PdfDocument document)
    {
        var metadata = new Dictionary<string, string>();
        try
        {
            var info = document.Information;
            foreach (var (key, value) in new[]
                     {
                         ("title", info.Title),
                         ("author", info.Author),
                         ("subject", info.Subject),
                         ("keywords", info.Keywords),
                         ("creator", info.Creator),
                         ("producer", info.Producer),
                         ("creationDate", info.CreationDate),
                         ("modDate", info.ModifiedDate),
                     })
            {
                if (!string.IsNullOrWhiteSpace(value))
                    metadata[key] = value.Trim();
            }
        }
        catch
        {
            // Missing or malformed Info dictionary: report nothing.
        }
        return metadata;
    }

    private static List<Dictionary<string, object?>> ReadBookmarks(PdfDocument document)
    {
        var result = new List<Dictionary<string, object?>>();
        try
        {
            if (!document.TryGetBookmarks(out var bookmarks))
                return result;
            foreach (var node in bookmarks.GetNodes())
            {
                result.Add(new Dictionary<string, object?>
                {
                    ["level"] = Math.Max(0, node.Level) + 1,
                    ["title"] = node.Title,
                    ["page"] = node is DocumentBookmarkNode destination
                        ? destination.PageNumber
                        : -1,
                });
            }
        }
        catch
        {
            // Ignore a broken outline; text extraction remains useful.
        }
        return result;
    }

    private static List<Dictionary<string, object?>> ReadLinks(Page page, int pageNumber)
    {
        var links = new List<Dictionary<string, object?>>();
        try
        {
            foreach (var annotation in page.GetAnnotations())
            {
                if (annotation.Type != AnnotationType.Link)
                    continue;
                var entry = new Dictionary<string, object?> { ["page"] = pageNumber };
                switch (annotation.Action)
                {
                    case UriAction uri:
                        entry["kind"] = 2;
                        entry["uri"] = uri.Uri;
                        break;
                    case AbstractGoToAction goTo:
                        entry["kind"] = 1;
                        entry["dest_page"] = Math.Max(
                            0, goTo.Destination.PageNumber - 1);
                        break;
                    default:
                        entry["kind"] = 0;
                        break;
                }
                entry["bbox"] = ToBbox(annotation.Rectangle, page.Height);
                links.Add(entry);
            }
        }
        catch
        {
            // Annotations are optional metadata; never fail the whole read.
        }
        return links;
    }

    private static List<Dictionary<string, object?>> ReadBlocks(
        Page page, int pageNumber, double pageHeight, bool extractImages,
        string? imageDirectory, ref int extractedImages)
    {
        var blocks = new List<Dictionary<string, object?>>();
        try
        {
            // PyMuPDF exposes one block per baseline row, with table cells as
            // separate line entries inside that block. Mirror that shape.
            foreach (var group in GroupWordsByBaseline(page.GetWords()))
            {
                var words = group.Where(word => word.Text.Length > 0).ToList();
                if (words.Count == 0)
                    continue;
                var left = words.Min(word => word.BoundingBox.Left);
                var right = words.Max(word => word.BoundingBox.Right);
                var bottom = words.Min(word => word.BoundingBox.Bottom);
                var top = words.Max(word => word.BoundingBox.Top);
                var entry = new Dictionary<string, object?>
                {
                    ["type"] = "text",
                    ["bbox"] = ToBbox(
                        new PdfRectangle(new PdfPoint(left, bottom),
                            new PdfPoint(right, top)),
                        pageHeight),
                };
                var lines = new List<Dictionary<string, object?>>();
                foreach (var segment in SplitColumnWordGroups(words))
                {
                    var spans = BuildSpans(segment);
                    if (spans.Count > 0)
                        lines.Add(new Dictionary<string, object?> { ["spans"] = spans });
                }
                entry["lines"] = lines;
                blocks.Add(entry);
            }
        }
        catch
        {
            // Text segmentation is best-effort; page.text still carries the content.
        }

        if (extractImages)
            blocks.AddRange(ExtractImageBlocks(
                page, pageNumber, pageHeight, imageDirectory, ref extractedImages));
        else
            blocks.AddRange(ReadImageBlocks(page, pageHeight));

        return blocks;
    }

    private static List<Dictionary<string, object?>> BuildSpans(List<Word> words)
    {
        var spans = new List<Dictionary<string, object?>>();
        var currentWords = new List<Word>();
        string? currentFont = null;
        double currentSize = 0;
        int currentColor = 0;

        void Flush()
        {
            if (currentWords.Count == 0)
                return;
            spans.Add(new Dictionary<string, object?>
            {
                ["text"] = JoinWords(currentWords),
                ["font"] = currentFont ?? "",
                ["size"] = Math.Round(currentSize, 1),
                ["color"] = currentColor,
            });
            currentWords.Clear();
        }

        foreach (var word in words)
        {
            var letter = word.Letters.FirstOrDefault();
            var font = StripSubset(word.FontName);
            var size = Math.Round(letter?.FontSize ?? 0, 1);
            var color = ColorValue(letter?.Color);
            if (currentWords.Count > 0
                && (font != currentFont || Math.Abs(size - currentSize) > 0.05
                    || color != currentColor))
                Flush();
            currentFont = font;
            currentSize = size;
            currentColor = color;
            currentWords.Add(word);
        }
        Flush();
        return spans;
    }

    private static string JoinWords(List<Word> words)
    {
        var builder = new StringBuilder();
        for (var index = 0; index < words.Count; index++)
        {
            var word = words[index];
            if (index > 0)
            {
                var previous = words[index - 1];
                var previousText = previous.Text;
                var text = word.Text;
                var previousChar = previousText.Length > 0 ? previousText[^1] : '\0';
                var nextChar = text.Length > 0 ? text[0] : '\0';
                var gap = word.BoundingBox.Left - previous.BoundingBox.Right;
                var fontSize = word.Letters.FirstOrDefault()?.FontSize
                               ?? previous.Letters.FirstOrDefault()?.FontSize
                               ?? 10;
                var spaced = gap > Math.Max(1.0, fontSize * 0.12)
                             || (!IsCjk(previousChar) && !IsCjk(nextChar));
                if (spaced)
                    builder.Append(' ');
            }
            builder.Append(word.Text);
        }
        return builder.ToString();
    }

    private static List<Dictionary<string, object?>> ExtractImageBlocks(
        Page page, int pageNumber, double pageHeight, string? imageDirectory,
        ref int extractedImages)
    {
        var blocks = new List<Dictionary<string, object?>>();
        try
        {
            var images = page.GetImages().ToList();
            for (var index = 0; index < images.Count; index++)
            {
                var image = images[index];
                var entry = new Dictionary<string, object?>
                {
                    ["type"] = "image",
                    ["bbox"] = ToBbox(image.Bounds, pageHeight),
                };
                if (imageDirectory is not null && image.TryGetPng(out var png))
                {
                    var path = Path.Combine(
                        imageDirectory, $"page{pageNumber}_img{index}.png");
                    File.WriteAllBytes(path, png);
                    entry["saved_path"] = path;
                    entry["image_width"] = image.WidthInSamples;
                    entry["image_height"] = image.HeightInSamples;
                    extractedImages++;
                }
                blocks.Add(entry);
            }
        }
        catch
        {
            // Damaged image streams are skipped, matching PyMuPDF's per-xref try/except.
        }
        return blocks;
    }

    /// <summary>Blocks returned when include_blocks=true but extraction is off.</summary>
    private static List<Dictionary<string, object?>> ReadImageBlocks(
        Page page, double pageHeight)
    {
        var blocks = new List<Dictionary<string, object?>>();
        try
        {
            foreach (var image in page.GetImages())
            {
                blocks.Add(new Dictionary<string, object?>
                {
                    ["type"] = "image",
                    ["bbox"] = ToBbox(image.Bounds, pageHeight),
                });
            }
        }
        catch
        {
            // Ignore broken image dictionaries.
        }
        return blocks;
    }

    private static List<string> ExtractPageTexts(
        PdfDocument document, out List<string> rawTexts)
    {
        var texts = new List<string>();
        rawTexts = new List<string>();
        for (var index = 1; index <= document.NumberOfPages; index++)
        {
            string raw;
            try
            {
                var page = document.GetPage(index);
                raw = ExtractPageText(page);
            }
            catch
            {
                raw = "";
            }
            rawTexts.Add(raw);
            texts.Add(raw.Trim());
        }
        return texts;
    }

    private static string SafePageText(Page page)
    {
        try
        {
            return ExtractPageText(page).Trim();
        }
        catch
        {
            return "";
        }
    }

    private static string ExtractPageText(Page page)
    {
        try
        {
            return BuildTextFromWords(page);
        }
        catch
        {
            try
            {
                return ContentOrderTextExtractor.GetText(page, false);
            }
            catch
            {
                return page.Text ?? "";
            }
        }
    }

    /// <summary>
    /// PyMuPDF-like text layout: words are grouped by baseline into visual lines,
    /// then split into segments when a large horizontal gap separates columns
    /// (table cells). Every segment becomes one output line, so a table row is
    /// emitted cell by cell instead of being flattened into one line.
    /// </summary>
    private static string BuildTextFromWords(Page page)
    {
        var groups = GroupWordsByBaseline(page.GetWords());
        if (groups.Count == 0)
            return "";
        var segments = new List<string>();
        foreach (var group in groups)
            segments.AddRange(SplitColumns(group));
        return string.Join("\n", segments)
            .Replace("\r\n", "\n").Replace('\r', '\n');
    }

    private static List<List<Word>> GroupWordsByBaseline(IEnumerable<Word> words)
    {
        var ordered = words
            .OrderByDescending(word => word.BoundingBox.Bottom)
            .ThenBy(word => word.BoundingBox.Left)
            .ToList();
        var groups = new List<List<Word>>();
        double baseline = 0;
        foreach (var word in ordered)
        {
            var bottom = word.BoundingBox.Bottom;
            var tolerance = Math.Max(
                1.5,
                (word.Letters.FirstOrDefault()?.FontSize ?? 10) * 0.25);
            if (groups.Count == 0 || Math.Abs(bottom - baseline) > tolerance)
            {
                groups.Add([word]);
                baseline = bottom;
            }
            else
            {
                groups[^1].Add(word);
            }
        }
        foreach (var group in groups)
            group.Sort((left, right) => left.BoundingBox.Left.CompareTo(right.BoundingBox.Left));
        return groups;
    }

    private static List<string> SplitColumns(List<Word> line)
    {
        var segments = new List<string>();
        foreach (var group in SplitColumnWordGroups(line))
            segments.Add(JoinWords(group));
        return segments;
    }

    private static List<List<Word>> SplitColumnWordGroups(List<Word> line)
    {
        var segments = new List<List<Word>>();
        var current = new List<Word>();
        foreach (var word in line)
        {
            if (current.Count > 0)
            {
                var previous = current[^1];
                var fontSize = word.Letters.FirstOrDefault()?.FontSize
                               ?? previous.Letters.FirstOrDefault()?.FontSize
                               ?? 10;
                var gap = word.BoundingBox.Left - previous.BoundingBox.Right;
                if (gap > Math.Max(8, fontSize * 2.5))
                {
                    segments.Add(current);
                    current = [];
                }
            }
            current.Add(word);
        }
        if (current.Count > 0)
            segments.Add(current);
        return segments;
    }

    private static string EnsureImageDirectory(string filePath)
    {
        var full = Path.GetFullPath(filePath);
        var directory = Path.GetDirectoryName(full)!;
        var stem = Path.GetFileNameWithoutExtension(full);
        var imageDirectory = Path.Combine(directory, $"{stem}_images");
        Directory.CreateDirectory(imageDirectory);
        return imageDirectory;
    }

    // ── primitive helpers ──────────────────────────────────────────

    private static List<double> ToBbox(PdfRectangle rectangle, double pageHeight) =>
    [
        Round(rectangle.Left),
        Round(pageHeight - rectangle.Top),
        Round(rectangle.Right),
        Round(pageHeight - rectangle.Bottom),
    ];

    private static double Round(double value) =>
        Math.Round(value, 2, MidpointRounding.AwayFromZero);

    private static int ColorValue(IColor? color)
    {
        if (color is null)
            return 0;
        try
        {
            var (red, green, blue) = color.ToRGBValues();
            return (ToByte(red) << 16) | (ToByte(green) << 8) | ToByte(blue);
        }
        catch
        {
            return 0;
        }
    }

    private static int ToByte(double value) =>
        Math.Clamp((int)Math.Round(value * 255, MidpointRounding.AwayFromZero), 0, 255);

    private static string StripSubset(string? fontName)
    {
        if (string.IsNullOrEmpty(fontName))
            return "";
        if (fontName.Length > 7 && fontName[6] == '+'
            && fontName.Take(6).All(char.IsAsciiLetterUpper))
            return fontName[7..];
        return fontName;
    }

    private static bool IsCjk(char c) =>
        (c >= 0x2E80 && c <= 0x9FFF)
        || (c >= 0xAC00 && c <= 0xD7AF)
        || (c >= 0xF900 && c <= 0xFAFF)
        || (c >= 0xFE30 && c <= 0xFE4F)
        || (c >= 0xFF00 && c <= 0xFFEF);

    /// <summary>
    /// Rasterizes every page with PDFium and runs the Tesseract CLI. Mirrors the
    /// Python engine's OCR fallback (OFFICE_TESSERACT_CMD / TESSERACT_CMD / PATH,
    /// MCP_OFFICE_OCR_LANG, MCP_OFFICE_OCR_DPI, MCP_OFFICE_OCR_MAX_PAGES, --psm 6).
    /// </summary>
    private static bool TryOcrDocument(
        string filePath, int totalPages, out List<string> pageTexts, out string message)
    {
        pageTexts = [];
        var executable = ResolveTesseract();
        if (executable is null)
        {
            message = OcrUnavailable;
            return false;
        }

        var language = Environment.GetEnvironmentVariable("MCP_OFFICE_OCR_LANG");
        if (string.IsNullOrWhiteSpace(language))
            language = "chi_sim+eng";
        var dpi = Math.Clamp(ReadIntEnvironment("MCP_OFFICE_OCR_DPI", 300), 50, 600);
        var maxPages = ReadIntEnvironment("MCP_OFFICE_OCR_MAX_PAGES", 0);
        var limit = maxPages > 0 ? Math.Min(totalPages, maxPages) : totalPages;

        var workDirectory = Path.Combine(
            Path.GetTempPath(), $"office-native-ocr-{Guid.NewGuid():N}");
        Directory.CreateDirectory(workDirectory);
        try
        {
            var bytes = File.ReadAllBytes(filePath);
            var results = new List<string>(limit);
            for (var page = 1; page <= limit; page++)
            {
                var pngPath = Path.Combine(workDirectory, $"page-{page:0000}.png");
                Conversion.SavePng(
                    pngPath, bytes, new Index(page - 1), null,
                    new RenderOptions { Dpi = dpi });
                var pageText = RunTesseract(executable, pngPath, language);
                if (pageText is null)
                {
                    message = $"OCR 失败（Tesseract 处理第 {page} 页出错），已返回 PDF 文本层";
                    return false;
                }
                results.Add(pageText.Trim());
            }
            if (results.All(string.IsNullOrWhiteSpace))
            {
                message = "OCR 未识别到文本，已返回 PDF 文本层";
                return false;
            }
            if (limit < totalPages)
            {
                results[^1] +=
                    $"\n\n_[OCR 仅处理前 {limit}/{totalPages} 页，"
                    + "受 MCP_OFFICE_OCR_MAX_PAGES 限制]_";
            }
            pageTexts = results;
            message = limit < totalPages
                ? $"已执行 OCR（{language}，前 {limit}/{totalPages} 页）"
                : $"已执行 OCR（{language}）";
            return true;
        }
        catch (Exception ex)
        {
            // OCR 失败不阻断 PDF 读取，降级返回文本层（与 Python 引擎一致）。
            message = $"OCR 失败（{ex.Message}），已返回 PDF 文本层";
            return false;
        }
        finally
        {
            try { Directory.Delete(workDirectory, recursive: true); }
            catch { /* best-effort cleanup */ }
        }
    }

    private static string? ResolveTesseract()
    {
        foreach (var name in new[] { "OFFICE_TESSERACT_CMD", "TESSERACT_CMD" })
        {
            var configured = Environment.GetEnvironmentVariable(name);
            if (!string.IsNullOrWhiteSpace(configured))
                return File.Exists(configured) ? configured : null;
        }

        var path = Environment.GetEnvironmentVariable("PATH");
        if (string.IsNullOrWhiteSpace(path))
            return null;
        foreach (var directory in path.Split(Path.PathSeparator))
        {
            if (string.IsNullOrWhiteSpace(directory))
                continue;
            foreach (var candidate in new[] { "tesseract.exe", "tesseract" })
            {
                var fullPath = Path.Combine(directory.Trim(), candidate);
                if (File.Exists(fullPath))
                    return fullPath;
            }
        }
        return null;
    }

    private static string? RunTesseract(
        string executable, string imagePath, string language)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = executable,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
            StandardOutputEncoding = new UTF8Encoding(false),
            StandardErrorEncoding = new UTF8Encoding(false),
        };
        startInfo.ArgumentList.Add(imagePath);
        startInfo.ArgumentList.Add("stdout");
        startInfo.ArgumentList.Add("-l");
        startInfo.ArgumentList.Add(language);
        startInfo.ArgumentList.Add("--psm");
        startInfo.ArgumentList.Add("6");

        using var process = Process.Start(startInfo);
        if (process is null)
            return null;
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(120_000))
        {
            try { process.Kill(entireProcessTree: true); }
            catch { /* already exited */ }
            return null;
        }
        var text = stdout.GetAwaiter().GetResult();
        _ = stderr.GetAwaiter().GetResult();
        return process.ExitCode == 0 ? text : null;
    }

    private static bool NeedsOcr(string text)
    {        if (Environment.GetEnvironmentVariable("MCP_OFFICE_OCR") == "0")
            return false;
        var minChars = ReadIntEnvironment("MCP_OFFICE_OCR_MIN_TEXT_CHARS", 500);
        var lines = text.Split('\n')
            .Select(line => new string(line.Where(c => !char.IsWhiteSpace(c)).ToArray()))
            .Where(line => line.Length >= 3)
            .ToList();
        var compact = string.Concat(lines);
        if (compact.Length < minChars)
            return lines.Count == 0;
        if (lines.Count >= 8)
        {
            var groups = lines.GroupBy(line => line)
                .Select(group => group.Count())
                .ToList();
            var repeated = groups.Max();
            if (groups.Count <= Math.Max(3, lines.Count / 8)
                || repeated / (double)lines.Count >= 0.85)
                return true;
        }
        return false;
    }

    private static int ReadIntEnvironment(string name, int fallback) =>
        int.TryParse(Environment.GetEnvironmentVariable(name), out var value)
            ? value
            : fallback;
}
