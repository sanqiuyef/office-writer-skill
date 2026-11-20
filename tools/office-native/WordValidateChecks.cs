using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeNative;

/// <summary>
/// validate-word-document 的检查实现：逐项端口 scripts/office_engine/checker.py
/// 的 validate_word_document。check id、severity 与 message/details 尽量保持与
/// Python 逐字一致；段落/run 的文本语义按 python-docx 处理（见各辅助方法）。
/// </summary>
internal static class WordValidateChecks
{
    private const string WordNamespace =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private const long MaxInputBytes = 100L * 1024 * 1024;
    private const long MaxUncompressedBytes = 500L * 1024 * 1024;

    private static readonly Regex HeadingStylePattern =
        new(@"(?:Heading|heading)\s*(\d+)", RegexOptions.CultureInvariant);

    private static readonly Regex PlaceholderPattern =
        new(@"\b(?:TODO|TBD|DRAFT)\b|待(?:完善|补充|确认)",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex PageFieldPattern =
        new(@"\bPAGE\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex TocFieldPattern =
        new(@"\bTOC\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex CjkPattern =
        new(@"[\u3400-\u9fff\uf900-\ufaff]", RegexOptions.CultureInvariant);

    private static readonly Regex LatinPattern =
        new(@"[A-Za-z]", RegexOptions.CultureInvariant);

    private static readonly Regex DecimalHeadingPattern =
        new(@"^(\d+(?:\.\d+)*)(?:\s*[.、．)）]|\s|$)", RegexOptions.CultureInvariant);

    private static readonly Regex CnParenHeadingPattern =
        new(@"^[（(]([一二三四五六七八九十百零〇两]+)[)）]", RegexOptions.CultureInvariant);

    private static readonly Regex CnChapterHeadingPattern =
        new(@"^第([一二三四五六七八九十百零〇两]+)[章节篇部分]", RegexOptions.CultureInvariant);

    private static readonly Regex CnEnumHeadingPattern =
        new(@"^([一二三四五六七八九十百零〇两]+)[、.．]\s*", RegexOptions.CultureInvariant);

    private static readonly Dictionary<char, int> CnDigits = new()
    {
        ['零'] = 0, ['〇'] = 0, ['一'] = 1, ['二'] = 2, ['两'] = 2, ['三'] = 3,
        ['四'] = 4, ['五'] = 5, ['六'] = 6, ['七'] = 7, ['八'] = 8, ['九'] = 9,
    };

    private static readonly (string Attribute, bool IsHeader, string Type)[] StoryAttributes =
    [
        ("header", true, "default"),
        ("footer", false, "default"),
        ("first_page_header", true, "first"),
        ("first_page_footer", false, "first"),
        ("even_page_header", true, "even"),
        ("even_page_footer", false, "even"),
    ];

    public static WordValidateReport Validate(string filePath)
    {
        var report = new WordValidateReport(filePath);

        if (!File.Exists(filePath))
        {
            report.Add("文件存在", false, "error", "文件不存在", checkId: "file_exists");
            return report;
        }

        var safetyError = ValidateInputFile(filePath);
        if (safetyError is not null)
        {
            report.Add("输入文件安全检查", false, "error", safetyError, checkId: "input_safe");
            return report;
        }

        WordprocessingDocument? document = null;
        try
        {
            document = WordprocessingDocument.Open(filePath, false);
            RunChecks(report, filePath, document);
        }
        catch (Exception ex)
        {
            report.Add("文档解析", false, "error", $"无法解析文档: {ex.Message}", checkId: "parse");
        }
        finally
        {
            document?.Dispose();
        }
        return report;
    }

    /// <summary>端口 utils.validate_input_file(file_path, (".docx",))。</summary>
    private static string? ValidateInputFile(string filePath)
    {
        var extension = Path.GetExtension(filePath);
        if (!string.Equals(extension, ".docx", StringComparison.OrdinalIgnoreCase))
            return "不支持的文件格式，仅支持: .docx";

        long fileSize;
        try
        {
            fileSize = new FileInfo(filePath).Length;
        }
        catch (Exception ex)
        {
            return ex.Message;
        }
        if (fileSize > MaxInputBytes)
            return "文件超过 100 MB 安全读取上限";

        try
        {
            using var archive = ZipFile.OpenRead(filePath);
            long uncompressed = 0;
            foreach (var entry in archive.Entries)
                uncompressed += entry.Length;
            if (uncompressed > MaxUncompressedBytes)
                return "Office 文件解压后超过 500 MB 安全上限";
        }
        catch (InvalidDataException)
        {
            return "Office 文件损坏或不是有效的 OOXML 压缩包";
        }
        catch (IOException)
        {
            return "Office 文件损坏或不是有效的 OOXML 压缩包";
        }
        return null;
    }

    private static void RunChecks(
        WordValidateReport report, string filePath, WordprocessingDocument document)
    {
        var mainPart = document.MainDocumentPart;
        var body = mainPart?.Document?.Body
                   ?? throw new InvalidOperationException("文档没有正文");
        var sections = CollectSections(body);
        var styleResolver = new StyleResolver(mainPart);
        var storyRoots = CollectStoryRoots(mainPart, body, sections);
        var bodyParagraphs = body.Elements<W.Paragraph>().ToList();
        var bodyTables = body.Elements<W.Table>().ToList();

        // ── 检查 1: 文件基础信息 ──
        var fileSize = new FileInfo(filePath).Length;
        var sizeOk = fileSize > 0;
        report.Add("文件大小检查", sizeOk, sizeOk ? "info" : "error",
            sizeOk ? $"{fileSize} 字节" : "空文件", checkId: "file_size");

        report.Add("段落数量", true, "info", $"{bodyParagraphs.Count} 个段落",
            checkId: "paragraph_count");

        // ── 检查 2: 空白页检查 ──
        var emptyParaCount = bodyParagraphs.Count(
            paragraph => ParagraphText(paragraph).Trim().Length == 0);
        var pageBreakCount = bodyParagraphs.Sum(
            paragraph => paragraph.Elements<W.Run>().Count(HasPageBreak));
        if (emptyParaCount > 5 && pageBreakCount > 2)
        {
            report.Add("空白页检测", false, "warn",
                $"检测到 {emptyParaCount} 个空段落和 {pageBreakCount} 个分页符，可能存在空白页",
                $"空段落: {emptyParaCount}, 分页符: {pageBreakCount}",
                "blank_pages");
        }
        else
        {
            report.Add("空白页检测", true, "info",
                $"空段落: {emptyParaCount}, 分页符: {pageBreakCount}",
                checkId: "blank_pages");
        }

        // ── 检查 3: 行间距一致性 ──
        var lineSpacings = new SortedSet<double>();
        foreach (var paragraph in bodyParagraphs)
        {
            if (ParagraphText(paragraph).Trim().Length == 0)
                continue;
            var lineSpacing = SafeLineSpacing(paragraph);
            if (lineSpacing is not null)
                lineSpacings.Add(lineSpacing.Value);
        }
        if (lineSpacings.Count > 3)
        {
            report.Add("行间距一致性", false, "warn",
                $"检测到 {lineSpacings.Count} 种不同行间距: {FormatNumberList(lineSpacings)}",
                checkId: "line_spacing");
        }
        else
        {
            report.Add("行间距一致性", true, "info",
                $"行间距: {(lineSpacings.Count > 0 ? FormatNumberList(lineSpacings) : "默认")}",
                checkId: "line_spacing");
        }

        // ── 检查 4: 表格完整性 ──
        var tableIssues = 0;
        foreach (var table in bodyTables)
        {
            var rows = table.Elements<W.TableRow>().ToList();
            if (rows.Count == 0)
            {
                tableIssues++;
                continue;
            }
            var hasContent = rows.Any(row => row.Elements<W.TableCell>()
                .Any(cell => TableCellText(cell).Trim().Length > 0));
            if (!hasContent)
                tableIssues++;
        }
        if (tableIssues > 0)
        {
            report.Add("表格完整性", false, "warn",
                $"共 {bodyTables.Count} 个表格，{tableIssues} 个可能存在空表问题",
                checkId: "table_integrity");
        }
        else
        {
            report.Add("表格完整性", true, "info", $"共 {bodyTables.Count} 个表格",
                checkId: "table_integrity");
        }

        // ── 检查 5: 标题层级连续（报告所有跳跃，不只是第一个）──
        var headingLevels = new List<(int Level, string Text)>();
        foreach (var paragraph in bodyParagraphs)
        {
            var level = HeadingLevel(paragraph, styleResolver);
            if (level is not null)
                headingLevels.Add((level.Value, Truncate(ParagraphText(paragraph).Trim(), 40)));
        }

        if (headingLevels.Count > 0)
        {
            var jumpsFound = false;
            for (var index = 1; index < headingLevels.Count; index++)
            {
                var previousLevel = headingLevels[index - 1].Level;
                var currentLevel = headingLevels[index].Level;
                if (currentLevel > previousLevel + 1)
                {
                    report.Add("标题层级连续性", false, "warn",
                        $"标题从 H{previousLevel} 跳转到 H{currentLevel}（第 {index + 1} 个标题）",
                        checkId: "heading_levels");
                    jumpsFound = true;
                }
            }
            if (!jumpsFound)
            {
                var levels = headingLevels.Select(item => item.Level).ToList();
                report.Add("标题层级连续性", true, "info",
                    $"标题层级: H{levels.Min()}-H{levels.Max()}",
                    checkId: "heading_levels");
            }
        }
        else
        {
            report.Add("标题层级连续性", true, "info", "无标题（可能为纯文档）",
                checkId: "heading_levels");
        }

        // ── 检查 6: 图片嵌入检测 ──
        var imageCount = bodyParagraphs.Sum(
            paragraph => paragraph.Elements<W.Run>().Count(HasDrawing));
        report.Add("图片嵌入检查", true, "info", $"检测到 {imageCount} 个图片",
            checkId: "images");

        // ── 检查 7: section/分节设置 ──
        if (sections.Count > 1)
            report.Add("分节设置", true, "info", $"共 {sections.Count} 个节", checkId: "sections");
        else
            report.Add("分节设置", true, "info", "单节文档", checkId: "sections");

        // ── 检查 8: 页眉页脚 ──
        var headerFooterFound = false;
        for (var index = 0; index < sections.Count && !headerFooterFound; index++)
        {
            var headerPart = ResolveHeader(mainPart, sections, index, "default");
            if (headerPart?.Header is { } header &&
                header.Elements<W.Paragraph>()
                    .Any(paragraph => ParagraphText(paragraph).Trim().Length > 0))
            {
                headerFooterFound = true;
                break;
            }
            var footerPart = ResolveFooter(mainPart, sections, index, "default");
            if (footerPart?.Footer is { } footer &&
                footer.Elements<W.Paragraph>()
                    .Any(paragraph => ParagraphText(paragraph).Trim().Length > 0))
            {
                headerFooterFound = true;
                break;
            }
        }
        if (headerFooterFound)
            report.Add("页眉页脚", true, "info", "已检测到页眉页脚内容",
                checkId: "header_footer");
        else
            report.Add("页眉页脚", true, "info", "无页眉页脚（可能是封面/简历等）",
                checkId: "header_footer");

        // ── 检查 9: 遗留占位标记 ──
        var textParts = new List<string>();
        foreach (var paragraph in bodyParagraphs)
            textParts.Add(ParagraphText(paragraph));
        foreach (var table in bodyTables)
        {
            foreach (var row in table.Elements<W.TableRow>())
            {
                foreach (var cell in row.Elements<W.TableCell>())
                    textParts.Add(TableCellText(cell));
            }
        }
        for (var index = 0; index < sections.Count; index++)
        {
            var headerPart = ResolveHeader(mainPart, sections, index, "default");
            if (headerPart?.Header is { } header)
                textParts.AddRange(header.Elements<W.Paragraph>().Select(ParagraphText));
            var footerPart = ResolveFooter(mainPart, sections, index, "default");
            if (footerPart?.Footer is { } footer)
                textParts.AddRange(footer.Elements<W.Paragraph>().Select(ParagraphText));
        }
        var markers = new SortedSet<string>(StringComparer.Ordinal);
        foreach (var text in textParts)
        {
            foreach (Match match in PlaceholderPattern.Matches(text))
                markers.Add(match.Value);
        }
        if (markers.Count > 0)
        {
            report.Add("遗留占位标记", false, "warn", $"发现: {string.Join(", ", markers)}",
                checkId: "placeholders");
        }
        else
        {
            report.Add("遗留占位标记", true, "info",
                "未发现 TODO/TBD/DRAFT 或待完善标记", checkId: "placeholders");
        }

        // ── 检查 10: 字体是否本机已安装 ──
        Guarded(report, "字体安装检查", "font_installed", "字体检查失败", () =>
        {
            var explicitFonts = CollectExplicitFonts(storyRoots);
            if (!OperatingSystem.IsWindows())
            {
                report.Add("字体安装检查", true, "info",
                    "非 Windows 环境，跳过字体安装检查", checkId: "font_installed");
            }
            else if (explicitFonts.Count == 0)
            {
                report.Add("字体安装检查", true, "info",
                    "未显式指定字体（继承样式/主题）", checkId: "font_installed");
            }
            else
            {
                var installed = WordValidateFonts.InstalledKeys();
                var missing = explicitFonts
                    .Where(font => !installed.Contains(WordValidateFonts.Key(font)))
                    .ToList();
                if (missing.Count > 0)
                {
                    report.Add("字体安装检查", false, "warn",
                        $"字体未安装: {string.Join(", ", missing)}",
                        $"文档显式引用 {explicitFonts.Count} 种字体: " +
                        $"{string.Join(", ", explicitFonts)}",
                        "font_installed");
                }
                else
                {
                    report.Add("字体安装检查", true, "info",
                        $"字体均已安装（{explicitFonts.Count} 种）",
                        checkId: "font_installed");
                }
            }
        });

        // ── 检查 11: 标题编号连续性 ──
        Guarded(report, "标题编号连续性", "heading_numbering", "标题编号检查失败", () =>
        {
            var (numberingIssues, numberedCount) = HeadingNumberingIssues(headingLevels);
            if (numberingIssues.Count > 0)
            {
                report.Add("标题编号连续性", false, "warn",
                    $"标题编号不连续（{numberingIssues.Count} 处）",
                    string.Join("；", numberingIssues.Take(8)),
                    "heading_numbering");
            }
            else if (numberedCount > 0)
            {
                report.Add("标题编号连续性", true, "info",
                    $"标题编号连续（{numberedCount} 个带编号标题）",
                    checkId: "heading_numbering");
            }
            else
            {
                report.Add("标题编号连续性", true, "info", "未检测到标题编号",
                    checkId: "heading_numbering");
            }
        });

        // ── 检查 12: PAGE 页码域 ──
        Guarded(report, "页码域检查", "page_number_field", "页码域检查失败", () =>
        {
            var instructions = FieldInstructions(storyRoots);
            var hasPageField = instructions.Any(instruction => PageFieldPattern.IsMatch(instruction));
            if (hasPageField)
            {
                report.Add("页码域检查", true, "info", "检测到 PAGE 域",
                    checkId: "page_number_field");
            }
            else
            {
                report.Add("页码域检查", false, "warn",
                    "未检测到 PAGE 域（页码可能为静态文本或缺失）",
                    checkId: "page_number_field");
            }
        });

        // ── 检查 13: TOC 目录域 ──
        Guarded(report, "目录域检查", "toc_field", "目录域检查失败", () =>
        {
            var instructions = FieldInstructions(storyRoots);
            var hasTocField = instructions.Any(instruction => TocFieldPattern.IsMatch(instruction));
            var headingCount = headingLevels.Count;
            if (hasTocField)
            {
                if (TocMarkedForUpdate(mainPart, storyRoots))
                {
                    report.Add("目录域检查", true, "info", "目录域存在且标记为需更新",
                        checkId: "toc_field");
                }
                else
                {
                    report.Add("目录域检查", false, "warn",
                        "目录域存在但未标记为需更新（dirty/updateFields）",
                        "Word 打开后不会自动刷新目录",
                        "toc_field");
                }
            }
            else if (headingCount >= 3)
            {
                report.Add("目录域检查", false, "warn", "未检测到目录域（TOC）",
                    $"文档有 {headingCount} 个标题；长文档建议插入 TOC 域",
                    "toc_field");
            }
            else
            {
                report.Add("目录域检查", true, "info",
                    $"标题较少（{headingCount} 个），未要求目录域",
                    checkId: "toc_field");
            }
        });

        // ── 检查 14: 表格列数与合并单元格结构 ──
        Guarded(report, "表格列数/合并结构", "table_columns", "表格结构检查失败", () =>
        {
            var columnIssues = new List<string>();
            for (var index = 0; index < bodyTables.Count; index++)
                columnIssues.AddRange(TableColumnIssues(bodyTables[index], index + 1));
            if (columnIssues.Count > 0)
            {
                report.Add("表格列数/合并结构", false, "warn",
                    $"表格列数或合并单元格异常（{columnIssues.Count} 处）",
                    string.Join("；", columnIssues.Take(8)),
                    "table_columns");
            }
            else
            {
                report.Add("表格列数/合并结构", true, "info",
                    $"表格列数与合并结构一致（{bodyTables.Count} 个表格）",
                    checkId: "table_columns");
            }
        });

        // ── 检查 15: 空段落/连续空行比例 ──
        Guarded(report, "空段落比例", "empty_paragraph_ratio", "空段落检查失败", () =>
        {
            var (totalParagraphs, emptyCount, longestEmptyRun) =
                EmptyParagraphStats(bodyParagraphs);
            var emptyRatio = totalParagraphs > 0
                ? (double)emptyCount / totalParagraphs
                : 0.0;
            var suspicious = longestEmptyRun >= 3 ||
                             (totalParagraphs >= 10 && emptyRatio > 0.3);
            var details =
                $"空段落 {emptyCount}/{totalParagraphs}（{FormatPercent(emptyRatio)}），" +
                $"最长连续空段落 {longestEmptyRun}";
            if (suspicious)
            {
                report.Add("空段落比例", false, "warn", "空段落或连续空行偏多", details,
                    "empty_paragraph_ratio");
            }
            else
            {
                report.Add("空段落比例", true, "info", details,
                    checkId: "empty_paragraph_ratio");
            }
        });

        // ── 检查 16: 中英文字体混用 ──
        Guarded(report, "中英文字体混用", "mixed_fonts", "中英文字体检查失败", () =>
        {
            var mixedFonts = MixedFontIssues(mainPart, body, sections);
            if (mixedFonts.Count > 0)
            {
                report.Add("中英文字体混用", false, "warn",
                    $"检测到 {mixedFonts.Count} 处中英文字体混用",
                    string.Join("；", mixedFonts.Take(8)),
                    "mixed_fonts");
            }
            else
            {
                report.Add("中英文字体混用", true, "info",
                    "同段/同 run 未发现中英文字体混用",
                    checkId: "mixed_fonts");
            }
        });
    }

    private static void Guarded(
        WordValidateReport report, string name, string checkId, string failurePrefix,
        Action action)
    {
        try
        {
            action();
        }
        catch (Exception ex)
        {
            report.Add(name, false, "warn", $"{failurePrefix}: {ex.Message}",
                checkId: checkId);
        }
    }

    // ── 结构辅助 ────────────────────────────────────────────────

    private static List<W.SectionProperties> CollectSections(W.Body body)
    {
        var sections = new List<W.SectionProperties>();
        foreach (var child in body.ChildElements)
        {
            if (child is W.Paragraph paragraph &&
                paragraph.ParagraphProperties?.SectionProperties is { } inParagraph)
                sections.Add(inParagraph);
            else if (child is W.SectionProperties direct)
                sections.Add(direct);
        }
        return sections;
    }

    private static List<OpenXmlElement> CollectStoryRoots(
        MainDocumentPart? mainPart, W.Body body, List<W.SectionProperties> sections)
    {
        var roots = new List<OpenXmlElement> { body };
        if (mainPart is null)
            return roots;
        for (var index = 0; index < sections.Count; index++)
        {
            foreach (var story in StoryAttributes)
            {
                if (story.IsHeader)
                {
                    var part = ResolveHeader(mainPart, sections, index, story.Type);
                    if (part?.Header is { } header)
                        roots.Add(header);
                }
                else
                {
                    var part = ResolveFooter(mainPart, sections, index, story.Type);
                    if (part?.Footer is { } footer)
                        roots.Add(footer);
                }
            }
        }
        return roots;
    }

    private static HeaderPart? ResolveHeader(
        MainDocumentPart part, List<W.SectionProperties> sections, int index, string type)
    {
        for (var current = index; current >= 0; current--)
        {
            var reference = sections[current].Elements<W.HeaderReference>()
                .FirstOrDefault(item => GetWordAttribute(item, "type") == type);
            if (reference is null)
                continue;
            var relationshipId = reference.Id?.Value;
            if (string.IsNullOrEmpty(relationshipId))
                return null;
            try
            {
                return part.GetPartById(relationshipId) as HeaderPart;
            }
            catch (Exception)
            {
                return null;
            }
        }
        return null;
    }

    private static FooterPart? ResolveFooter(
        MainDocumentPart part, List<W.SectionProperties> sections, int index, string type)
    {
        for (var current = index; current >= 0; current--)
        {
            var reference = sections[current].Elements<W.FooterReference>()
                .FirstOrDefault(item => GetWordAttribute(item, "type") == type);
            if (reference is null)
                continue;
            var relationshipId = reference.Id?.Value;
            if (string.IsNullOrEmpty(relationshipId))
                return null;
            try
            {
                return part.GetPartById(relationshipId) as FooterPart;
            }
            catch (Exception)
            {
                return null;
            }
        }
        return null;
    }

    // ── python-docx 文本语义 ────────────────────────────────────

    /// <summary>Paragraph.text：仅直接 w:r 与 w:hyperlink（字段结果 fldSimple 不计入）。</summary>
    private static string ParagraphText(W.Paragraph paragraph)
    {
        var builder = new StringBuilder();
        foreach (var child in paragraph.ChildElements)
        {
            if (child is W.Run run)
            {
                builder.Append(RunText(run));
            }
            else if (child is W.Hyperlink hyperlink)
            {
                foreach (var hyperlinkRun in hyperlink.Elements<W.Run>())
                    builder.Append(RunText(hyperlinkRun));
            }
        }
        return builder.ToString();
    }

    /// <summary>Run.text：w:t / w:tab / w:ptab / w:cr / 换行 br 的文本等价。</summary>
    private static string RunText(W.Run run)
    {
        var builder = new StringBuilder();
        foreach (var child in run.ChildElements)
        {
            switch (child.LocalName)
            {
                case "t":
                    builder.Append(child is W.Text text ? text.Text : child.InnerText);
                    break;
                case "tab":
                case "ptab":
                    builder.Append('\t');
                    break;
                case "cr":
                    builder.Append('\n');
                    break;
                case "br":
                    var breakType = GetWordAttribute(child, "type");
                    if (breakType is null or "textWrapping")
                        builder.Append('\n');
                    break;
                case "noBreakHyphen":
                    builder.Append('-');
                    break;
            }
        }
        return builder.ToString();
    }

    private static string TableCellText(W.TableCell cell) =>
        string.Join("\n", cell.Elements<W.Paragraph>().Select(ParagraphText));

    private static bool HasPageBreak(W.Run run) =>
        run.ChildElements.Any(child =>
            child.LocalName == "br" && GetWordAttribute(child, "type") == "page");

    private static bool HasDrawing(W.Run run) =>
        run.ChildElements.Any(child => child.LocalName == "drawing");

    // ── 样式与标题 ──────────────────────────────────────────────

    private sealed class StyleResolver
    {
        private readonly Dictionary<string, W.Style> _byId =
            new(StringComparer.Ordinal);
        private readonly string? _defaultName;

        public StyleResolver(MainDocumentPart? mainPart)
        {
            var styles = mainPart?.StyleDefinitionsPart?.Styles?
                .Elements<W.Style>().ToList() ?? [];
            foreach (var style in styles)
            {
                if (style.StyleId?.Value is { } styleId && !_byId.ContainsKey(styleId))
                    _byId[styleId] = style;
            }
            // python-docx default_for() 取文档顺序里最后一个 default 段落样式。
            W.Style? defaultStyle = null;
            foreach (var style in styles)
            {
                if (style.Type?.Value == W.StyleValues.Paragraph &&
                    (style.Default?.Value ?? false))
                    defaultStyle = style;
            }
            _defaultName = defaultStyle?.StyleName?.Val?.Value;
        }

        public string? ParagraphStyleName(W.Paragraph paragraph)
        {
            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (styleId is null)
                return _defaultName;
            if (_byId.TryGetValue(styleId, out var style) &&
                style.Type?.Value == W.StyleValues.Paragraph)
                return style.StyleName?.Val?.Value;
            return _defaultName;
        }
    }

    private static int? HeadingLevel(W.Paragraph paragraph, StyleResolver resolver)
    {
        var styleName = resolver.ParagraphStyleName(paragraph);
        if (string.IsNullOrEmpty(styleName))
            return null;
        var match = HeadingStylePattern.Match(styleName);
        if (!match.Success)
            return null;
        return int.TryParse(match.Groups[1].Value, NumberStyles.Integer,
            CultureInfo.InvariantCulture, out var level)
            ? level
            : null;
    }

    private static (List<string> Issues, int Numbered) HeadingNumberingIssues(
        List<(int Level, string Text)> headings)
    {
        var issues = new List<string>();
        var lastByLevel = new Dictionary<int, (int[] Number, string Text)>();
        var numbered = 0;
        foreach (var (level, text) in headings)
        {
            var number = HeadingNumber(text);
            if (number is null)
                continue;
            numbered++;
            if (lastByLevel.TryGetValue(level, out var previous) &&
                number[..^1].SequenceEqual(previous.Number[..^1]) &&
                number[^1] != previous.Number[^1] + 1)
            {
                issues.Add(
                    $"「{Truncate(previous.Text, 20)}」→「{Truncate(text, 20)}」编号 " +
                    $"{string.Join(".", previous.Number)}→{string.Join(".", number)} 不连续");
            }
            lastByLevel[level] = (number, text);
        }
        return (issues, numbered);
    }

    /// <summary>Extract a manual heading number such as 1.2, 一、 or 第一章.</summary>
    private static int[]? HeadingNumber(string text)
    {
        var value = (text ?? "").Trim();
        if (value.Length == 0)
            return null;

        var match = DecimalHeadingPattern.Match(value);
        if (match.Success)
        {
            var parts = new List<int>();
            foreach (var part in match.Groups[1].Value.Split('.'))
            {
                if (!int.TryParse(part, NumberStyles.None, CultureInfo.InvariantCulture,
                        out var parsed))
                    return null;
                parts.Add(parsed);
            }
            return parts.ToArray();
        }

        match = CnParenHeadingPattern.Match(value);
        if (match.Success)
        {
            var number = CnNumber(match.Groups[1].Value);
            return number is null ? null : [number.Value];
        }

        match = CnChapterHeadingPattern.Match(value);
        if (match.Success)
        {
            var number = CnNumber(match.Groups[1].Value);
            return number is null ? null : [number.Value];
        }

        match = CnEnumHeadingPattern.Match(value);
        if (match.Success)
        {
            var number = CnNumber(match.Groups[1].Value);
            return number is null ? null : [number.Value];
        }

        return null;
    }

    /// <summary>Parse Chinese numerals up to 999 (一、十二、二十三、一百零五).</summary>
    private static int? CnNumber(string text)
    {
        if (string.IsNullOrEmpty(text))
            return null;

        if (text.Contains('百'))
        {
            var index = text.IndexOf('百');
            var head = text[..index];
            var tail = text[(index + 1)..];
            var hundreds = head.Length > 0 ? CnDigitValue(head, 1) : 1;
            var rest = tail.Length > 0 ? CnNumber(tail) : 0;
            return hundreds * 100 + (rest ?? 0);
        }

        if (text.Contains('十'))
        {
            var index = text.IndexOf('十');
            var head = text[..index];
            var tail = text[(index + 1)..];
            var tens = head.Length > 0 ? CnDigitValue(head, 1) : 1;
            var ones = tail.Length > 0 ? CnDigitValue(tail, 0) : 0;
            return tens * 10 + ones;
        }

        var value = 0;
        foreach (var character in text)
        {
            if (!CnDigits.TryGetValue(character, out var digit))
                return null;
            value = value * 10 + digit;
        }
        return value != 0 ? value : null;
    }

    private static int CnDigitValue(string text, int fallback) =>
        text.Length == 1 && CnDigits.TryGetValue(text[0], out var digit)
            ? digit
            : fallback;

    // ── 行距 / 空段落 ───────────────────────────────────────────

    /// <summary>端口 _safe_line_spacing：返回 python-docx 的 line_spacing 四舍五入值。</summary>
    private static double? SafeLineSpacing(W.Paragraph paragraph)
    {
        var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
        if (spacing is null)
            return null;
        var lineText = GetWordAttribute(spacing, "line");
        if (string.IsNullOrEmpty(lineText))
            return null;
        if (!TryParseTwipsAsEmu(lineText, out var emu))
            return null;
        var rule = GetWordAttribute(spacing, "lineRule");
        // python-docx 在 @w:lineRule 缺失时按 MULTIPLE（auto）解释。
        var raw = rule is null or "auto" ? emu / 152400.0 : emu;
        return Math.Round(raw, 2, MidpointRounding.ToEven);
    }

    /// <summary>ST_SignedTwipsMeasure：裸数字按 twips 转 EMU，带单位时按通用度量转 EMU。</summary>
    private static bool TryParseTwipsAsEmu(string text, out double emu)
    {
        emu = 0;
        var value = text.Trim();
        if (value.Length == 0)
            return false;
        var unitMultipliers = new (string Unit, double Multiplier)[]
        {
            ("mm", 36000d), ("cm", 360000d), ("in", 914400d),
            ("pt", 12700d), ("pc", 152400d), ("pi", 152400d),
        };
        foreach (var (unit, multiplier) in unitMultipliers)
        {
            if (!value.EndsWith(unit, StringComparison.Ordinal))
                continue;
            if (!double.TryParse(value[..^unit.Length], NumberStyles.Float,
                    CultureInfo.InvariantCulture, out var quantity))
                return false;
            emu = Math.Round(quantity * multiplier, MidpointRounding.ToEven);
            return true;
        }
        if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture,
                out var twips))
            return false;
        emu = Math.Round(twips, MidpointRounding.ToEven) * 635.0;
        return true;
    }

    private static (int Total, int Empty, int Longest) EmptyParagraphStats(
        List<W.Paragraph> paragraphs)
    {
        var flags = new List<bool>();
        foreach (var paragraph in paragraphs)
        {
            if (ParagraphText(paragraph).Trim().Length > 0)
            {
                flags.Add(false);
                continue;
            }
            var hasVisualContent = paragraph.Elements<W.Run>()
                .Any(run => HasDrawing(run) || HasPageBreak(run));
            flags.Add(!hasVisualContent);
        }
        var emptyCount = flags.Count(flag => flag);
        var longest = 0;
        var current = 0;
        foreach (var flag in flags)
        {
            current = flag ? current + 1 : 0;
            longest = Math.Max(longest, current);
        }
        return (paragraphs.Count, emptyCount, longest);
    }

    // ── 表格列数/合并 ───────────────────────────────────────────

    private static List<string> TableColumnIssues(W.Table table, int tableIndex)
    {
        var issues = new List<string>();
        var grid = table.Elements<W.TableGrid>().FirstOrDefault();
        var gridColumns = grid?.Elements<W.GridColumn>().Count() ?? 0;
        var rows = table.Elements<W.TableRow>().ToList();
        int? firstTotal = null;
        var activeMerges = new HashSet<int>();

        for (var rowIndex = 1; rowIndex <= rows.Count; rowIndex++)
        {
            var cells = rows[rowIndex - 1].Elements<W.TableCell>().ToList();
            var total = 0;
            var nextMerges = new HashSet<int>();

            foreach (var cell in cells)
            {
                var span = 1;
                string? verticalMerge = null;
                var cellProperties = cell.Elements<W.TableCellProperties>().FirstOrDefault();
                if (cellProperties is not null)
                {
                    var gridSpan = cellProperties.Elements<W.GridSpan>().FirstOrDefault();
                    if (gridSpan is not null)
                    {
                        var spanText = GetWordAttribute(gridSpan, "val") ?? "1";
                        if (!int.TryParse(spanText, NumberStyles.Integer,
                                CultureInfo.InvariantCulture, out span))
                        {
                            span = 1;
                            issues.Add(
                                $"表 {tableIndex} 第 {rowIndex} 行 gridSpan 不是整数");
                        }
                        if (span < 1)
                        {
                            issues.Add(
                                $"表 {tableIndex} 第 {rowIndex} 行 gridSpan={span} 非法");
                            span = 1;
                        }
                    }
                    var merge = cellProperties.Elements<W.VerticalMerge>().FirstOrDefault();
                    if (merge is not null)
                    {
                        verticalMerge = (GetWordAttribute(merge, "val") ?? "continue")
                            .ToLowerInvariant();
                    }
                }

                if (gridColumns != 0 && span > gridColumns)
                {
                    issues.Add(
                        $"表 {tableIndex} 第 {rowIndex} 行 gridSpan={span} " +
                        $"超过网格 {gridColumns} 列");
                }
                if (verticalMerge == "continue" && !activeMerges.Contains(total))
                {
                    issues.Add(
                        $"表 {tableIndex} 第 {rowIndex} 行第 {total + 1} 列" +
                        "垂直合并缺少 restart");
                }
                if (verticalMerge is "continue" or "restart")
                {
                    for (var column = total; column < total + span; column++)
                        nextMerges.Add(column);
                }
                total += span;
            }

            if (firstTotal is null)
            {
                firstTotal = total;
            }
            else if (total != firstTotal)
            {
                issues.Add(
                    $"表 {tableIndex} 第 {rowIndex} 行 {total} 列，" +
                    $"与首行 {firstTotal} 列不一致");
            }
            if (gridColumns != 0 && total != gridColumns)
            {
                issues.Add(
                    $"表 {tableIndex} 第 {rowIndex} 行 {total} 列，" +
                    $"与表格网格 {gridColumns} 列不一致");
            }
            activeMerges = nextMerges;
        }
        return issues;
    }

    // ── 字体与字段 ──────────────────────────────────────────────

    private static List<string> CollectExplicitFonts(IEnumerable<OpenXmlElement> roots)
    {
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (var root in roots)
        {
            foreach (var runFonts in root.Descendants<W.RunFonts>())
            {
                foreach (var attribute in new[] { "ascii", "hAnsi", "eastAsia", "cs" })
                {
                    var value = GetWordAttribute(runFonts, attribute);
                    if (!string.IsNullOrEmpty(value) &&
                        !value.StartsWith("+", StringComparison.Ordinal))
                        names.Add(value.Trim());
                }
            }
        }
        return names
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static List<string> FieldInstructions(IEnumerable<OpenXmlElement> roots)
    {
        var instructions = new List<string>();
        foreach (var root in roots)
        {
            foreach (var simpleField in root.Descendants<W.SimpleField>())
                instructions.Add(GetWordAttribute(simpleField, "instr") ?? "");
            foreach (var fieldCode in root.Descendants<W.FieldCode>())
                instructions.Add(fieldCode.Text ?? "");
        }
        return instructions;
    }

    private static bool TocMarkedForUpdate(
        MainDocumentPart? mainPart, IEnumerable<OpenXmlElement> roots)
    {
        foreach (var root in roots)
        {
            foreach (var simpleField in root.Descendants<W.SimpleField>())
            {
                var instruction = GetWordAttribute(simpleField, "instr") ?? "";
                if (TocFieldPattern.IsMatch(instruction) &&
                    IsTruthy(GetWordAttribute(simpleField, "dirty")))
                    return true;
            }
            foreach (var fieldChar in root.Descendants<W.FieldChar>())
            {
                if (IsTruthy(GetWordAttribute(fieldChar, "dirty")))
                    return true;
            }
        }
        try
        {
            var settings = mainPart?.DocumentSettingsPart?.Settings;
            if (settings is not null)
            {
                foreach (var updateFields in settings.Descendants<W.UpdateFieldsOnOpen>())
                {
                    if (IsTruthy(GetWordAttribute(updateFields, "val")))
                        return true;
                }
            }
        }
        catch (Exception)
        {
            // 与 Python 的 try/except pass 一致。
        }
        return false;
    }

    private static bool IsTruthy(string? value) =>
        (value ?? "").ToLowerInvariant() is "true" or "1" or "on";

    // ── 中英文字体混用 ──────────────────────────────────────────

    private static List<string> MixedFontIssues(
        MainDocumentPart? mainPart, W.Body body, List<W.SectionProperties> sections)
    {
        var issues = new List<string>();
        foreach (var paragraph in ContentParagraphs(mainPart, body, sections))
        {
            var eastAsian = new SortedSet<string>(StringComparer.Ordinal);
            var latin = new SortedSet<string>(StringComparer.Ordinal);

            foreach (var run in paragraph.Elements<W.Run>())
            {
                var fonts = RunExplicitFonts(run);
                if (fonts.Count == 0)
                    continue;

                if (fonts.TryGetValue("eastAsia", out var eastAsiaValue))
                    eastAsian.Add(eastAsiaValue);
                foreach (var key in new[] { "ascii", "hAnsi" })
                {
                    if (fonts.TryGetValue(key, out var latinValue))
                        latin.Add(latinValue);
                }

                var text = RunText(run);
                var hasCjk = CjkPattern.IsMatch(text);
                var hasLatin = LatinPattern.IsMatch(text);
                if (hasCjk && !hasLatin && !fonts.ContainsKey("eastAsia") &&
                    fonts.TryGetValue("ascii", out var asciiOnly))
                    issues.Add($"中文 run 只设置西文字体 {asciiOnly}");
                if (hasLatin && !hasCjk &&
                    !(fonts.ContainsKey("ascii") || fonts.ContainsKey("hAnsi")) &&
                    fonts.TryGetValue("eastAsia", out var eastAsiaOnly))
                    issues.Add($"英文 run 只设置中文字体 {eastAsiaOnly}");
            }

            if (eastAsian.Count > 1)
                issues.Add("同段中文字体混用: " + string.Join("、", eastAsian));
            if (latin.Count > 1)
                issues.Add("同段英文字体混用: " + string.Join("、", latin));
        }

        var seen = new HashSet<string>(StringComparer.Ordinal);
        var unique = new List<string>();
        foreach (var issue in issues)
        {
            if (!seen.Add(issue))
                continue;
            unique.Add(issue);
            if (unique.Count >= 20)
                break;
        }
        return unique;
    }

    private static IEnumerable<W.Paragraph> ContentParagraphs(
        MainDocumentPart? mainPart, W.Body body, List<W.SectionProperties> sections)
    {
        foreach (var paragraph in body.Elements<W.Paragraph>())
            yield return paragraph;
        foreach (var table in body.Elements<W.Table>())
        {
            foreach (var row in table.Elements<W.TableRow>())
            {
                foreach (var cell in row.Elements<W.TableCell>())
                {
                    foreach (var paragraph in cell.Elements<W.Paragraph>())
                        yield return paragraph;
                }
            }
        }
        if (mainPart is null)
            yield break;
        for (var index = 0; index < sections.Count; index++)
        {
            foreach (var story in StoryAttributes)
            {
                if (story.IsHeader)
                {
                    var part = ResolveHeader(mainPart, sections, index, story.Type);
                    if (part?.Header is not { } header)
                        continue;
                    foreach (var paragraph in header.Elements<W.Paragraph>())
                        yield return paragraph;
                }
                else
                {
                    var part = ResolveFooter(mainPart, sections, index, story.Type);
                    if (part?.Footer is not { } footer)
                        continue;
                    foreach (var paragraph in footer.Elements<W.Paragraph>())
                        yield return paragraph;
                }
            }
        }
    }

    private static Dictionary<string, string> RunExplicitFonts(W.Run run)
    {
        var fonts = new Dictionary<string, string>(StringComparer.Ordinal);
        var runProperties = run.Elements<W.RunProperties>().FirstOrDefault();
        var runFonts = runProperties?.Elements<W.RunFonts>().FirstOrDefault();
        if (runFonts is null)
            return fonts;
        foreach (var attribute in new[] { "ascii", "hAnsi", "eastAsia" })
        {
            var value = GetWordAttribute(runFonts, attribute);
            if (!string.IsNullOrEmpty(value))
                fonts[attribute] = value;
        }
        return fonts;
    }

    // ── 通用工具 ────────────────────────────────────────────────

    private static string? GetWordAttribute(OpenXmlElement? element, string localName)
    {
        if (element is null)
            return null;
        foreach (var attribute in element.GetAttributes())
        {
            if (attribute.LocalName == localName &&
                attribute.NamespaceUri == WordNamespace)
                return attribute.Value;
        }
        return null;
    }

    private static string Truncate(string text, int length) =>
        text.Length <= length ? text : text[..length];

    private static string FormatNumberList(IEnumerable<double> values) =>
        "[" + string.Join(", ", values.Select(FormatPythonFloat)) + "]";

    private static string FormatPythonFloat(double value)
    {
        var text = value.ToString("R", CultureInfo.InvariantCulture);
        if (text.IndexOf('.') < 0 && text.IndexOf('e') < 0 && text.IndexOf('E') < 0)
            text += ".0";
        return text;
    }

    private static string FormatPercent(double ratio)
    {
        var rounded = Math.Round(ratio * 100.0, 0, MidpointRounding.ToEven);
        return rounded.ToString("0", CultureInfo.InvariantCulture) + "%";
    }
}
