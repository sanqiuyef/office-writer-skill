using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using XPresentation = DocumentFormat.OpenXml.Presentation;

namespace OfficeNative;

/// <summary>PowerPoint 文件级命令入口（主会话在 CommandRouter 中注册这些 Run* 方法）。</summary>
internal static class PptFileCommands
{
    // ── create-powerpoint ────────────────────────────────────────

    public static Dictionary<string, object?> RunCreate(JsonElement p)
    {
        var outputPath = NativeCommon.RequireString(p, "output_path");
        EnsurePptExtension(outputPath, "输出文件必须使用 .pptx 扩展名");
        var slides = PptModels.ParseSlides(p);
        var template = NativeCommon.OptString(p, "template");
        var resolved = ResolveOutput(outputPath, p);
        var work = PptFileOoxml.WorkPath(resolved);
        try
        {
            if (!string.IsNullOrWhiteSpace(template))
            {
                var templatePath = Path.GetFullPath(template!);
                NativeCommon.EnsureFile(templatePath, ".pptx", ".potx");
                File.Copy(templatePath, work, overwrite: true);
                if (Path.GetExtension(templatePath).Equals(".potx",
                        StringComparison.OrdinalIgnoreCase))
                    PptFileOoxml.NormalizePotxContentType(work);
                using var document = PptFileOoxml.Open(work, true);
                var presentationPart = PptFileOoxml.MainPart(document);
                if (FirstSlideSize(slides) is { } size)
                    PptFileOoxml.SetSlideSize(presentationPart, size.Width, size.Height);
                AddSlides(presentationPart, slides);
                document.Save();
            }
            else
            {
                var size = FirstSlideSize(slides) ?? (Width: 10.0, Height: 7.5);
                PptFileOoxml.CreateScratch(work, size.Width, size.Height);
                using var document = PptFileOoxml.Open(work, true);
                AddSlides(PptFileOoxml.MainPart(document), slides);
                document.Save();
            }

            var backupPath = PptFileOoxml.ReplaceWithBackup(work, resolved);
            var fileSize = new FileInfo(resolved).Length;
            return new Dictionary<string, object?>
            {
                ["success"] = true,
                ["file_path"] = Path.GetFullPath(resolved),
                ["file_size"] = fileSize,
                ["slides_count"] = slides.Count,
                ["message"] = $"PowerPoint 演示文稿创建成功: {Path.GetFullPath(resolved)}",
                ["backup_path"] = backupPath,
            };
        }
        finally
        {
            PptFileOoxml.EnsureCleanup(work);
        }
    }

    private static (double Width, double Height)? FirstSlideSize(List<PptSlideDef> slides)
    {
        foreach (var slide in slides)
        {
            if (slide.SlideWidth is not null && slide.SlideHeight is not null)
                return (slide.SlideWidth.Value, slide.SlideHeight.Value);
        }
        return null;
    }

    private static void AddSlides(PresentationPart presentationPart,
        IEnumerable<PptSlideDef> slides)
    {
        foreach (var slide in slides)
        {
            var writer = PptFileContent.AddSlide(presentationPart,
                PptModels.GetLayoutIndex(slide.Layout), slide.BgColor);
            writer.SetPlaceholderText(slide.Title ?? "", "title", "ctrTitle");
            if (!string.IsNullOrEmpty(slide.Subtitle)
                && slide.Layout is "title" or "section")
                writer.SetPlaceholderText(slide.Subtitle!, "subTitle", "body");
            else
                writer.SetPlaceholderText(slide.Subtitle ?? "", "subTitle");
            writer.WriteContent(slide.Body, PptDefaults.Body, 1);
            if (slide.Left is not null || slide.Right is not null)
            {
                writer.WriteContent(slide.Left, PptDefaults.LeftColumn, 1);
                writer.WriteContent(slide.Right, PptDefaults.RightColumn, 2);
            }
            if (!string.IsNullOrEmpty(slide.Notes))
                PptFileOoxml.AddNotes(writer.Part, slide.Notes!);
            writer.Save();
        }
    }

    // ── edit-powerpoint ──────────────────────────────────────────

    public static Dictionary<string, object?> RunEdit(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        NativeCommon.EnsureFile(filePath, ".pptx");
        var slides = PptModels.ParseSlides(p, "append_slides");
        var fullPath = Path.GetFullPath(filePath);
        // 文件已在 PowerPoint/WPS 中打开时优先走活动文档 COM，避免 OOXML 写回被锁。
        if (OpenEditPptCommands.TryEditOpenPowerpoint(fullPath, slides) is { } comResult)
            return comResult;
        var work = PptFileOoxml.WorkPath(fullPath);
        try
        {
            File.Copy(fullPath, work, overwrite: true);
            using (var document = PptFileOoxml.Open(work, true))
            {
                AddSlides(PptFileOoxml.MainPart(document), slides);
                document.Save();
            }
            var backupPath = PptFileOoxml.ReplaceWithBackup(work, fullPath);
            return new Dictionary<string, object?>
            {
                ["success"] = true,
                ["file_path"] = fullPath,
                ["operation_count"] = slides.Count,
                ["message"] = $"已追加 {slides.Count} 张幻灯片到 PPT",
                ["backup_path"] = backupPath,
            };
        }
        finally
        {
            PptFileOoxml.EnsureCleanup(work);
        }
    }

    // ── read-powerpoint ──────────────────────────────────────────

    public static Dictionary<string, object?> RunRead(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        NativeCommon.EnsureFile(filePath, ".pptx");
        var maxChars = NativeCommon.OptInt(p, "max_chars", 50000);
        if (maxChars < 1)
            throw new NativeUsageException("max_chars 必须大于 0");

        using var document = PptFileOoxml.Open(filePath, false);
        var presentationPart = PptFileOoxml.MainPart(document);
        var slides = PptFileOoxml.SlideParts(presentationPart);
        var parts = new List<string>();
        var contentLength = 0;
        var truncated = false;
        var slideNumber = 0;
        var includedSlides = 0;

        foreach (var slidePart in slides)
        {
            slideNumber++;
            var slideDocument = PptFileOoxml.SlideDocument(slidePart);
            var lines = new List<string> { $"## Slide {slideNumber}" };
            var layoutName = slidePart.SlideLayoutPart is { } layoutPart
                ? PptFileOoxml.LayoutName(layoutPart)
                : "";
            lines.Add($"**Layout:** {layoutName}");

            foreach (var shape in PptFileOoxml.ShapeElements(slideDocument))
            {
                var textBody = shape.Element(PptFileOoxml.P + "txBody");
                if (textBody is not null)
                {
                    foreach (var paragraph in textBody.Elements(PptFileOoxml.A + "p"))
                    {
                        var text = PptFileOoxml.ParagraphText(paragraph).Trim();
                        if (text.Length == 0)
                            continue;
                        var level = PptFileOoxml.ParagraphLevel(paragraph);
                        var prefix = new string(' ', 2 * level) + (level > 0 ? "- " : "");
                        lines.Add(prefix + text);
                    }
                }

                if (shape.Name == PptFileOoxml.P + "graphicFrame")
                {
                    var table = shape.Element(PptFileOoxml.A + "graphic")
                        ?.Element(PptFileOoxml.A + "graphicData")
                        ?.Element(PptFileOoxml.A + "tbl");
                    if (table is not null)
                    {
                        lines.Add("");
                        lines.Add("[Table]");
                        var rowIndex = 0;
                        foreach (var row in table.Elements(PptFileOoxml.A + "tr"))
                        {
                            var cells = row.Elements(PptFileOoxml.A + "tc")
                                .Select(cell => PptFileOoxml
                                    .TxBodyText(cell.Element(PptFileOoxml.A + "txBody"))
                                    .Trim())
                                .ToList();
                            lines.Add("| " + string.Join(" | ", cells) + " |");
                            if (rowIndex == 0)
                                lines.Add("| " + string.Join(" | ",
                                    cells.Select(_ => "---")) + " |");
                            rowIndex++;
                        }
                    }
                }
            }

            var notes = PptFileOoxml.ReadNotesText(slidePart);
            if (!string.IsNullOrEmpty(notes))
                lines.Add($"\n*Notes:* {notes}");

            var slideText = string.Join("\n", lines);
            var separator = parts.Count > 0 ? "\n\n" : "";
            if (contentLength + separator.Length + slideText.Length > maxChars)
            {
                truncated = true;
                break;
            }
            parts.Add(slideText);
            contentLength += separator.Length + slideText.Length;
            includedSlides++;
        }

        var content = truncated
            ? WithTruncationMarker(parts, maxChars)
            : string.Join("\n\n", parts);
        var actualSlides = truncated ? slideNumber - 1 : includedSlides;
        var message = $"成功读取演示文稿，共 {slides.Count} 张幻灯片";
        if (truncated)
            message += $"（已截断至 {maxChars} 字符，显示前 {actualSlides} 张）";

        var result = new Dictionary<string, object?>
        {
            ["success"] = true,
            ["file_path"] = filePath,
            ["content"] = content,
            ["slides_count"] = actualSlides,
            ["truncated"] = truncated,
            ["message"] = message,
        };
        return new Dictionary<string, object?>
        {
            ["markdown"] = ReadResultToMarkdown(result),
            ["result"] = result,
        };
    }

    private static string WithTruncationMarker(List<string> parts, int maxChars)
    {
        const string marker = "_(内容已截断)_";
        var content = string.Join("\n\n", parts);
        if (content.Length == 0)
            return marker[..Math.Min(maxChars, marker.Length)];
        var candidate = $"{content}\n\n{marker}";
        return candidate.Length <= maxChars ? candidate : content;
    }

    public static string ReadResultToMarkdown(Dictionary<string, object?> result)
    {
        if (result.GetValueOrDefault("success") is false)
            return $"读取失败: {result.GetValueOrDefault("message") ?? "未知错误"}";
        var parts = new List<string> { result.GetValueOrDefault("content")?.ToString() ?? "" };
        var stats = result
            .Where(pair => pair.Key is not ("success" or "file_path" or "content"
                or "message" or "truncated") && pair.Value is not null)
            .Select(pair => $"{pair.Key}: {ValueText(pair.Value)}")
            .ToList();
        if (stats.Count > 0)
            parts.Add("> " + string.Join("；", stats));
        if (result.GetValueOrDefault("truncated") is true)
            parts.Add("> 内容过长已截断");
        return string.Join("\n\n", parts.Where(part => part.Length > 0));
    }

    private static string ValueText(object? value) => value switch
    {
        null => "",
        bool boolean => boolean ? "True" : "False",
        _ => Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? "",
    };

    // ── ppt-list-color-schemes ──────────────────────────────────

    public static Dictionary<string, object?> RunListColorSchemes(JsonElement p) =>
        new() { ["schemes"] = PptModels.ListSchemes() };

    // ── ppt-apply-color-scheme ──────────────────────────────────

    public static Dictionary<string, object?> RunApplyColorScheme(JsonElement p) =>
        PptFileEnhance.ApplyColorScheme(p);

    // ── ppt-optimize-slide-layout ───────────────────────────────

    public static Dictionary<string, object?> RunOptimizeSlideLayout(JsonElement p) =>
        PptFileEnhance.OptimizeSlideLayout(p);

    // ── ppt-clean-unused-placeholders ───────────────────────────

    public static Dictionary<string, object?> RunCleanUnusedPlaceholders(JsonElement p) =>
        PptFileEnhance.CleanUnusedPlaceholders(p);

    // ── ppt-analyze-template ────────────────────────────────────

    public static Dictionary<string, object?> RunAnalyzeTemplate(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        NativeCommon.EnsureFile(filePath, ".pptx", ".potx");
        var fullPath = Path.GetFullPath(filePath);
        var extension = Path.GetExtension(fullPath).ToLowerInvariant();
        string? temporary = null;
        try
        {
            if (extension == ".potx")
            {
                temporary = PptFileOoxml.WorkPath(fullPath);
                File.Copy(fullPath, temporary, overwrite: true);
                PptFileOoxml.NormalizePotxContentType(temporary);
            }
            using var document = PptFileOoxml.Open(temporary ?? fullPath, false);
            var presentationPart = PptFileOoxml.MainPart(document);
            var presentation = presentationPart.Presentation;
            var slideWidth = PptFileOoxml.ToInches(
                presentation.SlideSize?.Cx?.Value ?? 9144000);
            var slideHeight = PptFileOoxml.ToInches(
                presentation.SlideSize?.Cy?.Value ?? 6858000);
            var layoutNames = PptFileOoxml.LayoutParts(presentationPart)
                .Select(PptFileOoxml.LayoutName).ToList();

            var fonts = new HashSet<string>(StringComparer.Ordinal);
            var colors = new HashSet<string>(StringComparer.Ordinal);
            var slideDocuments = PptFileOoxml.SlideParts(presentationPart)
                .Select(PptFileOoxml.SlideDocument).ToList();
            foreach (var slideDocument in slideDocuments)
            {
                foreach (var textBody in PptFileOoxml.TextBodiesInDocument(slideDocument))
                {
                    foreach (var run in textBody.Descendants(PptFileOoxml.A + "r"))
                        CollectRunProperties(run.Element(PptFileOoxml.A + "rPr"), fonts, colors);
                }
            }
            // 模板可能只把字体/颜色写在段落级 defRPr 或域（a:fld）里；
            // 仅在 run 级完全无收获时回退，保持与 Python run.font 行为一致。
            if (fonts.Count == 0 && colors.Count == 0)
            {
                foreach (var slideDocument in slideDocuments)
                {
                    foreach (var textBody in PptFileOoxml.TextBodiesInDocument(slideDocument))
                    {
                        foreach (var properties in textBody.Descendants()
                                     .Where(element => element.Name == PptFileOoxml.A + "rPr"
                                         || element.Name == PptFileOoxml.A + "defRPr"
                                         || element.Name == PptFileOoxml.A + "endParaRPr"))
                            CollectRunProperties(properties, fonts, colors);
                    }
                }
            }

            var sortedFonts = fonts.OrderBy(font => font, StringComparer.Ordinal).ToList();
            var sortedColors = colors.OrderBy(color => color, StringComparer.Ordinal).ToList();
            return new Dictionary<string, object?>
            {
                ["success"] = true,
                ["file_path"] = filePath,
                ["slide_count"] = PptFileOoxml.SlideParts(presentationPart).Count,
                ["slide_width"] = Math.Round(slideWidth, 2),
                ["slide_height"] = Math.Round(slideHeight, 2),
                ["layout_names"] = layoutNames,
                ["fonts_used"] = sortedFonts,
                ["color_hexes_found"] = sortedColors,
                ["message"] = $"模板分析完成: {layoutNames.Count} 种布局, "
                              + $"{sortedFonts.Count} 种字体, {sortedColors.Count} 种颜色",
            };
        }
        finally
        {
            if (temporary is not null)
                PptFileOoxml.EnsureCleanup(temporary);
        }
    }

    // ── ppt-set-hyperlinks ──────────────────────────────────────

    public static Dictionary<string, object?> RunSetHyperlinks(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        NativeCommon.EnsureFile(filePath, ".pptx");
        var links = PptModels.ParseHyperlinks(p);
        var fullPath = Path.GetFullPath(filePath);
        var work = PptFileOoxml.WorkPath(fullPath);
        try
        {
            File.Copy(fullPath, work, overwrite: true);
            var applied = 0;
            using (var document = PptFileOoxml.Open(work, true))
            {
                var presentationPart = PptFileOoxml.MainPart(document);
                var slides = PptFileOoxml.SlideParts(presentationPart);
                foreach (var link in links)
                {
                    if (link.Slide > slides.Count)
                        throw new NativeUsageException(
                            $"link.slide={link.Slide} 超出范围：文件共有 {slides.Count} 张幻灯片");
                    var slidePart = slides[link.Slide - 1];
                    var slideDocument = PptFileOoxml.SlideDocument(slidePart);
                    var shapes = PptFileOoxml.ShapeElements(slideDocument).ToList();

                    XElement? shape = null;
                    if (link.ShapeIndex is not null)
                    {
                        if (link.ShapeIndex.Value > shapes.Count)
                            throw new NativeUsageException(
                                $"第 {link.Slide} 张幻灯片只有 {shapes.Count} 个形状，"
                                + $"shape_index={link.ShapeIndex} 越界");
                        shape = shapes[link.ShapeIndex.Value - 1];
                    }
                    else
                    {
                        var comparison = link.MatchCase
                            ? StringComparison.Ordinal
                            : StringComparison.OrdinalIgnoreCase;
                        shape = shapes.FirstOrDefault(candidate =>
                        {
                            var textBody = candidate.Element(PptFileOoxml.P + "txBody");
                            return textBody is not null
                                   && PptFileOoxml.TxBodyText(textBody)
                                       .Contains(link.ShapeText!, comparison);
                        });
                        if (shape is null)
                            throw new NativeUsageException(
                                $"第 {link.Slide} 张幻灯片中找不到包含文字 "
                                + $"{link.ShapeText!} 的形状");
                    }

                    if (!string.IsNullOrEmpty(link.TextMatch))
                    {
                        var comparison = link.MatchCase
                            ? StringComparison.Ordinal
                            : StringComparison.OrdinalIgnoreCase;
                        var textBody = shape.Element(PptFileOoxml.P + "txBody");
                        var matched = false;
                        if (textBody is not null)
                        {
                            var relationshipId = PptFileContent
                                .AddHyperlinkRelationship(slidePart, link.Url);
                            foreach (var run in textBody.Descendants(PptFileOoxml.A + "r"))
                            {
                                var text = run.Element(PptFileOoxml.A + "t")?.Value ?? "";
                                if (!text.Contains(link.TextMatch!, comparison))
                                    continue;
                                var runProperties = run.Element(PptFileOoxml.A + "rPr");
                                if (runProperties is null)
                                {
                                    runProperties = new XElement(PptFileOoxml.A + "rPr");
                                    run.AddFirst(runProperties);
                                }
                                runProperties.Element(PptFileOoxml.A + "hlinkClick")?.Remove();
                                runProperties.Add(new XElement(PptFileOoxml.A + "hlinkClick",
                                    new XAttribute(PptFileOoxml.R + "id", relationshipId)));
                                matched = true;
                            }
                        }
                        if (!matched)
                            throw new NativeUsageException(
                                $"第 {link.Slide} 张幻灯片的目标形状中找不到匹配文字 "
                                + $"{link.TextMatch!}");
                    }
                    else
                    {
                        var cnvPr = PptFileOoxml.CnvPr(shape);
                        if (cnvPr is null)
                            throw new NativeUsageException(
                                $"第 {link.Slide} 张幻灯片的形状不支持整体点击超链接");
                        var relationshipId = PptFileContent
                            .AddHyperlinkRelationship(slidePart, link.Url);
                        cnvPr.Element(PptFileOoxml.A + "hlinkClick")?.Remove();
                        cnvPr.Add(new XElement(PptFileOoxml.A + "hlinkClick",
                            new XAttribute(PptFileOoxml.R + "id", relationshipId)));
                    }
                    PptFileOoxml.SaveSlide(slidePart, slideDocument);
                    applied++;
                }
                document.Save();
            }
            var backupPath = PptFileOoxml.ReplaceWithBackup(work, fullPath);
            return new Dictionary<string, object?>
            {
                ["success"] = true,
                ["file_path"] = filePath,
                ["applied"] = applied,
                ["backup_path"] = backupPath,
            };
        }
        finally
        {
            PptFileOoxml.EnsureCleanup(work);
        }
    }

    // ── merge-office-template（pptx 分支）───────────────────────

    private static readonly Regex PlaceholderRegex = new(
        @"\{\{\s*([^{}]+?)\s*\}\}", RegexOptions.Compiled);

    public static Dictionary<string, object?> RunMergeTemplate(JsonElement p)
    {
        var templatePath = NativeCommon.RequireString(p, "template_path");
        var outputPath = NativeCommon.RequireString(p, "output_path");
        var templateFull = Path.GetFullPath(templatePath);
        NativeCommon.EnsureFile(templateFull, ".pptx");
        EnsurePptExtension(outputPath, "输出文件必须使用 .pptx 扩展名");
        if (p.ValueKind != JsonValueKind.Object || !p.TryGetProperty("data", out var data)
            || data.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException(
                "缺少参数 data（通过 --params 或 --params-file 传入 JSON 对象）");

        var resolved = ResolveOutput(outputPath, p);
        var values = FlattenData(data);
        var used = new HashSet<string>(StringComparer.Ordinal);
        var replaced = 0;
        var work = PptFileOoxml.WorkPath(resolved);
        try
        {
            File.Copy(templateFull, work, overwrite: true);
            using (var document = PptFileOoxml.Open(work, true))
            {
                var presentationPart = PptFileOoxml.MainPart(document);
                foreach (var slidePart in PptFileOoxml.SlideParts(presentationPart))
                {
                    var slideDocument = PptFileOoxml.SlideDocument(slidePart);
                    replaced += ReplaceInTextBodies(
                        PptFileOoxml.TextBodiesInDocument(slideDocument), values, used);
                    PptFileOoxml.SaveSlide(slidePart, slideDocument);

                    var notesPart = slidePart.NotesSlidePart;
                    if (notesPart?.NotesSlide is not null)
                    {
                        var notesDocument = XDocument.Parse(notesPart.NotesSlide.OuterXml);
                        replaced += ReplaceInTextBodies(
                            PptFileOoxml.TextBodiesInDocument(notesDocument), values, used);
                        notesPart.NotesSlide = new XPresentation.NotesSlide(
                            notesDocument.Root!.ToString(
                                System.Xml.Linq.SaveOptions.DisableFormatting));
                    }
                }
                document.Save();
            }
            PptFileOoxml.ReplaceWithBackup(work, resolved);

            var unresolved = new HashSet<string>(StringComparer.Ordinal);
            using (var document = PptFileOoxml.Open(resolved, false))
            {
                var presentationPart = PptFileOoxml.MainPart(document);
                foreach (var slidePart in PptFileOoxml.SlideParts(presentationPart))
                {
                    foreach (var textBody in PptFileOoxml.AllTextBodies(slidePart))
                    {
                        foreach (var paragraph in textBody.Elements(PptFileOoxml.A + "p"))
                        {
                            foreach (Match match in PlaceholderRegex.Matches(
                                         PptFileOoxml.ParagraphText(paragraph)))
                                unresolved.Add(match.Groups[1].Value.Trim());
                        }
                    }
                }
            }

            var resolvedPath = Path.GetFullPath(resolved);
            var usedKeys = used.OrderBy(key => key, StringComparer.Ordinal).ToList();
            var unresolvedKeys = unresolved.OrderBy(key => key, StringComparer.Ordinal).ToList();
            return new Dictionary<string, object?>
            {
                ["success"] = true,
                ["template_path"] = templateFull,
                ["file_path"] = resolvedPath,
                ["file_size"] = new FileInfo(resolvedPath).Length,
                ["replaced_count"] = replaced,
                ["used_keys"] = usedKeys,
                ["unresolved_placeholders"] = unresolvedKeys,
                ["message"] = $"模板合并完成，替换 {replaced} 处，"
                              + $"占位符未解析 {unresolvedKeys.Count} 个",
            };
        }
        finally
        {
            PptFileOoxml.EnsureCleanup(work);
        }
    }

    private static int ReplaceInTextBodies(IEnumerable<XElement> textBodies,
        Dictionary<string, string> values, HashSet<string> used)
    {
        var count = 0;
        foreach (var textBody in textBodies)
        {
            foreach (var paragraph in textBody.Elements(PptFileOoxml.A + "p"))
            {
                var runs = paragraph.Elements(PptFileOoxml.A + "r").ToList();
                var original = string.Concat(runs
                    .Select(run => run.Element(PptFileOoxml.A + "t")?.Value ?? ""));
                if (!original.Contains("{{", StringComparison.Ordinal))
                    continue;
                var replacements = 0;
                var changed = PlaceholderRegex.Replace(original, match =>
                {
                    var key = match.Groups[1].Value.Trim();
                    if (!values.TryGetValue(key, out var value))
                        return match.Value;
                    used.Add(key);
                    replacements++;
                    return value;
                });
                if (replacements == 0)
                    continue;
                count += replacements;
                if (changed.Contains('\n'))
                {
                    // 重建段落只能处理纯 run；若段落还有字段/换行，重建会重复或丢失节点。
                    if (paragraph.Elements().Any(element =>
                            element.Name == PptFileOoxml.A + "fld"
                            || element.Name == PptFileOoxml.A + "br"))
                        throw new NativeRuntimeException(
                            "模板合并失败：替换值包含换行，且目标段落含字段或换行（a:fld/a:br）；"
                            + "请改用不含换行的值，或把占位符移到独立段落");
                    var firstRunProperties = runs.Count > 0
                        ? runs[0].Element(PptFileOoxml.A + "rPr")
                        : null;
                    foreach (var run in runs)
                        run.Remove();
                    PptFileContent.AppendTextContent(paragraph, changed, firstRunProperties);
                }
                else if (runs.Count > 0)
                {
                    var firstText = runs[0].Element(PptFileOoxml.A + "t");
                    if (firstText is null)
                    {
                        firstText = PptFileContent.TextElement(changed);
                        runs[0].Add(firstText);
                    }
                    else
                    {
                        firstText.Value = changed;
                    }
                    for (var index = 1; index < runs.Count; index++)
                    {
                        var text = runs[index].Element(PptFileOoxml.A + "t");
                        if (text is not null)
                            text.Value = "";
                    }
                }
                else
                {
                    paragraph.Add(new XElement(PptFileOoxml.A + "r",
                        PptFileContent.TextElement(changed)));
                }
            }
        }
        return count;
    }

    private static Dictionary<string, string> FlattenData(JsonElement data)
    {
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var property in data.EnumerateObject())
            values[property.Name] = Stringify(property.Value);

        void Visit(JsonElement value, string prefix)
        {
            if (value.ValueKind == JsonValueKind.Object)
            {
                foreach (var property in value.EnumerateObject())
                {
                    var path = $"{prefix}.{property.Name}";
                    values.TryAdd(path, Stringify(property.Value));
                    Visit(property.Value, path);
                }
            }
            else if (value.ValueKind == JsonValueKind.Array)
            {
                var index = 0;
                foreach (var item in value.EnumerateArray())
                {
                    var path = $"{prefix}[{index}]";
                    values.TryAdd(path, Stringify(item));
                    Visit(item, path);
                    index++;
                }
            }
        }

        foreach (var property in data.EnumerateObject())
            Visit(property.Value, property.Name);
        return values;
    }

    private static readonly JsonSerializerOptions RelaxedJson = new()
    {
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    private static string Stringify(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.Null or JsonValueKind.Undefined => "",
        JsonValueKind.String => value.GetString() ?? "",
        JsonValueKind.True => "True",
        JsonValueKind.False => "False",
        JsonValueKind.Object or JsonValueKind.Array => SerializePythonJson(value),
        _ => value.GetRawText(),
    };

    /// <summary>按 Python json.dumps(ensure_ascii=False) 的默认分隔符输出嵌套值。</summary>
    private static string SerializePythonJson(JsonElement value)
    {
        switch (value.ValueKind)
        {
            case JsonValueKind.Object:
                return "{" + string.Join(", ", value.EnumerateObject().Select(property =>
                    JsonSerializer.Serialize(property.Name, RelaxedJson) + ": "
                    + SerializePythonJson(property.Value))) + "}";
            case JsonValueKind.Array:
                return "[" + string.Join(", ", value.EnumerateArray()
                    .Select(SerializePythonJson)) + "]";
            case JsonValueKind.String:
                return JsonSerializer.Serialize(value.GetString()!, RelaxedJson);
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

    // ── 输出路径策略 ────────────────────────────────────────────

    private static void CollectRunProperties(XElement? properties,
        HashSet<string> fonts, HashSet<string> colors)
    {
        if (properties is null)
            return;
        foreach (var font in new[]
                 {
                     properties.Element(PptFileOoxml.A + "latin"),
                     properties.Element(PptFileOoxml.A + "ea"),
                 })
        {
            var typeface = (string?)font?.Attribute("typeface");
            if (!string.IsNullOrWhiteSpace(typeface) && !typeface!.StartsWith('+'))
                fonts.Add(typeface);
        }
        var color = (string?)properties
            .Element(PptFileOoxml.A + "solidFill")
            ?.Element(PptFileOoxml.A + "srgbClr")?.Attribute("val");
        if (!string.IsNullOrWhiteSpace(color))
            colors.Add(color!.ToUpperInvariant());
    }

    public static void EnsurePptExtension(string path, string message)
    {
        if (!Path.GetExtension(path).Equals(".pptx", StringComparison.OrdinalIgnoreCase))
            throw new NativeRuntimeException(message);
    }

    public static string ResolveOutput(string outputPath, JsonElement p)
    {
        var full = Path.GetFullPath(outputPath);
        var mode = NativeCommon.OptString(p, "if_exists") ?? "error";
        var confirm = PptModels.OptBool(p, "confirm_overwrite", false);
        if (!File.Exists(full) || new FileInfo(full).Length == 0)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(full)!);
            return full;
        }
        switch (mode)
        {
            case "overwrite":
                NativeCommon.EnsureOutput(full, overwrite: true, confirm: confirm);
                return full;
            case "rename":
                var directory = Path.GetDirectoryName(full)!;
                var stem = Path.GetFileNameWithoutExtension(full);
                var suffix = Path.GetExtension(full);
                for (var index = 1;; index++)
                {
                    var candidate = Path.Combine(directory, $"{stem} ({index}){suffix}");
                    if (!File.Exists(candidate))
                        return candidate;
                }
            case "error":
                throw new NativeUsageException($"文件已存在: {full}");
            default:
                throw new NativeUsageException($"不支持的 if_exists 值: {mode}");
        }
    }
}
