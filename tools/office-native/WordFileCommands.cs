using System.Globalization;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeNative;

internal static class WordFileCommands
{
    // ── create-word-document ────────────────────────────────────

    public static Dictionary<string, object?> RunCreateWordDocument(JsonElement p)
    {
        var outputPath = NativeCommon.RequireString(p, "output_path");
        if (!WordJson.TryGet(p, "content", out var contentElement)
            || contentElement.ValueKind == JsonValueKind.Null)
            throw new NativeUsageException("缺少参数 content");
        var content = WordContent.Parse(contentElement);
        var template = NativeCommon.OptString(p, "template");
        var pageSetup = PageSetup.ParseOptional(p, "page_setup");
        var headerFooter = HeaderFooter.ParseOptional(p, "header_footer");
        var defaultFont = FontConfig.ParseOptional(p, "default_font");
        var cover = CoverConfig.ParseOptional(p, "cover_config");
        var scene = NativeCommon.OptString(p, "scene");
        try
        {
            return CreateDocumentResult(
                p, outputPath, content, template, pageSetup, headerFooter, defaultFont,
                cover, scene);
        }
        catch (NativeUsageException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new NativeRuntimeException($"Word 文档创建失败: {ex.Message}");
        }
    }

    internal static Dictionary<string, object?> CreateDocumentResult(
        JsonElement parent,
        string outputPath,
        WordContent content,
        string? template,
        PageSetup? pageSetup,
        HeaderFooter? headerFooter,
        FontConfig? defaultFont,
        CoverConfig? cover,
        string? sceneName)
    {
        ValidateDocxOutputPath(outputPath);
        var sceneProfile = WordFileScenes.Get(sceneName);
        var effectivePageSetup = MergePageSetup(sceneProfile, pageSetup);
        var effectiveHeaderFooter = MergeHeaderFooter(sceneProfile, headerFooter);
        var effectiveFont = MergeFont(sceneProfile, defaultFont);
        var generateCover = cover is not null || sceneProfile?.ShowCover == true;
        var mode = NativeCommon.OptString(parent, "if_exists");
        var confirm = NativeCommon.OptBool(parent, "confirm_overwrite", false);
        var resolvedPath = ResolveOutputPath(outputPath, mode, confirm);
        string? backupPath = null;
        if (File.Exists(resolvedPath) && new FileInfo(resolvedPath).Length > 0)
            backupPath = NativeCommon.Backup(resolvedPath);
        WordFileWriter.CreateDocument(
            resolvedPath, content, template, effectivePageSetup, effectiveHeaderFooter,
            effectiveFont, sceneProfile, cover, generateCover);
        var size = new FileInfo(resolvedPath).Length;
        var message = $"Word 文档创建成功: {resolvedPath}"
                      + (string.IsNullOrEmpty(sceneName) ? "" : $"（场景: {sceneName}）")
                      + (generateCover ? "（含封面）" : "");
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["file_path"] = resolvedPath,
            ["file_size"] = size,
            ["sheets_count"] = null,
            ["slides_count"] = null,
            ["message"] = message,
            ["backup_path"] = backupPath,
        };
    }

    // Python word_writer._merge_* semantics: scene defaults, explicit parameters win.

    private static PageSetup? MergePageSetup(SceneProfile? scene, PageSetup? explicitSetup)
    {
        if (scene is null && explicitSetup is null)
            return null;
        return new PageSetup
        {
            Orientation = explicitSetup?.Orientation,
            PaperSize = explicitSetup?.PaperSize
                        ?? scene?.PaperSize
                        ?? "A4",
            MarginTop = explicitSetup?.MarginTop ?? scene?.MarginTop,
            MarginBottom = explicitSetup?.MarginBottom ?? scene?.MarginBottom,
            MarginLeft = explicitSetup?.MarginLeft ?? scene?.MarginLeft,
            MarginRight = explicitSetup?.MarginRight ?? scene?.MarginRight,
            Columns = explicitSetup?.Columns,
            ColumnSpacing = explicitSetup?.ColumnSpacing,
            PageBorder = explicitSetup?.PageBorder,
            PageBorderColor = explicitSetup?.PageBorderColor,
            WatermarkText = explicitSetup?.WatermarkText,
            WatermarkColor = explicitSetup?.WatermarkColor,
        };
    }

    private static HeaderFooter? MergeHeaderFooter(
        SceneProfile? scene, HeaderFooter? explicitHeaderFooter)
    {
        if (scene is null && explicitHeaderFooter is null)
            return null;
        string? footer = scene is not null && scene.PageNumberStyle != "none"
            ? "— {page} —"
            : null;
        string? header = null;
        if (explicitHeaderFooter is not null)
        {
            if (explicitHeaderFooter.HeaderText is not null)
                header = explicitHeaderFooter.HeaderText;
            if (explicitHeaderFooter.FooterText is not null)
                footer = explicitHeaderFooter.FooterText;
        }
        return new HeaderFooter
        {
            HeaderText = header,
            FooterText = footer,
        };
    }

    private static FontConfig? MergeFont(SceneProfile? scene, FontConfig? explicitFont)
    {
        if (scene is null && explicitFont is null)
            return null;
        return new FontConfig
        {
            Name = explicitFont?.Name ?? scene?.FontName,
            SizePt = explicitFont?.SizePt ?? scene?.BodySizePt,
            Bold = explicitFont?.Bold,
            Italic = explicitFont?.Italic,
            Color = explicitFont?.Color,
        };
    }

    // ── edit-word-document ──────────────────────────────────────

    public static Dictionary<string, object?> RunEditWordDocument(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        NativeCommon.EnsureFile(filePath, ".docx");
        if (!WordJson.TryGet(p, "append_sections", out var sectionsElement)
            || sectionsElement.ValueKind != JsonValueKind.Array)
            throw new NativeUsageException(
                "缺少参数 append_sections（通过 --params 或 --params-file 传入 JSON 数组）");
        var sections = sectionsElement.EnumerateArray().Select(SectionItem.Parse).ToList();
        var defaultFont = FontConfig.ParseOptional(p, "default_font");
        try
        {
            var backupPath = NativeCommon.Backup(filePath);
            WordFileWriter.AppendToDocument(filePath, sections, defaultFont);
            return new Dictionary<string, object?>
            {
                ["success"] = true,
                ["file_path"] = filePath,
                ["operation_count"] = sections.Count,
                ["message"] = $"已追加 {sections.Count} 个内容块到 Word 文档",
                ["backup_path"] = backupPath,
            };
        }
        catch (NativeUsageException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new NativeRuntimeException($"编辑 Word 文档失败: {ex.Message}");
        }
    }

    // ── word-set-properties ─────────────────────────────────────

    public static Dictionary<string, object?> RunWordSetProperties(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        NativeCommon.EnsureFile(filePath, ".docx");
        WordFileOoxml.ValidateOfficeInput(filePath);
        var wantList = NativeCommon.OptBool(p, "list", false);
        WordProperties? properties = null;
        var effectiveFields = new List<string>();
        if (WordJson.TryGet(p, "properties", out var propertiesElement)
            && propertiesElement.ValueKind != JsonValueKind.Null)
        {
            properties = WordProperties.Parse(propertiesElement, rejectUnknown: true);
            effectiveFields = properties.NonNullFields();
        }
        if (wantList || properties is null || effectiveFields.Count == 0)
        {
            return new Dictionary<string, object?>
            {
                ["success"] = true,
                ["file_path"] = filePath,
                ["properties"] = ReadProperties(filePath),
                ["backup_path"] = null,
            };
        }
        var backupPath = NativeCommon.Backup(filePath);
        try
        {
            WordFileOoxml.SaveThroughTemp(filePath, document =>
                WordFileOoxml.ApplyCoreProperties(document, properties));
        }
        catch (Exception ex) when (ex is not NativeUsageException)
        {
            throw new NativeRuntimeException($"写入 Word 属性失败: {ex.Message}");
        }
        effectiveFields.Sort(StringComparer.Ordinal);
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["file_path"] = filePath,
            ["properties"] = ReadProperties(filePath),
            ["updated_fields"] = effectiveFields,
            ["backup_path"] = backupPath,
        };
    }

    private static Dictionary<string, object?> ReadProperties(string filePath)
    {
        using var document = WordprocessingDocument.Open(filePath, false);
        return WordFileOoxml.ReadCoreProperties(document);
    }

    // ── word-manage-comments ────────────────────────────────────

    public static Dictionary<string, object?> RunWordManageComments(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        NativeCommon.EnsureFile(filePath, ".docx");
        WordFileOoxml.ValidateOfficeInput(filePath);
        var wantList = NativeCommon.OptBool(p, "list", false);
        var additions = WordCommentDef.ParseList(p, "add");
        string? backupPath = null;
        var added = 0;
        if (additions.Count > 0)
        {
            try
            {
                WordFileOoxml.SaveThroughTemp(filePath, document =>
                {
                    var main = document.MainDocumentPart
                               ?? throw new NativeRuntimeException("文档没有主部件");
                    var body = main.Document?.Body
                               ?? throw new NativeRuntimeException("文档没有正文");
                    // Resolve every anchor before mutating, like the Python entry.
                    var resolved = additions
                        .Select(comment => (Comment: comment, Runs: ResolveCommentRuns(body, comment)))
                        .ToList();
                    backupPath = NativeCommon.Backup(filePath);
                    var commentsPart = main.WordprocessingCommentsPart
                                       ?? main.AddNewPart<WordprocessingCommentsPart>();
                    commentsPart.Comments ??= new Comments();
                    var nextId = commentsPart.Comments.Elements<Comment>()
                        .Select(item => int.TryParse(item.Id?.Value,
                            NumberStyles.Integer, CultureInfo.InvariantCulture, out var id) ? id : -1)
                        .DefaultIfEmpty(-1)
                        .Max() + 1;
                    foreach (var (comment, runs) in resolved)
                    {
                        var commentId = nextId.ToString(CultureInfo.InvariantCulture);
                        commentsPart.Comments.Append(
                            BuildComment(comment, commentId));
                        InsertCommentAnchors(runs, commentId);
                        nextId++;
                    }
                    commentsPart.Comments.Save();
                    main.Document.Save();
                });
            }
            catch (NativeUsageException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new NativeRuntimeException($"添加 Word 批注失败: {ex.Message}");
            }
            added = additions.Count;
        }
        var includeComments = wantList || additions.Count == 0;
        var comments = includeComments ? ReadComments(filePath) : [];
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["file_path"] = filePath,
            ["added"] = added,
            ["comments"] = comments,
            ["backup_path"] = backupPath,
        };
    }

    private static Comment BuildComment(WordCommentDef definition, string commentId)
    {
        var now = DateTime.UtcNow;
        var comment = new Comment
        {
            Id = commentId,
            Author = definition.Author,
            Initials = definition.Initials,
            Date = new DateTime(
                now.Ticks - now.Ticks % TimeSpan.TicksPerSecond,
                DateTimeKind.Utc),
        };
        if (definition.Text.Length == 0)
        {
            comment.Append(new Paragraph());
            return comment;
        }
        var lines = definition.Text.Split('\n');
        comment.Append(new Paragraph(new Run(new Text(lines[0])
        {
            Space = SpaceProcessingModeValues.Preserve,
        })));
        foreach (var line in lines.Skip(1))
        {
            var paragraph = new Paragraph();
            if (line.Length > 0)
            {
                paragraph.Append(new Run(new Text(line)
                {
                    Space = SpaceProcessingModeValues.Preserve,
                }));
            }
            comment.Append(paragraph);
        }
        return comment;
    }

    private static List<Run> ResolveCommentRuns(Body body, WordCommentDef comment)
    {
        if (!string.IsNullOrEmpty(comment.AnchorText))
        {
            var comparison = comment.MatchCase
                ? StringComparison.Ordinal
                : StringComparison.OrdinalIgnoreCase;
            foreach (var paragraph in body.Elements<Paragraph>())
            {
                foreach (var run in paragraph.Elements<Run>())
                {
                    if (WordFileOoxml.RunText(run)
                        .Contains(comment.AnchorText, comparison))
                        return [run];
                }
            }
            throw new NativeUsageException(
                $"未找到包含锚定文字的 run：{comment.AnchorText}");
        }
        var paragraphs = body.Elements<Paragraph>().ToList();
        var index = comment.ParagraphIndex ?? 0;
        if (index < 1 || index > paragraphs.Count)
            throw new NativeUsageException(
                $"paragraph_index 超出范围：{index}（文档共 {paragraphs.Count} 个正文段落）");
        var runs = paragraphs[index - 1].Elements<Run>().ToList();
        if (runs.Count == 0)
            throw new NativeUsageException($"第 {index} 个正文段落没有 run，无法添加批注");
        return runs;
    }

    private static void InsertCommentAnchors(List<Run> runs, string commentId)
    {
        var first = runs[0];
        var last = runs[^1];
        var firstParent = first.Parent
                          ?? throw new NativeRuntimeException("批注锚定 run 没有父节点");
        var lastParent = last.Parent
                         ?? throw new NativeRuntimeException("批注锚定 run 没有父节点");
        var start = new CommentRangeStart { Id = commentId };
        firstParent.InsertBefore(start, first);
        var end = new CommentRangeEnd { Id = commentId };
        lastParent.InsertAfter(end, last);
        end.InsertAfterSelf(new Run(new CommentReference { Id = commentId }));
    }

    private static List<Dictionary<string, object?>> ReadComments(string filePath)
    {
        using var document = WordprocessingDocument.Open(filePath, false);
        var main = document.MainDocumentPart;
        var body = main?.Document?.Body;
        var result = new List<Dictionary<string, object?>>();
        var comments = main?.WordprocessingCommentsPart?.Comments;
        if (comments is null)
            return result;
        foreach (var comment in comments.Elements<Comment>())
        {
            var rawId = comment.Id?.Value ?? "0";
            int.TryParse(rawId, NumberStyles.Integer, CultureInfo.InvariantCulture, out var commentId);
            var text = string.Join("\n",
                comment.Elements<Paragraph>().Select(WordFileOoxml.ParagraphText));
            result.Add(new Dictionary<string, object?>
            {
                ["comment_id"] = commentId,
                ["author"] = comment.Author?.Value,
                ["initials"] = comment.Initials?.Value,
                ["text"] = text,
                ["timestamp"] = WordFileOoxml.IsoUtc(comment.Date?.Value),
                ["anchor_text"] = body is null ? null : ExtractAnchorText(body, rawId),
            });
        }
        return result;
    }

    private static string? ExtractAnchorText(Body body, string commentId)
    {
        var capturing = false;
        var builder = new StringBuilder();
        foreach (var element in body.Descendants())
        {
            if (element is CommentRangeStart start)
            {
                if (start.Id?.Value == commentId)
                    capturing = true;
                continue;
            }
            if (element is CommentRangeEnd end)
            {
                if (end.Id?.Value == commentId)
                    break;
                continue;
            }
            if (capturing && element is Text text && text.Text is not null)
                builder.Append(text.Text);
        }
        return builder.Length > 0 ? builder.ToString() : null;
    }

    // ── merge-office-template (docx branch) ─────────────────────

    public static Dictionary<string, object?> RunMergeOfficeTemplate(JsonElement p) =>
        WordFileTemplate.RunMergeOfficeTemplate(p);

    // ── shared output resolution ────────────────────────────────

    internal static void ValidateDocxOutputPath(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new NativeRuntimeException("输出路径不能为空");
        if (!Path.GetExtension(path).Equals(".docx", StringComparison.OrdinalIgnoreCase))
            throw new NativeRuntimeException("输出文件必须使用以下扩展名之一: .docx");
    }

    internal static string ResolveOutputPath(string path, string? modeRaw, bool confirm)
    {
        var mode = modeRaw ?? "error";
        if (mode is not ("error" or "overwrite" or "rename"))
            throw new NativeUsageException(
                $"if_exists 必须是 error, overwrite, rename 之一，收到 '{mode}'");
        var fullPath = Path.GetFullPath(path);
        var existsNonEmpty = File.Exists(fullPath) && new FileInfo(fullPath).Length > 0;
        if (!existsNonEmpty)
            return path;
        if (mode == "rename")
            return MakeRenamedPath(fullPath);
        NativeCommon.EnsureOutput(path, mode == "overwrite", confirm);
        return path;
    }

    private static string MakeRenamedPath(string fullPath)
    {
        var directory = Path.GetDirectoryName(fullPath)!;
        var stem = Path.GetFileNameWithoutExtension(fullPath);
        var extension = Path.GetExtension(fullPath);
        for (var counter = 1; ; counter++)
        {
            var candidate = Path.Combine(directory, $"{stem} ({counter}){extension}");
            if (!File.Exists(candidate))
                return candidate;
        }
    }
}
