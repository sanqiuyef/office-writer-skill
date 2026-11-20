using System.Globalization;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeNative;

internal static class WordCommands
{
    // ── 修订 ─────────────────────────────────────────────────────

    public static Dictionary<string, object?> RunRevisions(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        var action = NativeCommon.OptString(p, "action") ?? "list";
        NativeCommon.EnsureFile(filePath, ".docx");
        return action switch
        {
            "list" => ListRevisions(filePath),
            "accept-all" => ResolveRevisions(filePath, accept: true),
            "reject-all" => ResolveRevisions(filePath, accept: false),
            "track-insert" => TrackInsert(filePath, p),
            "track-delete" => TrackDelete(filePath, p),
            _ => throw new NativeUsageException(
                "action 必须是 list/accept-all/reject-all/track-insert/track-delete"),
        };
    }

    private static Dictionary<string, object?> ListRevisions(string filePath)
    {
        using var doc = WordprocessingDocument.Open(filePath, false);
        var roots = RevisionStoryRoots(doc);
        var insertions = new List<Dictionary<string, object?>>();
        var deletions = new List<Dictionary<string, object?>>();
        var rowInsertions = new List<Dictionary<string, object?>>();
        var rowDeletions = new List<Dictionary<string, object?>>();

        foreach (var root in roots)
        {
            insertions.AddRange(root.Descendants<InsertedRun>().Select(item =>
                new Dictionary<string, object?>
                {
                    ["type"] = "insert",
                    ["id"] = item.Id?.Value,
                    ["author"] = item.Author?.Value,
                    ["date"] = item.Date?.Value.ToString("o"),
                    ["text"] = string.Concat(
                        item.Descendants<Text>().Select(text => text.Text)),
                }));
            deletions.AddRange(root.Descendants<DeletedRun>().Select(item =>
                new Dictionary<string, object?>
                {
                    ["type"] = "delete",
                    ["id"] = item.Id?.Value,
                    ["author"] = item.Author?.Value,
                    ["date"] = item.Date?.Value.ToString("o"),
                    ["text"] = string.Concat(
                        item.Descendants<DeletedText>().Select(text => text.Text)),
                }));
            foreach (var row in root.Descendants<TableRow>())
            {
                var properties = row.TableRowProperties;
                if (properties is null)
                    continue;
                if (properties.GetFirstChild<Inserted>() is { } inserted)
                    rowInsertions.Add(RowRevisionInfo("insert_row", inserted, row));
                if (properties.GetFirstChild<Deleted>() is { } deleted)
                    rowDeletions.Add(RowRevisionInfo("delete_row", deleted, row));
            }
        }

        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["file_path"] = Path.GetFullPath(filePath),
            ["insertions"] = insertions.Count,
            ["deletions"] = deletions.Count,
            ["row_insertions"] = rowInsertions.Count,
            ["row_deletions"] = rowDeletions.Count,
            ["revisions"] = insertions.Concat(deletions)
                .Concat(rowInsertions).Concat(rowDeletions).ToList(),
        };
    }

    private static Dictionary<string, object?> RowRevisionInfo(
        string type, TrackChangeType marker, TableRow row)
    {
        var text = string.Concat(row.Descendants<Text>().Select(item => item.Text));
        if (text.Length == 0)
        {
            text = string.Concat(
                row.Descendants<DeletedText>().Select(item => item.Text));
        }
        return new Dictionary<string, object?>
        {
            ["type"] = type,
            ["id"] = marker.Id?.Value,
            ["author"] = marker.Author?.Value,
            ["date"] = marker.Date?.Value.ToString("o"),
            ["text"] = text,
        };
    }

    private static Dictionary<string, object?> ResolveRevisions(string filePath, bool accept)
    {
        var backupPath = NativeCommon.Backup(filePath);
        var insertionsResolved = 0;
        var deletionsResolved = 0;
        var rowInsertionsResolved = 0;
        var rowDeletionsResolved = 0;
        WordFileOoxml.SaveThroughTemp(filePath, document =>
        {
            var roots = RevisionStoryRoots(document);
            foreach (var root in roots)
            {
                var rows = ResolveRowRevisions(root, accept);
                rowInsertionsResolved += rows.Insertions;
                rowDeletionsResolved += rows.Deletions;
            }
            foreach (var root in roots)
            {
                if (accept)
                {
                    // Accept inner deletions before unwrapping insertions so the
                    // cloned insertion content never carries stale w:del markup.
                    var deletionCount = root.Descendants<DeletedRun>().Count();
                    foreach (var deletion in root.Descendants<DeletedRun>().ToList())
                        deletion.Remove();
                    deletionsResolved += deletionCount;
                    var insertionCount = root.Descendants<InsertedRun>().Count();
                    foreach (var insertion in root.Descendants<InsertedRun>()
                                 .OrderByDescending(item =>
                                     item.Ancestors<InsertedRun>().Count())
                                 .ToList())
                    {
                        var parent = insertion.Parent
                                     ?? throw new NativeRuntimeException(
                                         "修订节点没有父节点");
                        foreach (var child in insertion.ChildElements.ToList())
                            parent.InsertBefore(child.CloneNode(true), insertion);
                        insertion.Remove();
                    }
                    insertionsResolved += insertionCount;
                }
                else
                {
                    // Rejecting an insertion removes its whole subtree, so count
                    // nested deletions before discarding that subtree.
                    var insertionCount = root.Descendants<InsertedRun>().Count();
                    var deletionCount = root.Descendants<DeletedRun>().Count();
                    foreach (var insertion in root.Descendants<InsertedRun>().ToList())
                        insertion.Remove();
                    insertionsResolved += insertionCount;
                    foreach (var deletion in root.Descendants<DeletedRun>().ToList())
                    {
                        var parent = deletion.Parent
                                     ?? throw new NativeRuntimeException(
                                         "修订节点没有父节点");
                        foreach (var deletedText in
                                 deletion.Descendants<DeletedText>().ToList())
                        {
                            var run = new Run();
                            var sourceRun = deletedText.Ancestors<Run>().FirstOrDefault();
                            var runProperties = sourceRun
                                ?.GetFirstChild<RunProperties>()?.CloneNode(true)
                                as RunProperties;
                            if (runProperties is not null)
                                run.Append(runProperties);
                            run.Append(new Text(deletedText.Text ?? string.Empty)
                            {
                                Space = SpaceProcessingModeValues.Preserve,
                            });
                            parent.InsertBefore(run, deletion);
                        }
                        deletion.Remove();
                    }
                    deletionsResolved += deletionCount;
                }
            }
            document.MainDocumentPart!.Document?.Save();
        });
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["file_path"] = Path.GetFullPath(filePath),
            ["action"] = accept ? "accept-all" : "reject-all",
            ["insertions_resolved"] = insertionsResolved,
            ["deletions_resolved"] = deletionsResolved,
            ["row_insertions_resolved"] = rowInsertionsResolved,
            ["row_deletions_resolved"] = rowDeletionsResolved,
            ["backup_path"] = backupPath,
        };
    }

    private static Dictionary<string, object?> TrackInsert(string filePath, JsonElement p)
    {
        var text = NativeCommon.RequireString(p, "text");
        var backupPath = NativeCommon.Backup(filePath);
        string? revisionId = null;
        WordFileOoxml.SaveThroughTemp(filePath, document =>
        {
            var body = document.MainDocumentPart?.Document.Body
                       ?? throw new NativeRuntimeException("文档没有正文");
            var paragraph = FindParagraph(body, p);
            var inserted = new InsertedRun
            {
                Author = NativeCommon.OptString(p, "author") ?? "office-native",
                Date = RevisionDate(p),
                Id = NextRevisionId(body),
            };
            inserted.Append(new Run(new Text(text)
            {
                Space = SpaceProcessingModeValues.Preserve,
            }));
            paragraph.Append(inserted);
            revisionId = inserted.Id?.Value;
            document.MainDocumentPart!.Document.Save();
        });
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["file_path"] = Path.GetFullPath(filePath),
            ["action"] = "track-insert",
            ["revision_id"] = revisionId,
            ["backup_path"] = backupPath,
        };
    }

    private static Dictionary<string, object?> TrackDelete(string filePath, JsonElement p)
    {
        var anchor = NativeCommon.RequireString(p, "text");
        if (string.IsNullOrEmpty(anchor))
            throw new NativeUsageException("track-delete 的 text 不能为空");
        var matchCase = NativeCommon.OptBool(p, "match_case", false);
        var backupPath = NativeCommon.Backup(filePath);
        var comparison = matchCase
            ? StringComparison.Ordinal
            : StringComparison.OrdinalIgnoreCase;
        var deletedText = anchor;
        List<string> revisionIds = [];
        WordFileOoxml.SaveThroughTemp(filePath, document =>
        {
            var body = document.MainDocumentPart?.Document.Body
                       ?? throw new NativeRuntimeException("文档没有正文");
            var match = FindTextInSingleParagraph(body, anchor, comparison)
                        ?? throw new NativeUsageException(
                            $"无法精确定位要删除的文本：{anchor}"
                            + "（未在单个段落中找到完整匹配）");
            deletedText = match.MatchedText;
            revisionIds = DeleteTextAcrossRuns(
                body, match.Paragraph, match.Index, match.MatchedText.Length, p);
            document.MainDocumentPart!.Document.Save();
        });
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["file_path"] = Path.GetFullPath(filePath),
            ["action"] = "track-delete",
            ["deleted_text"] = deletedText,
            ["revision_id"] = revisionIds.FirstOrDefault(),
            ["revision_ids"] = revisionIds,
            ["backup_path"] = backupPath,
        };
    }

    private static (Paragraph Paragraph, int Index, string MatchedText)?
        FindTextInSingleParagraph(Body body, string anchor, StringComparison comparison)
    {
        foreach (var paragraph in body.Descendants<Paragraph>())
        {
            var text = VisibleText(paragraph);
            var index = text.IndexOf(anchor, comparison);
            if (index >= 0)
                return (paragraph, index, text.Substring(index, anchor.Length));
        }
        return null;
    }

    private static string VisibleText(Paragraph paragraph)
    {
        var builder = new StringBuilder();
        foreach (var text in paragraph.Descendants<Text>())
        {
            var owner = text.Ancestors<Paragraph>().FirstOrDefault();
            if (!ReferenceEquals(owner, paragraph))
                continue;
            builder.Append(text.Text);
        }
        return builder.ToString();
    }

    private sealed class RunTextSpan
    {
        public required Run Run { get; init; }
        public required int Start { get; init; }
        public StringBuilder Content { get; } = new();
        public int End => Start + Content.Length;
    }

    private static List<string> DeleteTextAcrossRuns(
        Body body, Paragraph paragraph, int matchStart, int matchLength, JsonElement p)
    {
        var textNodes = paragraph.Descendants<Text>()
            .Where(text => ReferenceEquals(
                text.Ancestors<Paragraph>().FirstOrDefault(), paragraph))
            .ToList();
        var spans = new List<RunTextSpan>();
        var offset = 0;
        foreach (var textNode in textNodes)
        {
            var run = textNode.Parent as Run
                      ?? throw new NativeUsageException(
                          "无法精确定位要删除的文本：文本节点不在 run 内");
            if (spans.Count == 0 || !ReferenceEquals(spans[^1].Run, run))
                spans.Add(new RunTextSpan { Run = run, Start = offset });
            var value = textNode.Text ?? string.Empty;
            spans[^1].Content.Append(value);
            offset += value.Length;
        }
        var deleteStart = matchStart;
        var deleteEnd = matchStart + matchLength;
        if (deleteStart < 0 || deleteEnd > offset)
        {
            throw new NativeUsageException(
                "无法精确定位要删除的文本：匹配区间超出段落文本范围");
        }

        var author = NativeCommon.OptString(p, "author") ?? "office-native";
        var date = RevisionDate(p);
        var revisionIds = new List<string>();
        foreach (var span in spans)
        {
            var overlapStart = Math.Max(span.Start, deleteStart);
            var overlapEnd = Math.Min(span.End, deleteEnd);
            if (overlapStart >= overlapEnd)
                continue;
            if (span.Run.ChildElements.Any(child => child is not RunProperties
                    && child is not Text && child is not LastRenderedPageBreak))
            {
                throw new NativeUsageException(
                    "无法精确定位要删除的文本：命中的 run 含有无法安全拆分的"
                    + "非文本内容（如制表符、换行、域或图片）");
            }
            var parent = span.Run.Parent
                         ?? throw new NativeUsageException(
                             "无法精确定位要删除的文本：run 没有父节点");
            var runProperties = span.Run.GetFirstChild<RunProperties>()
                ?.CloneNode(true) as RunProperties;
            var content = span.Content.ToString();
            var localStart = overlapStart - span.Start;
            var localLength = overlapEnd - overlapStart;
            if (localStart > 0)
            {
                parent.InsertBefore(
                    NewTextRun(runProperties, content[..localStart]), span.Run);
            }
            var deletion = new DeletedRun
            {
                Author = author,
                Date = date,
                Id = NextRevisionId(body),
            };
            var deletedRun = new Run();
            if (runProperties is not null)
                deletedRun.Append(runProperties.CloneNode(true));
            deletedRun.Append(new DeletedText(
                content.Substring(localStart, localLength))
            {
                Space = SpaceProcessingModeValues.Preserve,
            });
            deletion.Append(deletedRun);
            parent.InsertBefore(deletion, span.Run);
            revisionIds.Add(deletion.Id!.Value);
            var suffixStart = localStart + localLength;
            if (suffixStart < content.Length)
            {
                parent.InsertBefore(
                    NewTextRun(runProperties, content[suffixStart..]), span.Run);
            }
            span.Run.Remove();
        }
        if (revisionIds.Count == 0)
        {
            throw new NativeUsageException(
                "无法精确定位要删除的文本：没有可拆分的文本 run");
        }
        return revisionIds;
    }

    private static Run NewTextRun(RunProperties? properties, string text)
    {
        var run = new Run();
        if (properties is not null)
            run.Append(properties.CloneNode(true));
        run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }

    private static List<OpenXmlElement> RevisionStoryRoots(
        WordprocessingDocument document)
    {
        var main = document.MainDocumentPart
                   ?? throw new NativeRuntimeException("文档没有主部件");
        var roots = new List<OpenXmlElement>();
        if (main.Document?.Body is { } body)
            roots.Add(body);
        foreach (var part in main.HeaderParts)
        {
            if (part.Header is { } header)
                roots.Add(header);
        }
        foreach (var part in main.FooterParts)
        {
            if (part.Footer is { } footer)
                roots.Add(footer);
        }
        if (main.FootnotesPart?.Footnotes is { } footnotes)
            roots.Add(footnotes);
        if (main.EndnotesPart?.Endnotes is { } endnotes)
            roots.Add(endnotes);
        if (main.WordprocessingCommentsPart?.Comments is { } comments)
            roots.Add(comments);
        return roots;
    }

    private static (int Insertions, int Deletions) ResolveRowRevisions(
        OpenXmlElement root, bool accept)
    {
        var insertions = 0;
        var deletions = 0;
        foreach (var row in root.Descendants<TableRow>().ToList())
        {
            var properties = row.TableRowProperties;
            if (properties is null)
                continue;
            var inserted = properties.GetFirstChild<Inserted>();
            var deleted = properties.GetFirstChild<Deleted>();
            if (inserted is null && deleted is null)
                continue;
            if (accept ? deleted is not null : inserted is not null)
            {
                if (accept)
                    deletions++;
                else
                    insertions++;
                row.Remove();
                continue;
            }
            if (inserted is not null)
            {
                insertions++;
                inserted.Remove();
            }
            if (deleted is not null)
            {
                deletions++;
                deleted.Remove();
            }
        }
        return (insertions, deletions);
    }

    private static DateTime RevisionDate(JsonElement p)
    {
        var raw = NativeCommon.OptString(p, "date");
        return raw is not null && DateTime.TryParse(raw, CultureInfo.InvariantCulture,
            DateTimeStyles.RoundtripKind, out var parsed)
            ? parsed.ToUniversalTime()
            : DateTime.UtcNow;
    }

    private static StringValue NextRevisionId(Body body)
    {
        var ids = body.Descendants<InsertedRun>().Select(item => item.Id?.Value)
            .Concat(body.Descendants<DeletedRun>().Select(item => item.Id?.Value))
            .Where(value => value is not null)
            .Select(value => int.TryParse(value, out var parsed) ? parsed : 0)
            .DefaultIfEmpty(0);
        return (ids.Max() + 1).ToString(CultureInfo.InvariantCulture);
    }

    private static Paragraph FindParagraph(Body body, JsonElement p)
    {
        var paragraphIndex = NativeCommon.OptInt(p, "paragraph_index", 0);
        if (paragraphIndex > 0)
        {
            var paragraph = body.Descendants<Paragraph>().ElementAtOrDefault(paragraphIndex - 1);
            if (paragraph is null)
                throw new NativeRuntimeException($"paragraph_index={paragraphIndex} 超出范围");
            return paragraph;
        }
        var anchor = NativeCommon.OptString(p, "anchor_text");
        if (string.IsNullOrEmpty(anchor))
            throw new NativeUsageException("需要 anchor_text 或 paragraph_index 定位段落");
        var matchCase = NativeCommon.OptBool(p, "match_case", false);
        var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        return body.Descendants<Paragraph>()
            .FirstOrDefault(paragraph => paragraph.InnerText.Contains(anchor, comparison))
            ?? throw new NativeRuntimeException($"找不到包含文本的段落: {anchor}");
    }

    // ── 脚注 ─────────────────────────────────────────────────────

    public static Dictionary<string, object?> RunFootnotes(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        var action = NativeCommon.OptString(p, "action") ?? "list";
        NativeCommon.EnsureFile(filePath, ".docx");
        return action switch
        {
            "add" => AddFootnote(filePath, p),
            "list" => ListFootnotes(filePath),
            _ => throw new NativeUsageException("action 必须是 add/list"),
        };
    }

    private static Dictionary<string, object?> AddFootnote(string filePath, JsonElement p)
    {
        var text = NativeCommon.RequireString(p, "text");
        var backupPath = NativeCommon.Backup(filePath);
        var footnoteId = 0;
        WordFileOoxml.SaveThroughTemp(filePath, document =>
        {
            var main = document.MainDocumentPart
                       ?? throw new NativeRuntimeException("文档没有主部件");
            var body = main.Document.Body
                       ?? throw new NativeRuntimeException("文档没有正文");
            var paragraph = FindParagraph(body, p);

            var footnotesPart = main.FootnotesPart ?? main.AddNewPart<FootnotesPart>();
            footnotesPart.Footnotes ??= new Footnotes();
            var footnotes = footnotesPart.Footnotes;
            if (!footnotes.Elements<Footnote>().Any(item => item.Id?.Value == -1))
            {
                footnotes.Append(new Footnote(
                    new Paragraph(new Run(new SeparatorMark())))
                { Id = -1, Type = FootnoteEndnoteValues.Separator });
            }
            if (!footnotes.Elements<Footnote>().Any(item => item.Id?.Value == 0))
            {
                footnotes.Append(new Footnote(
                    new Paragraph(new Run(new ContinuationSeparatorMark())))
                { Id = 0, Type = FootnoteEndnoteValues.ContinuationSeparator });
            }
            var nextId = footnotes.Elements<Footnote>()
                .Select(item => item.Id?.Value ?? 0)
                .Where(value => value > 0)
                .DefaultIfEmpty(0)
                .Max() + 1;
            EnsureFootnoteStyle(main);
            footnotes.Append(new Footnote(
                new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId { Val = "FootnoteText" }),
                    new Run(
                        new RunProperties(new RunStyle { Val = "FootnoteReference" }),
                        new FootnoteReferenceMark()),
                    new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })))
            { Id = nextId });

            paragraph.Append(new Run(
                new RunProperties(new RunStyle { Val = "FootnoteReference" }),
                new FootnoteReference { Id = nextId }));
            main.Document.Save();
            footnoteId = (int)nextId;
        });
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["file_path"] = Path.GetFullPath(filePath),
            ["footnote_id"] = footnoteId,
            ["backup_path"] = backupPath,
        };
    }

    private static void EnsureFootnoteStyle(MainDocumentPart main)
    {
        var stylesPart = main.StyleDefinitionsPart ?? main.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles ??= new Styles();
        var styles = stylesPart.Styles;
        if (!styles.Elements<Style>().Any(style => style.StyleId == "FootnoteText"))
        {
            styles.Append(new Style(
                new StyleName { Val = "footnote text" },
                new BasedOn { Val = "Normal" },
                new NextParagraphStyle { Val = "Normal" },
                new StyleRunProperties(new FontSize { Val = "20" }))
            {
                Type = StyleValues.Paragraph,
                StyleId = "FootnoteText",
            });
        }
        if (!styles.Elements<Style>().Any(style => style.StyleId == "FootnoteReference"))
        {
            styles.Append(new Style(
                new StyleName { Val = "footnote reference" },
                new StyleRunProperties(new VerticalTextAlignment
                {
                    Val = VerticalPositionValues.Superscript,
                }))
            {
                Type = StyleValues.Character,
                StyleId = "FootnoteReference",
            });
        }
        styles.Save();
    }

    private static Dictionary<string, object?> ListFootnotes(string filePath)
    {
        using var doc = WordprocessingDocument.Open(filePath, false);
        var main = doc.MainDocumentPart
                   ?? throw new NativeRuntimeException("文档没有主部件");
        var footnotes = main.FootnotesPart?.Footnotes?.Elements<Footnote>()
            .Where(item => (item.Id?.Value ?? 0) > 0)
            .Select(item => new Dictionary<string, object?>
            {
                ["id"] = item.Id?.Value,
                ["text"] = string.Concat(item.Descendants<Text>().Select(text => text.Text)),
            }).ToList() ?? [];
        var references = main.Document.Body?.Descendants<FootnoteReference>().Count() ?? 0;
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["file_path"] = Path.GetFullPath(filePath),
            ["count"] = footnotes.Count,
            ["references"] = references,
            ["footnotes"] = footnotes,
        };
    }

    // ── 内容控件 ─────────────────────────────────────────────────

    public static Dictionary<string, object?> RunContentControls(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        var action = NativeCommon.OptString(p, "action") ?? "list";
        NativeCommon.EnsureFile(filePath, ".docx");
        return action switch
        {
            "add-text" => AddContentControl(filePath, p, "text"),
            "add-date" => AddContentControl(filePath, p, "date"),
            "add-dropdown" => AddContentControl(filePath, p, "dropdown"),
            "add-checkbox" => AddContentControl(filePath, p, "checkbox"),
            "list" => ListContentControls(filePath),
            "set" => SetContentControl(filePath, p),
            "remove" => RemoveContentControl(filePath, p),
            _ => throw new NativeUsageException(
                "action 必须是 add-text/add-date/add-dropdown/add-checkbox/list/set/remove"),
        };
    }

    private static Dictionary<string, object?> AddContentControl(
        string filePath, JsonElement p, string kind)
    {
        var tag = NativeCommon.RequireString(p, "tag");
        var alias = NativeCommon.OptString(p, "alias") ?? tag;
        var backupPath = NativeCommon.Backup(filePath);
        int? sdtId = null;
        WordFileOoxml.SaveThroughTemp(filePath, document =>
        {
            var body = document.MainDocumentPart?.Document.Body
                       ?? throw new NativeRuntimeException("文档没有正文");
            var nextSdtId = NextContentControlId(body);
            sdtId = nextSdtId;
            SdtProperties properties;
            string display;
            if (kind == "checkbox")
            {
                var checkedValue = NativeCommon.OptBool(p, "checked", true);
                properties = BuildCheckboxProperties(alias, tag, nextSdtId, checkedValue);
                display = checkedValue ? "☒" : "☐";
            }
            else
            {
                properties = new SdtProperties(
                    new SdtAlias { Val = alias },
                    new Tag { Val = tag },
                    new SdtId { Val = nextSdtId });
                switch (kind)
                {
                    case "text":
                        display = NativeCommon.OptString(p, "text") ?? string.Empty;
                        properties.Append(new SdtContentText());
                        break;
                    case "date":
                        var dateText = NativeCommon.OptString(p, "date")
                                       ?? DateTime.Today.ToString("yyyy-MM-dd");
                        var dateFormat = NativeCommon.OptString(p, "date_format")
                                         ?? "yyyy-MM-dd";
                        var parsedDate = DateTime.TryParse(dateText,
                            CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind,
                            out var date) ? date : DateTime.Today;
                        properties.Append(new SdtContentDate
                        {
                            FullDate = new DateTimeValue(parsedDate),
                            DateFormat = new DateFormat { Val = dateFormat },
                        });
                        display = dateText;
                        break;
                    case "dropdown":
                        var items = NativeCommon.OptStringList(p, "items");
                        if (items.Count == 0)
                            throw new NativeUsageException(
                                "add-dropdown 需要非空 items 数组");
                        var list = new SdtContentDropDownList();
                        foreach (var item in items)
                            list.Append(new ListItem { DisplayText = item, Value = item });
                        display = NativeCommon.OptString(p, "value") ?? items[0];
                        list.LastValue = display;
                        properties.Append(list);
                        break;
                    default:
                        throw new NativeUsageException($"未知内容控件类型: {kind}");
                }
            }
            var content = new SdtContentBlock(new Paragraph(new Run(
                new Text(display) { Space = SpaceProcessingModeValues.Preserve })));
            var control = new SdtBlock(properties, content);
            var paragraph = FindParagraphOrNull(body, p);
            if (paragraph is not null)
                paragraph.InsertAfterSelf(control);
            else
                WordFileOoxml.AppendBodyElement(body, control);

            document.MainDocumentPart!.Document.Save();
        });
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["file_path"] = Path.GetFullPath(filePath),
            ["action"] = $"add-{kind}",
            ["tag"] = tag,
            ["sdt_id"] = sdtId,
            ["backup_path"] = backupPath,
        };
    }

    private static Paragraph? FindParagraphOrNull(Body body, JsonElement p)
    {
        try { return FindParagraph(body, p); }
        catch (NativeUsageException) { return null; }
        catch (NativeRuntimeException) { return null; }
    }

    private static Dictionary<string, object?> ListContentControls(string filePath)
    {
        using var doc = WordprocessingDocument.Open(filePath, false);
        var body = doc.MainDocumentPart?.Document.Body
                   ?? throw new NativeRuntimeException("文档没有正文");
        var controls = body.Descendants<SdtElement>()
            .Select(item => DescribeContentControl(item))
            .Where(item => item is not null)
            .Cast<Dictionary<string, object?>>()
            .ToList();
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["file_path"] = Path.GetFullPath(filePath),
            ["count"] = controls.Count,
            ["controls"] = controls,
        };
    }

    private static Dictionary<string, object?>? DescribeContentControl(SdtElement control)
    {
        var properties = control.SdtProperties;
        if (properties is null)
            return null;
        var textControl = properties.GetFirstChild<SdtContentText>();
        var dateControl = properties.GetFirstChild<SdtContentDate>();
        var dropDown = properties.GetFirstChild<SdtContentDropDownList>();
        var checkBox = FindCheckbox(properties);
        var kind = textControl is not null ? "text"
            : dateControl is not null ? "date"
            : dropDown is not null ? "dropdown"
            : checkBox is not null ? "checkbox"
            : "other";
        var items = dropDown?.Elements<ListItem>()
            .Select(item => item.DisplayText?.Value ?? item.Value?.Value ?? string.Empty)
            .ToList();
        return new Dictionary<string, object?>
        {
            ["tag"] = properties.GetFirstChild<Tag>()?.Val?.Value,
            ["alias"] = properties.GetFirstChild<SdtAlias>()?.Val?.Value,
            ["id"] = properties.GetFirstChild<SdtId>()?.Val?.Value,
            ["type"] = kind,
            ["text"] = string.Concat(control.Descendants<Text>().Select(text => text.Text)),
            ["items"] = items,
            ["checked"] = checkBox is not null && CheckboxChecked(checkBox),
        };
    }

    private static Dictionary<string, object?> SetContentControl(string filePath, JsonElement p)
    {
        var tag = NativeCommon.RequireString(p, "tag");
        var value = NativeCommon.OptString(p, "value")
                    ?? NativeCommon.OptString(p, "text")
                    ?? string.Empty;
        var backupPath = NativeCommon.Backup(filePath);
        WordFileOoxml.SaveThroughTemp(filePath, document =>
        {
            var body = document.MainDocumentPart?.Document.Body
                       ?? throw new NativeRuntimeException("文档没有正文");
            var control = FindContentControl(body, tag);
            var properties = control.SdtProperties!;
            var checkBox = FindCheckbox(properties);
            if (checkBox is not null)
            {
                var isChecked = NativeCommon.OptBool(p, "checked", true);
                SetCheckbox(properties, isChecked);
                value = isChecked ? "☒" : "☐";
            }
            else if (properties.GetFirstChild<SdtContentDropDownList>() is { } dropDown)
            {
                dropDown.LastValue = value;
            }
            else if (properties.GetFirstChild<SdtContentDate>() is { } dateControl
                     && DateTime.TryParse(value, CultureInfo.InvariantCulture,
                         DateTimeStyles.RoundtripKind, out var date))
            {
                dateControl.FullDate = new DateTimeValue(date);
            }
            ReplaceContentControlContent(control, value);
            document.MainDocumentPart!.Document.Save();
        });
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["file_path"] = Path.GetFullPath(filePath),
            ["tag"] = tag,
            ["value"] = value,
            ["backup_path"] = backupPath,
        };
    }

    private static void ReplaceContentControlContent(SdtElement control, string value)
    {
        var text = new Text(value) { Space = SpaceProcessingModeValues.Preserve };
        switch (control)
        {
            case SdtBlock block:
                foreach (var content in block.Elements<SdtContentBlock>().ToList())
                    content.Remove();
                block.Append(new SdtContentBlock(new Paragraph(new Run(text))));
                break;
            case SdtRun run:
                foreach (var content in run.Elements<SdtContentRun>().ToList())
                    content.Remove();
                run.Append(new SdtContentRun(new Run(text)));
                break;
            case SdtCell cell:
            {
                var content = cell.Elements<SdtContentCell>().FirstOrDefault();
                if (content is null)
                {
                    content = new SdtContentCell();
                    cell.Append(content);
                }
                var tableCell = content.Elements<TableCell>().FirstOrDefault();
                if (tableCell is null)
                {
                    tableCell = new TableCell();
                    content.Append(tableCell);
                }
                SetCellText(tableCell, text);
                break;
            }
            case SdtRow row:
            {
                var content = row.Elements<SdtContentRow>().FirstOrDefault();
                if (content is null)
                {
                    content = new SdtContentRow();
                    row.Append(content);
                }
                var tableRow = content.Elements<TableRow>().FirstOrDefault();
                if (tableRow is null)
                {
                    tableRow = new TableRow();
                    content.Append(tableRow);
                }
                var tableCell = tableRow.Elements<TableCell>().FirstOrDefault();
                if (tableCell is null)
                {
                    tableCell = new TableCell();
                    tableRow.Append(tableCell);
                }
                SetCellText(tableCell, text);
                break;
            }
            default:
                throw new NativeRuntimeException(
                    $"不支持的内容控件类型: {control.GetType().Name}");
        }
    }

    private static void SetCellText(TableCell cell, Text text)
    {
        foreach (var child in cell.ChildElements.ToList())
        {
            if (child is TableCellProperties)
                continue;
            child.Remove();
        }
        cell.Append(new Paragraph(new Run(text)));
    }

    private static Dictionary<string, object?> RemoveContentControl(string filePath, JsonElement p)
    {
        var tag = NativeCommon.RequireString(p, "tag");
        var backupPath = NativeCommon.Backup(filePath);
        WordFileOoxml.SaveThroughTemp(filePath, document =>
        {
            var body = document.MainDocumentPart?.Document.Body
                       ?? throw new NativeRuntimeException("文档没有正文");
            var control = FindContentControl(body, tag);
            control.Remove();
            document.MainDocumentPart!.Document.Save();
        });
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["file_path"] = Path.GetFullPath(filePath),
            ["tag"] = tag,
            ["removed"] = true,
            ["backup_path"] = backupPath,
        };
    }

    private static SdtElement FindContentControl(Body body, string tag) =>
        body.Descendants<SdtElement>().FirstOrDefault(
            item => item.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value == tag)
        ?? throw new NativeRuntimeException($"找不到 tag={tag} 的内容控件");

    private static int NextContentControlId(Body body) =>
        body.Descendants<SdtElement>()
            .Select(item => item.SdtProperties?.GetFirstChild<SdtId>()?.Val?.Value ?? 0)
            .DefaultIfEmpty(0)
            .Max() + 1;

    // SDK 3.5 未建模 w14:checkbox，用原始 XML 注入；读写都按本地名+命名空间识别。
    private const string W14Namespace = "http://schemas.microsoft.com/office/word/2010/wordml";

    private static OpenXmlElement? FindCheckbox(SdtProperties properties) =>
        properties.ChildElements.FirstOrDefault(element => element.LocalName == "checkbox"
                                                           && element.NamespaceUri == W14Namespace);

    private static SdtProperties BuildCheckboxProperties(
        string alias, string tag, int sdtId, bool isChecked)
    {
        var xml = "<w:sdtPr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
                  + "xmlns:w14=\"" + W14Namespace + "\">"
                  + "<w:alias w:val=\"" + System.Security.SecurityElement.Escape(alias) + "\"/>"
                  + "<w:tag w:val=\"" + System.Security.SecurityElement.Escape(tag) + "\"/>"
                  + "<w:id w:val=\"" + sdtId + "\"/>"
                  + "<w14:checkbox><w14:checked w14:val=\"" + (isChecked ? "1" : "0") + "\"/>"
                  + "<w14:checkedState w14:val=\"2612\" w14:font=\"MS Gothic\"/>"
                  + "<w14:uncheckedState w14:val=\"2610\" w14:font=\"MS Gothic\"/>"
                  + "</w14:checkbox></w:sdtPr>";
        return new SdtProperties(xml);
    }

    private static bool CheckboxChecked(OpenXmlElement checkBox)
    {
        var checkedElement = checkBox.ChildElements
            .FirstOrDefault(element => element.LocalName == "checked");
        var value = checkedElement?.GetAttributes()
            .FirstOrDefault(attribute => attribute.LocalName == "val").Value;
        return value is "1" or "true" or "on";
    }

    private static void SetCheckbox(OpenXmlElement checkBox, bool isChecked)
    {
        var checkedElement = checkBox.ChildElements
            .FirstOrDefault(element => element.LocalName == "checked");
        checkedElement?.SetAttribute(new OpenXmlAttribute(
            "w14", "val", W14Namespace, isChecked ? "1" : "0"));
    }
}
