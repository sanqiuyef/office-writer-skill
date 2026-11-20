using System.Text.Json;
using System.Text.RegularExpressions;

namespace OfficeNative;

internal static class OpenReplaceCommands
{
    private sealed record ReplaceSection(
        string HeadingText, string HeadingMatch, string? EndHeadingText,
        bool KeepHeading, List<string> Lines, FontConfig? Font);

    private sealed record FontConfig(
        string? Name, object? Size, bool? Bold, bool? Italic);

    private sealed record TextReplacement(string Find, string Replace, bool MatchCase);

    private static readonly HashSet<string> WordExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".doc", ".docx", ".dot", ".dotx", ".docm",
    };
    private static readonly HashSet<string> PptExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".ppt", ".pptx", ".pot", ".potx", ".pptm",
    };
    private static readonly HashSet<string> ExcelExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".xls", ".xlsx", ".xlt", ".xltx", ".xlsm",
    };

    public static Dictionary<string, object?> RunReplace(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        var sections = ParseSections(p);
        var replacements = ParseReplacements(p);
        if (sections.Count == 0 && replacements.Count == 0)
            throw new NativeUsageException("replace_sections 与 replace_texts 至少提供一个");
        NativeCommon.EnsureFile(filePath);

        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        if (WordExtensions.Contains(extension))
            return ReplaceWord(filePath, sections, replacements);
        if (PptExtensions.Contains(extension))
        {
            if (sections.Count > 0)
                throw new NativeRuntimeException("PPTX 不支持 replace_sections，请用 replace_texts");
            return ReplacePowerPoint(filePath, replacements);
        }
        if (ExcelExtensions.Contains(extension))
        {
            if (sections.Count > 0)
                throw new NativeRuntimeException("XLSX 不支持 replace_sections，请用 replace_texts");
            return ReplaceExcel(filePath, replacements);
        }
        throw new NativeRuntimeException($"不支持的文件类型: {extension}");
    }

    private static List<ReplaceSection> ParseSections(JsonElement p)
    {
        var result = new List<ReplaceSection>();
        if (p.ValueKind != JsonValueKind.Object || !p.TryGetProperty("replace_sections", out var array)
            || array.ValueKind != JsonValueKind.Array)
            return result;
        foreach (var item in array.EnumerateArray())
        {
            var heading = NativeCommon.OptString(item, "heading_text");
            if (string.IsNullOrWhiteSpace(heading))
                throw new NativeUsageException("replace_sections[].heading_text 不能为空");
            var lines = NativeCommon.OptStringList(item, "lines");
            FontConfig? font = null;
            if (item.TryGetProperty("font", out var fontElement) && fontElement.ValueKind == JsonValueKind.Object)
            {
                font = new FontConfig(
                    NativeCommon.OptString(fontElement, "name"),
                    OptScalar(fontElement, "size"),
                    OptBool(fontElement, "bold"),
                    OptBool(fontElement, "italic"));
            }
            var headingMatch = NativeCommon.OptString(item, "heading_match") ?? "contains";
            if (headingMatch is not ("exact" or "contains"))
                throw new NativeUsageException("heading_match 必须是 exact/contains");
            result.Add(new ReplaceSection(
                heading,
                headingMatch,
                NativeCommon.OptString(item, "end_heading_text"),
                NativeCommon.OptBool(item, "keep_heading", true),
                lines, font));
        }
        return result;
    }

    private static List<TextReplacement> ParseReplacements(JsonElement p)
    {
        var result = new List<TextReplacement>();
        if (p.ValueKind != JsonValueKind.Object || !p.TryGetProperty("replace_texts", out var array)
            || array.ValueKind != JsonValueKind.Array)
            return result;
        foreach (var item in array.EnumerateArray())
        {
            var find = NativeCommon.OptString(item, "find");
            if (string.IsNullOrEmpty(find))
                throw new NativeUsageException("replace_texts[].find 不能为空");
            result.Add(new TextReplacement(
                find,
                NativeCommon.OptString(item, "replace") ?? string.Empty,
                NativeCommon.OptBool(item, "match_case", false)));
        }
        return result;
    }

    // ── Word ─────────────────────────────────────────────────────

    private static Dictionary<string, object?> ReplaceWord(
        string filePath, List<ReplaceSection> sections, List<TextReplacement> replacements)
    {
        dynamic? app = FindWord(filePath, out var document, out var progId)
            ?? throw new NativeRuntimeException(
                $"未在 Word/WPS 中找到已打开文档: {filePath}；请先打开文件");
        try
        {
            if ((bool)document.ReadOnly)
                throw new NativeRuntimeException("活动 Word 文档为只读，无法保存");
            dynamic? undo = null;
            try { undo = app.UndoRecord; undo.StartCustomRecord("office-native replace"); }
            catch { undo = null; }
            var changed = 0;
            try
            {
                foreach (var section in sections)
                {
                    ReplaceWordSection(document, section);
                    changed++;
                }
                changed += ReplaceWordTexts(document, replacements);
                document.Save();
            }
            catch
            {
                try { if (undo is not null) undo.EndCustomRecord(); } catch { }
                try { app.Undo(); } catch { }
                throw;
            }
            try { if (undo is not null) undo.EndCustomRecord(); } catch { }
            return EditResult(filePath, changed,
                $"已通过 {progId} 替换 {changed} 处并保存");
        }
        finally
        {
            NativeCommon.ReleaseCom(document);
            NativeCommon.ReleaseCom(app);
        }
    }

    private static dynamic? FindWord(string filePath, out dynamic? document, out string progId)
    {
        foreach (var candidate in new[] { "Word.Application", "kwps.Application" })
        {
            dynamic? app;
            try { app = GenericOpenCommands.GetActive(candidate); }
            catch { app = null; }
            if (app is null)
                continue;
            try
            {
                dynamic documents = app.Documents;
                for (var index = 1; index <= (int)documents.Count; index++)
                {
                    dynamic current = documents.Item(index);
                    if (PathsEqual((string)current.FullName, filePath))
                    {
                        document = current;
                        progId = candidate;
                        return app;
                    }
                }
            }
            catch { }
            NativeCommon.ReleaseCom(app);
        }
        document = null;
        progId = string.Empty;
        return null;
    }

    private static void ReplaceWordSection(dynamic document, ReplaceSection section)
    {
        var headingIndex = FindWordParagraph(document, section.HeadingText, section.HeadingMatch)
            ?? throw new NativeRuntimeException($"未找到标题: {section.HeadingText}");
        dynamic heading = document.Paragraphs(headingIndex);
        var level = WordOutlineLevel(heading);
        var paragraphCount = (int)document.Paragraphs.Count;
        var endIndex = paragraphCount + 1;
        if (!string.IsNullOrWhiteSpace(section.EndHeadingText))
        {
            var found = FindWordParagraph(document, section.EndHeadingText!, "contains")
                ?? throw new NativeRuntimeException($"未找到结束标题: {section.EndHeadingText}");
            if (found <= headingIndex)
                throw new NativeRuntimeException($"未找到结束标题: {section.EndHeadingText}");
            endIndex = found;
        }
        else
        {
            for (var index = headingIndex + 1; index <= paragraphCount; index++)
            {
                if (WordOutlineLevel(document.Paragraphs(index)) <= level)
                {
                    endIndex = index;
                    break;
                }
            }
        }
        var start = section.KeepHeading ? (int)heading.Range.End : (int)heading.Range.Start;
        var end = endIndex <= paragraphCount
            ? (int)document.Paragraphs(endIndex).Range.Start
            : (int)document.Content.End;
        if (end > start)
            document.Range(start, end).Delete();
        if (section.Lines.Count > 0)
        {
            var payload = string.Join("\r", section.Lines) + "\r";
            var range = document.Range(start, start);
            range.InsertAfter(payload);
            ApplyWordFont(document.Range(start, start + payload.Length), section.Font);
        }
    }

    private static int? FindWordParagraph(dynamic document, string text, string match)
    {
        var count = (int)document.Paragraphs.Count;
        for (var index = 1; index <= count; index++)
        {
            var current = WordText(document.Paragraphs(index));
            if (match == "exact")
            {
                if (string.Equals(current, text.Trim(), StringComparison.Ordinal))
                    return index;
            }
            else if (current.Contains(text.Trim(), StringComparison.Ordinal))
            {
                return index;
            }
        }
        return null;
    }

    private static string WordText(dynamic paragraph)
    {
        try { return ((string)paragraph.Range.Text).Trim(); }
        catch { return string.Empty; }
    }

    private static int WordOutlineLevel(dynamic paragraph)
    {
        try { return (int)paragraph.OutlineLevel; }
        catch { return 10; }
    }

    private static void ApplyWordFont(dynamic range, FontConfig? font)
    {
        if (font is null)
            return;
        if (!string.IsNullOrWhiteSpace(font.Name))
        {
            range.Font.NameFarEast = font.Name;
            var western = font.Name is "宋体" or "黑体" or "楷体" or "仿宋"
                ? "Times New Roman"
                : font.Name;
            range.Font.NameAscii = western;
            range.Font.NameOther = western;
        }
        if (font.Size is not null)
        {
            try { range.Font.Size = ChineseSizeToPoints(font.Size.ToString()!); }
            catch { }
        }
        if (font.Bold is not null)
            range.Font.Bold = font.Bold.Value ? -1 : 0;
        if (font.Italic is not null)
            range.Font.Italic = font.Italic.Value ? -1 : 0;
    }

    private static int ReplaceWordTexts(dynamic document, List<TextReplacement> replacements)
    {
        var changed = 0;
        foreach (var item in replacements)
        {
            dynamic finder = document.Content.Find;
            finder.ClearFormatting();
            finder.Replacement.ClearFormatting();
            finder.Text = item.Find;
            finder.Replacement.Text = item.Replace;
            finder.MatchCase = item.MatchCase;
            finder.Forward = true;
            finder.Wrap = 1; // wdFindContinue
            if ((bool)finder.Execute(item.Find, item.MatchCase, false, false, false, false,
                    true, 1, false, item.Replace, 2,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing))
                changed++;
        }
        return changed;
    }

    // ── PowerPoint ──────────────────────────────────────────────

    private static Dictionary<string, object?> ReplacePowerPoint(
        string filePath, List<TextReplacement> replacements)
    {
        dynamic? app = FindPowerPoint(filePath, out var presentation, out var progId)
            ?? throw new NativeRuntimeException(
                $"未在 PowerPoint/WPS 中找到已打开演示文稿: {filePath}；请先打开文件");
        try
        {
            var changed = 0;
            foreach (var slide in presentation.Slides)
            {
                foreach (var shape in slide.Shapes)
                {
                    try
                    {
                        if (!(bool)shape.HasTextFrame || !(bool)shape.TextFrame.HasText)
                            continue;
                        dynamic textRange = shape.TextFrame.TextRange;
                        var original = (string)textRange.Text;
                        var updated = original;
                        foreach (var item in replacements)
                        {
                            updated = item.MatchCase
                                ? updated.Replace(item.Find, item.Replace, StringComparison.Ordinal)
                                : Regex.Replace(updated, Regex.Escape(item.Find),
                                    item.Replace.Replace("$", "$$"), RegexOptions.IgnoreCase);
                        }
                        if (updated != original)
                        {
                            textRange.Text = updated;
                            changed++;
                        }
                    }
                    catch { }
                }
            }
            try
            {
                presentation.Save();
            }
            catch (Exception saveError)
            {
                try { presentation.SaveAs(presentation.FullName); }
                catch
                {
                    throw new NativeRuntimeException(
                        "PowerPoint 保存被拒绝（文件可能被占用），请在 PowerPoint 中按 Ctrl+S 手工保存: "
                        + saveError.Message);
                }
            }
            return EditResult(filePath, changed, $"已通过 {progId} 替换 {changed} 处文本并保存");
        }
        finally
        {
            NativeCommon.ReleaseCom(presentation);
            NativeCommon.ReleaseCom(app);
        }
    }

    private static dynamic? FindPowerPoint(
        string filePath, out dynamic? presentation, out string progId)
    {
        foreach (var candidate in new[] { "PowerPoint.Application", "KWPP.Application" })
        {
            dynamic? app;
            try { app = GenericOpenCommands.GetActive(candidate); }
            catch { app = null; }
            if (app is null)
                continue;
            try
            {
                dynamic presentations = app.Presentations;
                for (var index = 1; index <= (int)presentations.Count; index++)
                {
                    dynamic current = presentations.Item(index);
                    if (PathsEqual((string)current.FullName, filePath))
                    {
                        presentation = current;
                        progId = candidate;
                        return app;
                    }
                }
            }
            catch { }
            NativeCommon.ReleaseCom(app);
        }
        presentation = null;
        progId = string.Empty;
        return null;
    }

    // ── Excel ────────────────────────────────────────────────────

    private static Dictionary<string, object?> ReplaceExcel(
        string filePath, List<TextReplacement> replacements)
    {
        dynamic? app = FindExcel(filePath, out var workbook, out var progId)
            ?? throw new NativeRuntimeException(
                $"未在 Excel/WPS 中找到已打开工作簿: {filePath}；请先打开文件");
        try
        {
            var changed = 0;
            foreach (var sheet in workbook.Worksheets)
            {
                foreach (var item in replacements)
                {
                    try
                    {
                        dynamic used = sheet.UsedRange;
                        if ((int)used.Cells.Count <= 0)
                            continue;
                        if ((bool)used.Replace(item.Find, item.Replace, 2, 1, item.MatchCase))
                            changed++;
                    }
                    catch { }
                }
            }
            workbook.Save();
            return EditResult(filePath, changed, $"已通过 {progId} 替换 {changed} 处文本并保存");
        }
        finally
        {
            NativeCommon.ReleaseCom(workbook);
            NativeCommon.ReleaseCom(app);
        }
    }

    private static dynamic? FindExcel(string filePath, out dynamic? workbook, out string progId)
    {
        foreach (var candidate in new[] { "Excel.Application", "KET.Application", "et.Application" })
        {
            dynamic? app;
            try { app = GenericOpenCommands.GetActive(candidate); }
            catch { app = null; }
            if (app is null)
                continue;
            try
            {
                dynamic workbooks = app.Workbooks;
                for (var index = 1; index <= (int)workbooks.Count; index++)
                {
                    dynamic current = workbooks.Item(index);
                    if (PathsEqual((string)current.FullName, filePath))
                    {
                        workbook = current;
                        progId = candidate;
                        return app;
                    }
                }
            }
            catch { }
            NativeCommon.ReleaseCom(app);
        }
        workbook = null;
        progId = string.Empty;
        return null;
    }

    // ── 通用 ────────────────────────────────────────────────────

    private static bool PathsEqual(string left, string right)
    {
        try
        {
            return string.Equals(
                Path.GetFullPath(left).Replace('/', '\\'),
                Path.GetFullPath(right).Replace('/', '\\'),
                StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
        }
    }

    private static object? OptScalar(JsonElement element, string name)
    {
        if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(name, out var value))
            return null;
        return value.ValueKind switch
        {
            JsonValueKind.String => value.GetString(),
            JsonValueKind.Number => value.GetDouble(),
            _ => null,
        };
    }

    private static bool? OptBool(JsonElement element, string name)
    {
        if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(name, out var value))
            return null;
        return value.ValueKind switch
        {
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.String when bool.TryParse(value.GetString(), out var parsed) => parsed,
            _ => null,
        };
    }

    private static int ChineseSizeToPoints(string size)
    {
        if (int.TryParse(size, out var points))
            return points;
        return size switch
        {
            "初号" => 42, "小初" => 36, "一号" => 26, "小一" => 24,
            "二号" => 22, "小二" => 18, "三号" => 16, "小三" => 15,
            "四号" => 14, "小四" => 12, "五号" => 10, "小五" => 9,
            "六号" => 8, "小六" => 7,
            _ => 12,
        };
    }

    private static Dictionary<string, object?> EditResult(
        string filePath, int operationCount, string message) => new()
    {
        ["success"] = true,
        ["file_path"] = filePath,
        ["operation_count"] = operationCount,
        ["message"] = message,
        ["backup_path"] = null,
    };
}
