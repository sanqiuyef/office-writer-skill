using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeNative;

/// <summary>Word 批量编辑：/body/p[N]、/body/tbl[N]、行/单元格路径的 set/add/remove/move/replace。</summary>
internal static class BatchEditWord
{
    private static readonly Regex ParagraphPath = new(@"^/body/p\[(\d+)\]$");
    private static readonly Regex TablePath = new(@"^/body/tbl\[(\d+)\]$");
    private static readonly Regex CellPath =
        new(@"^/body/tbl\[(\d+)\]/row\[(\d+)\]/cell\[(\d+)\]$");

    private static readonly HashSet<string> ParagraphProps = new(StringComparer.Ordinal)
    {
        "text", "style", "bold", "italic", "font_name", "font_size",
        "font_color", "alignment",
    };

    private static readonly string[] ParagraphOrder =
    [
        "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
        "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
        "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
        "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
        "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
        "suppressOverlap", "jc", "textDirection", "textAlignment",
        "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
        "pPrChange",
    ];

    private static readonly string[] RunOrder =
    [
        "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
        "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
        "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
        "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
        "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
        "eastAsianLayout", "specVanish", "oMath",
    ];

    public static List<string> Apply(
        string workPath, IReadOnlyList<OfficeOperation> operations)
    {
        using var document = WordprocessingDocument.Open(workPath, true);
        var main = document.MainDocumentPart
                   ?? throw new NativeRuntimeException("文档没有 MainDocumentPart");
        var body = main.Document?.Body
                   ?? throw new NativeRuntimeException("文档没有正文");
        var messages = new List<string>(operations.Count);
        foreach (var operation in operations)
        {
            try
            {
                messages.Add(ApplyOperation(main, body, operation));
            }
            catch (BatchEditException ex)
            {
                throw new BatchEditException(
                    ex.Code,
                    $"第 {operation.Index + 1} 个操作 (op={operation.Op}, path={operation.Path}) 失败: {ex.Message}",
                    ex.Path ?? operation.Path, ex.Suggestion, ex.ValidValues);
            }
            catch (Exception ex) when (ex is not NativeRuntimeException
                                       and not NativeUsageException)
            {
                throw new BatchEditException(
                    "operation_failed",
                    $"第 {operation.Index + 1} 个操作 (op={operation.Op}, path={operation.Path}) 执行失败: {ex.Message}",
                    operation.Path);
            }
        }
        main.Document.Save();
        return messages;
    }

    private static string ApplyOperation(
        MainDocumentPart main, W.Body body, OfficeOperation operation)
    {
        var (kind, element) = Resolve(body, operation.Path);
        return operation.Op switch
        {
            "set" => Set(main, kind, element, operation),
            "add" => Add(main, body, kind, operation),
            "remove" => Remove(kind, element, operation),
            "move" => Move(body, kind, element, operation),
            "replace" => Replace(kind, element, operation),
            _ => throw new BatchEditException(
                "unsupported_operation", $"不支持的操作: {operation.Op}",
                operation.Path),
        };
    }

    // ── 路径解析 ─────────────────────────────────────────────────

    private static (string Kind, OpenXmlElement Element) Resolve(
        W.Body body, string path)
    {
        var match = ParagraphPath.Match(path);
        if (match.Success)
        {
            var paragraphs = body.Elements<W.Paragraph>().ToList();
            return ("paragraph", OneBased(paragraphs,
                int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture), path));
        }
        match = TablePath.Match(path);
        if (match.Success)
        {
            var tables = body.Elements<W.Table>().ToList();
            return ("table", OneBased(tables,
                int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture), path));
        }
        match = CellPath.Match(path);
        if (match.Success)
        {
            var tables = body.Elements<W.Table>().ToList();
            var table = OneBased(tables,
                int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture), path);
            var rows = table.Elements<W.TableRow>().ToList();
            var rowIndex = int.Parse(match.Groups[2].Value,
                CultureInfo.InvariantCulture);
            OneBased(rows, rowIndex, path);
            // python-docx 按布局网格展开 gridSpan，并把 vMerge=continue 指向顶层
            // restart 单元格；cell[N] 必须使用同一解析结果。
            var cells = WordTableGrid.ExpandRow(rows, rowIndex - 1);
            return ("cell", OneBased(cells,
                int.Parse(match.Groups[3].Value, CultureInfo.InvariantCulture), path));
        }
        if (path == "/body")
            return ("body", body);
        throw new BatchEditException(
            "invalid_path", $"不支持的 Word 路径: {path}", path,
            suggestion: "使用 inspect_office_document 获取路径",
            validValues:
            [
                "/body", "/body/p[N]", "/body/tbl[N]",
                "/body/tbl[N]/row[N]/cell[N]",
            ]);
    }

    private static T OneBased<T>(IReadOnlyList<T> items, int index, string path)
        where T : class
    {
        if (index < 1 || index > items.Count)
            throw new BatchEditException(
                "not_found", $"路径索引超出范围: {path}", path,
                suggestion: $"有效索引为 1..{items.Count}");
        return items[index - 1];
    }

    // ── set ──────────────────────────────────────────────────────

    private static string Set(
        MainDocumentPart main, string kind, OpenXmlElement element,
        OfficeOperation operation)
    {
        switch (kind)
        {
            case "paragraph":
                SetParagraph(main, (W.Paragraph)element, operation);
                break;
            case "cell":
                BatchEditJson.EnsureAllowed(operation, ["text"], "Word 单元格");
                if (BatchEditJson.TryGet(operation, "text", out var text))
                    SetCellText((W.TableCell)element,
                        BatchEditJson.PythonString(text));
                break;
            default:
                throw new BatchEditException(
                    "unsupported_operation", $"set 不支持目标类型: {kind}",
                    operation.Path);
        }
        return $"已更新 {operation.Path}";
    }

    private static void SetParagraph(
        MainDocumentPart main, W.Paragraph paragraph, OfficeOperation operation)
    {
        BatchEditJson.EnsureAllowed(operation, ParagraphProps, "Word 段落");

        if (BatchEditJson.TryGet(operation, "text", out var text))
            SetParagraphText(paragraph, BatchEditJson.PythonString(text));

        if (BatchEditJson.TryGet(operation, "style", out var style))
        {
            var styleName = BatchEditJson.PythonString(style);
            SetParagraphStyle(paragraph,
                ResolveStyleId(main, styleName, operation.Path));
        }

        if (BatchEditJson.TryGet(operation, "alignment", out var alignment))
        {
            var value = BatchEditJson.PythonString(alignment).ToLowerInvariant();
            if (!TryJustification(value, out var justification))
                throw new BatchEditException(
                    "invalid_value", $"无效的对齐方式: {value}", operation.Path,
                    validValues: ["left", "center", "right", "justify"]);
            SetJustification(paragraph, justification);
        }

        var bold = BatchEditJson.TryGet(operation, "bold", out var boldValue)
            ? BatchEditJson.Truthy(boldValue) : (bool?)null;
        var italic = BatchEditJson.TryGet(operation, "italic", out var italicValue)
            ? BatchEditJson.Truthy(italicValue) : (bool?)null;
        var fontName = BatchEditJson.OptString(operation, "font_name");
        var fontSize = BatchEditJson.TryGet(operation, "font_size", out _)
            ? BatchEditJson.RequireDouble(operation, "font_size") : (double?)null;
        var fontColor = BatchEditJson.TryGet(operation, "font_color", out var colorValue)
            ? BatchEditJson.RequireHexColor(operation, "font_color", colorValue) : null;

        if (bold is null && italic is null && fontName is null
            && fontSize is null && fontColor is null)
            return;

        foreach (var run in paragraph.Elements<W.Run>())
            ApplyRunProperties(run, bold, italic, fontName, fontSize, fontColor);
    }

    private static void SetParagraphText(W.Paragraph paragraph, string text)
    {
        var runs = paragraph.Elements<W.Run>().ToList();
        if (runs.Count > 0)
        {
            SetRunContent(runs[0], text);
            for (var index = 1; index < runs.Count; index++)
                SetRunContent(runs[index], "");
        }
        else
        {
            var run = new W.Run();
            AppendRunContent(run, text);
            paragraph.Append(run);
        }
    }

    private static void SetCellText(W.TableCell cell, string text)
    {
        foreach (var child in cell.ChildElements
                     .Where(child => child is not W.TableCellProperties).ToList())
            child.Remove();
        var paragraph = new W.Paragraph();
        var run = new W.Run();
        AppendRunContent(run, text);
        paragraph.Append(run);
        cell.Append(paragraph);
    }

    private static void SetRunContent(W.Run run, string text)
    {
        foreach (var child in run.ChildElements
                     .Where(child => child is not W.RunProperties).ToList())
            child.Remove();
        AppendRunContent(run, text);
    }

    private static void AppendRunContent(W.Run run, string text)
    {
        var buffer = new StringBuilder();
        void Flush()
        {
            if (buffer.Length == 0)
                return;
            run.Append(new W.Text(buffer.ToString())
            {
                Space = SpaceProcessingModeValues.Preserve,
            });
            buffer.Clear();
        }

        for (var index = 0; index < text.Length; index++)
        {
            var character = text[index];
            if (character == '\r' || character == '\n')
            {
                Flush();
                run.Append(new W.Break());
                if (character == '\r' && index + 1 < text.Length
                    && text[index + 1] == '\n')
                    index++;
            }
            else if (character == '\t')
            {
                Flush();
                run.Append(new W.TabChar());
            }
            else
            {
                buffer.Append(character);
            }
        }
        Flush();
    }

    private static void ApplyRunProperties(
        W.Run run, bool? bold, bool? italic, string? fontName,
        double? fontSize, string? fontColor)
    {
        var properties = run.RunProperties;
        if (properties is null)
        {
            properties = new W.RunProperties();
            run.PrependChild(properties);
        }
        if (bold is { } boldValue)
            SetChildInOrder(properties, new W.Bold { Val = boldValue }, RunOrder);
        if (italic is { } italicValue)
            SetChildInOrder(properties, new W.Italic { Val = italicValue }, RunOrder);
        if (fontName is not null)
        {
            var fonts = properties.GetFirstChild<W.RunFonts>()
                        ?? new W.RunFonts();
            fonts.Ascii = fontName;
            fonts.HighAnsi = fontName;
            if (fonts.Parent is null)
                SetChildInOrder(properties, fonts, RunOrder);
        }
        if (fontSize is { } size)
        {
            var halfPoints = ((int)Math.Round(size * 2,
                MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture);
            SetChildInOrder(properties, new W.FontSize { Val = halfPoints }, RunOrder);
            SetChildInOrder(properties,
                new W.FontSizeComplexScript { Val = halfPoints }, RunOrder);
        }
        if (fontColor is not null)
            SetChildInOrder(properties, new W.Color { Val = fontColor }, RunOrder);
    }

    private static void SetParagraphStyle(W.Paragraph paragraph, string styleId)
    {
        var properties = EnsureParagraphProperties(paragraph);
        var existing = properties.GetFirstChild<W.ParagraphStyleId>();
        if (existing is not null)
        {
            existing.Val = styleId;
            return;
        }
        properties.PrependChild(new W.ParagraphStyleId { Val = styleId });
    }

    private static void SetJustification(
        W.Paragraph paragraph, W.JustificationValues value)
    {
        var properties = EnsureParagraphProperties(paragraph);
        var existing = properties.GetFirstChild<W.Justification>();
        if (existing is not null)
        {
            existing.Val = value;
            return;
        }
        SetChildInOrder(properties, new W.Justification { Val = value },
            ParagraphOrder);
    }

    private static W.ParagraphProperties EnsureParagraphProperties(
        W.Paragraph paragraph)
    {
        var properties = paragraph.ParagraphProperties;
        if (properties is not null)
            return properties;
        properties = new W.ParagraphProperties();
        paragraph.PrependChild(properties);
        return properties;
    }

    private static bool TryJustification(
        string value, out W.JustificationValues justification)
    {
        switch (value)
        {
            case "left":
                justification = W.JustificationValues.Left;
                return true;
            case "center":
                justification = W.JustificationValues.Center;
                return true;
            case "right":
                justification = W.JustificationValues.Right;
                return true;
            case "justify":
                justification = W.JustificationValues.Both;
                return true;
            default:
                justification = W.JustificationValues.Left;
                return false;
        }
    }

    // ── add ──────────────────────────────────────────────────────

    private static string Add(
        MainDocumentPart main, W.Body body, string kind,
        OfficeOperation operation)
    {
        if (kind != "body")
            throw new BatchEditException(
                "unsupported_operation", "Word add 当前仅支持 /body",
                operation.Path);

        var elementType = BatchEditJson.RequireElementType(
            operation, "Word", ["paragraph", "heading", "page_break", "table"]);
        switch (elementType)
        {
            case "paragraph":
            {
                var paragraph = new W.Paragraph();
                if (BatchEditJson.TryGet(operation, "style", out var style))
                    SetParagraphStyle(paragraph,
                        ResolveStyleId(main, BatchEditJson.PythonString(style),
                            operation.Path));
                var text = BatchEditJson.PythonStringProp(operation, "text");
                if (text.Length > 0)
                {
                    var run = new W.Run();
                    AppendRunContent(run, text);
                    paragraph.Append(run);
                }
                WordFileOoxml.AppendBodyElement(body, paragraph);
                break;
            }
            case "heading":
            {
                var level = BatchEditJson.OptInt(operation, "level", 1);
                var styleName = level == 0 ? "Title" : $"Heading {level}";
                var paragraph = new W.Paragraph();
                SetParagraphStyle(paragraph,
                    ResolveStyleId(main, styleName, operation.Path));
                var text = BatchEditJson.PythonStringProp(operation, "text");
                if (text.Length > 0)
                {
                    var run = new W.Run();
                    AppendRunContent(run, text);
                    paragraph.Append(run);
                }
                WordFileOoxml.AppendBodyElement(body, paragraph);
                break;
            }
            case "page_break":
            {
                var paragraph = new W.Paragraph();
                var run = new W.Run(new W.Break { Type = W.BreakValues.Page });
                paragraph.Append(run);
                WordFileOoxml.AppendBodyElement(body, paragraph);
                break;
            }
            case "table":
                WordFileOoxml.AppendBodyElement(body,
                    BuildTable(main, body, operation));
                break;
            default:
                throw new BatchEditException(
                    "invalid_value",
                    $"不支持的 Word 元素类型: {elementType}", operation.Path,
                    validValues: ["paragraph", "heading", "page_break", "table"]);
        }
        return $"已添加 {elementType}";
    }

    private static W.Table BuildTable(
        MainDocumentPart main, W.Body body, OfficeOperation operation)
    {
        var headers = new List<string>();
        if (BatchEditJson.TryGet(operation, "headers", out var headersElement)
            && headersElement.ValueKind != JsonValueKind.Null)
        {
            if (headersElement.ValueKind != JsonValueKind.Array)
                throw new BatchEditException(
                    "invalid_value", "table.headers 必须是数组", operation.Path);
            headers = BatchEditJson.ArrayStrings(
                headersElement, "table.headers", operation.Path);
        }

        var rows = new List<List<string>>();
        if (BatchEditJson.TryGet(operation, "rows", out var rowsElement)
            && rowsElement.ValueKind != JsonValueKind.Null)
        {
            if (rowsElement.ValueKind != JsonValueKind.Array)
                throw new BatchEditException(
                    "invalid_value", "table.rows 必须是二维数组", operation.Path);
            foreach (var rowElement in rowsElement.EnumerateArray())
            {
                if (rowElement.ValueKind != JsonValueKind.Array)
                    throw new BatchEditException(
                        "invalid_value", "table.rows 必须是二维数组", operation.Path);
                rows.Add(BatchEditJson.ArrayStrings(
                    rowElement, "table.rows", operation.Path));
            }
        }

        var columnCount = Math.Max(headers.Count,
            rows.Count == 0 ? 0 : rows.Max(row => row.Count));
        if (columnCount < 1)
            throw new BatchEditException(
                "invalid_value", "table 至少需要一列", operation.Path);

        var styleId = ResolveStyleId(main, "Table Grid", operation.Path);
        var columnWidth = Math.Max(1, TextBlockWidth(body) / columnCount);
        var widthText = columnWidth.ToString(CultureInfo.InvariantCulture);
        var table = new W.Table();
        table.Append(new W.TableProperties(
            new W.TableStyle { Val = styleId },
            new W.TableWidth
            {
                Type = W.TableWidthUnitValues.Auto,
                Width = "0",
            }));
        var grid = new W.TableGrid();
        for (var column = 0; column < columnCount; column++)
            grid.Append(new W.GridColumn { Width = widthText });
        table.Append(grid);

        var allRows = new List<List<string>>();
        if (headers.Count > 0)
            allRows.Add(headers);
        allRows.AddRange(rows);
        foreach (var values in allRows)
        {
            var row = new W.TableRow();
            for (var column = 0; column < columnCount; column++)
            {
                var cell = new W.TableCell();
                cell.Append(new W.TableCellProperties(
                    new W.TableCellWidth
                    {
                        Type = W.TableWidthUnitValues.Dxa,
                        Width = widthText,
                    }));
                SetCellText(cell, column < values.Count ? values[column] : "");
                row.Append(cell);
            }
            table.Append(row);
        }
        return table;
    }

    private static int TextBlockWidth(W.Body body)
    {
        var section = body.Elements<W.SectionProperties>().LastOrDefault();
        var pageWidth = (int)(section?.GetFirstChild<W.PageSize>()?.Width?.Value
                              ?? 12240U);
        var margin = section?.GetFirstChild<W.PageMargin>();
        var left = (int)(margin?.Left?.Value ?? 1440U);
        var right = (int)(margin?.Right?.Value ?? 1440U);
        var width = pageWidth - left - right;
        return width > 0 ? width : 9360;
    }

    // ── remove ───────────────────────────────────────────────────

    private static string Remove(
        string kind, OpenXmlElement element, OfficeOperation operation)
    {
        switch (kind)
        {
            case "paragraph":
            case "table":
                element.Remove();
                break;
            case "cell":
                SetCellText((W.TableCell)element, "");
                break;
            default:
                throw new BatchEditException(
                    "unsupported_operation", "不能删除整个 Word 正文",
                    operation.Path);
        }
        return $"已删除 {operation.Path}";
    }

    // ── move ─────────────────────────────────────────────────────

    private static string Move(
        W.Body body, string kind, OpenXmlElement element,
        OfficeOperation operation)
    {
        if (kind is not ("paragraph" or "table")
            || string.IsNullOrEmpty(operation.TargetPath))
            throw new BatchEditException(
                "invalid_value", "Word move 需要正文段落/表格和 target_path",
                operation.Path);

        var (targetKind, moveTarget) = Resolve(body, operation.TargetPath);
        if (targetKind is not ("paragraph" or "table"))
            throw new BatchEditException(
                "invalid_path", "target_path 必须是正文段落或表格",
                operation.TargetPath);
        if (ReferenceEquals(element, moveTarget))
            throw new BatchEditException(
                "invalid_value", "源路径和目标路径不能相同", operation.Path);

        element.Remove();
        if (operation.Position == "before")
            moveTarget.InsertBeforeSelf(element);
        else
            moveTarget.InsertAfterSelf(element);
        return $"已移动 {operation.Path}";
    }

    // ── replace ──────────────────────────────────────────────────

    private static string Replace(
        string kind, OpenXmlElement element, OfficeOperation operation)
    {
        var find = BatchEditJson.PythonStringProp(operation, "find");
        var replacement = BatchEditJson.PythonStringProp(operation, "replace");
        if (find.Length == 0)
            throw new BatchEditException(
                "invalid_value", "replace 操作的 find 不能为空", operation.Path);

        IEnumerable<W.Paragraph> paragraphs;
        switch (kind)
        {
            case "body":
                paragraphs = ParagraphsInBody((W.Body)element);
                break;
            case "paragraph":
                paragraphs = new[] { (W.Paragraph)element };
                break;
            case "cell":
                paragraphs = ((W.TableCell)element).Elements<W.Paragraph>();
                break;
            default:
                throw new BatchEditException(
                    "unsupported_operation", "replace 仅支持正文、段落或单元格",
                    operation.Path);
        }

        var count = paragraphs.Sum(paragraph =>
            ReplaceParagraphText(paragraph, find, replacement));
        return $"已替换 {count} 处文本";
    }

    private static IEnumerable<W.Paragraph> ParagraphsInBody(W.Body body)
    {
        foreach (var paragraph in body.Elements<W.Paragraph>())
            yield return paragraph;
        foreach (var table in body.Elements<W.Table>())
        foreach (var row in table.Elements<W.TableRow>())
        foreach (var cell in row.Elements<W.TableCell>())
        foreach (var paragraph in cell.Elements<W.Paragraph>())
            yield return paragraph;
    }

    private static int ReplaceParagraphText(
        W.Paragraph paragraph, string find, string replacement)
    {
        var original = WordFileOoxml.ParagraphText(paragraph);
        var count = CountOccurrences(original, find);
        if (count == 0)
            return 0;
        var changed = original.Replace(find, replacement, StringComparison.Ordinal);
        var runs = paragraph.Elements<W.Run>().ToList();
        if (runs.Count > 0)
        {
            SetRunContent(runs[0], changed);
            for (var index = 1; index < runs.Count; index++)
                SetRunContent(runs[index], "");
        }
        else
        {
            var run = new W.Run();
            AppendRunContent(run, changed);
            paragraph.Append(run);
        }
        return count;
    }

    private static int CountOccurrences(string text, string find)
    {
        var count = 0;
        var offset = 0;
        while (offset <= text.Length - find.Length)
        {
            var index = text.IndexOf(find, offset, StringComparison.Ordinal);
            if (index < 0)
                break;
            count++;
            offset = index + find.Length;
        }
        return count;
    }

    // ── styles / schema ordering ─────────────────────────────────

    private static string ResolveStyleId(
        MainDocumentPart main, string styleName, string path)
    {
        var styleId = FindStyleId(main, styleName);
        if (styleId is not null)
            return styleId;
        WordFileOoxml.EnsureRequiredStyles(main);
        styleId = FindStyleId(main, styleName);
        if (styleId is not null)
            return styleId;
        throw new BatchEditException(
            "not_found", $"Word 样式不存在: {styleName}", path);
    }

    private static string? FindStyleId(MainDocumentPart main, string styleName)
    {
        var styles = main.StyleDefinitionsPart?.Styles;
        if (styles is null)
            return null;
        var byName = styles.Elements<W.Style>().FirstOrDefault(style =>
            string.Equals(style.StyleName?.Val?.Value, styleName,
                StringComparison.OrdinalIgnoreCase));
        if (byName?.StyleId?.Value is { } nameId)
            return nameId;
        var byId = styles.Elements<W.Style>().FirstOrDefault(style =>
            string.Equals(style.StyleId?.Value, styleName,
                StringComparison.OrdinalIgnoreCase));
        if (byId?.StyleId?.Value is { } id)
            return id;
        var compact = RemoveWhitespace(styleName);
        var byCompactId = styles.Elements<W.Style>().FirstOrDefault(style =>
            RemoveWhitespace(style.StyleId?.Value ?? "") is { Length: > 0 } candidate
            && string.Equals(candidate, compact, StringComparison.OrdinalIgnoreCase));
        return byCompactId?.StyleId?.Value ?? BuiltinStyleId(styleName);
    }

    private static string? BuiltinStyleId(string name)
    {
        var compact = RemoveWhitespace(name);
        if (compact.Equals("Normal", StringComparison.OrdinalIgnoreCase))
            return "Normal";
        if (compact.Equals("Title", StringComparison.OrdinalIgnoreCase))
            return "Title";
        if (compact.Equals("Subtitle", StringComparison.OrdinalIgnoreCase))
            return "Subtitle";
        if (compact.Equals("TableGrid", StringComparison.OrdinalIgnoreCase))
            return "TableGrid";
        if (compact.Equals("ListParagraph", StringComparison.OrdinalIgnoreCase))
            return "ListParagraph";
        if (compact.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)
            && int.TryParse(compact["Heading".Length..], out var level)
            && level is >= 1 and <= 9)
            return $"Heading{level}";
        return null;
    }

    private static string RemoveWhitespace(string value) =>
        new(value.Where(character => !char.IsWhiteSpace(character)).ToArray());

    private static void SetChildInOrder(
        OpenXmlElement parent, OpenXmlElement child, string[] order)
    {
        var existing = parent.ChildElements.FirstOrDefault(item =>
            item.LocalName == child.LocalName);
        if (existing is not null)
        {
            parent.ReplaceChild(child, existing);
            return;
        }
        var childOrder = IndexInOrder(child.LocalName, order);
        foreach (var item in parent.ChildElements)
        {
            if (IndexInOrder(item.LocalName, order) > childOrder)
            {
                item.InsertBeforeSelf(child);
                return;
            }
        }
        parent.Append(child);
    }

    private static int IndexInOrder(string localName, string[] order)
    {
        var index = Array.IndexOf(order, localName);
        return index < 0 ? int.MaxValue : index;
    }
}
