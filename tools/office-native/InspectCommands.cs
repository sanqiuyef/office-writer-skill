using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeNative;

internal static class InspectCommands
{
    private sealed class NodeBudget
    {
        public NodeBudget(int maximum) => Maximum = maximum;
        public int Maximum { get; }
        public int Count { get; private set; } = 1;
        public bool Truncated { get; private set; }

        public bool Add(Dictionary<string, object?> parent, Dictionary<string, object?> child)
        {
            if (Count >= Maximum)
            {
                Truncated = true;
                return false;
            }
            ((List<Dictionary<string, object?>>)parent["children"]!).Add(child);
            Count++;
            return true;
        }
    }

    public static Dictionary<string, object?> RunInspect(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        var maxNodes = NativeCommon.OptInt(p, "max_nodes", 1000);
        if (maxNodes is < 1 or > 10000)
            throw new NativeUsageException("max_nodes 必须在 1 到 10000 之间");
        NativeCommon.EnsureFile(filePath);
        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        var budget = new NodeBudget(maxNodes);
        var (format, root) = extension switch
        {
            ".docx" => ("docx", InspectWord(filePath, budget)),
            ".xlsx" or ".xlsm" => ("xlsx", InspectExcel(filePath, budget)),
            ".pptx" => ("pptx", InspectPowerpoint(filePath, budget)),
            _ => throw new NativeRuntimeException(
                $"不支持的文件格式: {extension}（支持 .docx/.xlsx/.xlsm/.pptx）"),
        };
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["file_path"] = filePath,
            ["format"] = format,
            ["node_count"] = budget.Count,
            ["truncated"] = budget.Truncated,
            ["root"] = root,
        };
    }

    private static Dictionary<string, object?> Node(
        string path, string type, string? text = null, Dictionary<string, object?>? format = null,
        string? name = null, object? value = null)
    {
        var node = new Dictionary<string, object?>
        {
            ["path"] = path,
            ["type"] = type,
            ["children"] = new List<Dictionary<string, object?>>(),
        };
        if (text is not null)
            node["text"] = text;
        if (format is not null)
            node["format"] = format;
        if (name is not null)
            node["name"] = name;
        if (value is not null)
            node["value"] = value;
        return node;
    }

    // ── Word ─────────────────────────────────────────────────────

    private static Dictionary<string, object?> InspectWord(string filePath, NodeBudget budget)
    {
        using var document = WordprocessingDocument.Open(filePath, false);
        var body = document.MainDocumentPart?.Document.Body
                   ?? throw new NativeRuntimeException("文档没有正文");
        var styleNames = document.MainDocumentPart?.StyleDefinitionsPart?.Styles?
            .Elements<W.Style>()
            .Where(style => style.StyleId?.Value is not null)
            .ToDictionary(
                style => style.StyleId!.Value!,
                style => style.StyleName?.Val?.Value ?? style.StyleId!.Value!,
                StringComparer.Ordinal)
            ?? new Dictionary<string, string>(StringComparer.Ordinal);

        var root = Node("/", "document");
        root["format"] = "docx";
        var bodyNode = Node("/body", "body");
        if (!budget.Add(root, bodyNode))
            return root;

        var paragraphIndex = 0;
        foreach (var paragraph in body.Elements<W.Paragraph>())
        {
            paragraphIndex++;
            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            var styleName = styleId is null
                ? "Normal"
                : styleNames.TryGetValue(styleId, out var display) ? display : styleId;
            if (styleName.StartsWith("heading ", StringComparison.OrdinalIgnoreCase))
                styleName = "Heading " + styleName[8..];
            var justification = paragraph.ParagraphProperties?.Justification?.Val;
            string? alignment = justification is null
                ? null
                : justification == W.JustificationValues.Left ? "LEFT (0)"
                : justification == W.JustificationValues.Center ? "CENTER (1)"
                : justification == W.JustificationValues.Right ? "RIGHT (2)"
                : justification == W.JustificationValues.Both ? "JUSTIFY (3)"
                : justification.ToString();
            var node = Node(
                $"/body/p[{paragraphIndex}]", "paragraph",
                WordParagraphText(paragraph),
                new Dictionary<string, object?>
                {
                    ["alignment"] = alignment,
                });
            node["style"] = styleName;
            if (!budget.Add(bodyNode, node))
                break;
        }

        var tableIndex = 0;
        foreach (var table in body.Elements<W.Table>())
        {
            tableIndex++;
            var rows = table.Elements<W.TableRow>().ToList();
            var columns = rows.Count == 0
                ? 0
                : rows.Max(row => row.Elements<W.TableCell>()
                    .Sum(WordTableGrid.GridSpan));
            var tableNode = Node(
                $"/body/tbl[{tableIndex}]", "table",
                format: new Dictionary<string, object?>
                {
                    ["rows"] = rows.Count,
                    ["columns"] = columns,
                });
            if (!budget.Add(bodyNode, tableNode))
                break;
            for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                var rowNode = Node(
                    $"/body/tbl[{tableIndex}]/row[{rowIndex + 1}]", "row");
                if (!budget.Add(tableNode, rowNode))
                    break;
                var cells = rows[rowIndex].Elements<W.TableCell>().ToList();
                var gridColumn = 0;
                foreach (var cell in cells)
                {
                    // python-docx 按网格列数展开合并单元格，vMerge=continue 的文本
                    // 取自顶层 restart 单元格；路径必须与它对齐。
                    var effective = WordTableGrid.RootCell(rows, rowIndex, cell);
                    for (var span = 0; span < WordTableGrid.GridSpan(effective); span++)
                    {
                        gridColumn++;
                        var text = string.Join("\n", effective.Elements<W.Paragraph>()
                            .Select(WordParagraphText));
                        var cellNode = Node(
                            $"/body/tbl[{tableIndex}]/row[{rowIndex + 1}]/cell[{gridColumn}]",
                            "cell", text);
                        if (!budget.Add(rowNode, cellNode))
                            break;
                    }
                }
            }
        }
        return root;
    }

    private static string WordParagraphText(W.Paragraph paragraph)
    {
        var builder = new System.Text.StringBuilder();
        foreach (var element in paragraph.Descendants())
        {
            switch (element)
            {
                case W.Text text:
                    builder.Append(text.Text);
                    break;
                case W.TabChar:
                    builder.Append('\t');
                    break;
                case W.Break or W.CarriageReturn:
                    builder.Append('\n');
                    break;
            }
        }
        return builder.ToString();
    }

    // ── Excel ────────────────────────────────────────────────────

    private static Dictionary<string, object?> InspectExcel(string filePath, NodeBudget budget)
    {
        using var document = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = document.WorkbookPart
                           ?? throw new NativeRuntimeException("工作簿没有 WorkbookPart");
        var shared = workbookPart.SharedStringTablePart?.SharedStringTable?
            .Elements<S.SharedStringItem>()
            .Select(item => item.InnerText)
            .ToList() ?? [];
        var stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet;
        var cellFormats = stylesheet?.CellFormats?.Elements<S.CellFormat>().ToList() ?? [];
        var fonts = stylesheet?.Fonts?.Elements<S.Font>().ToList() ?? [];
        var fills = stylesheet?.Fills?.Elements<S.Fill>().ToList() ?? [];
        var numberFormats = stylesheet?.NumberingFormats?.Elements<S.NumberingFormat>()
            .Where(item => item.NumberFormatId?.Value is not null)
            .ToDictionary(item => item.NumberFormatId!.Value, item => item.FormatCode?.Value)
            ?? new Dictionary<uint, string?>();

        var root = Node("/", "workbook");
        root["format"] = "xlsx";
        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            var worksheet = worksheetPart.Worksheet
                            ?? throw new NativeRuntimeException("工作表为空");
            var sheetElement = workbookPart.Workbook.Descendants<S.Sheet>()
                .FirstOrDefault(sheet => sheet.Id?.Value ==
                    workbookPart.GetIdOfPart(worksheetPart));
            var title = sheetElement?.Name?.Value ?? "Sheet";
            var sheetData = worksheet.GetFirstChild<S.SheetData>();
            var cells = sheetData?.Elements<S.Row>()
                .SelectMany(row => row.Elements<S.Cell>())
                .Where(cell => cell.CellReference?.Value is not null)
                .ToList() ?? [];
            var maxRow = cells.Count == 0 ? 1 : cells.Max(cell => RowIndex(cell.CellReference!.Value!));
            var maxColumn = cells.Count == 0 ? 1 : cells.Max(cell => ColumnIndex(cell.CellReference!.Value!));
            var stateValue = sheetElement?.State?.Value;
            var state = stateValue == S.SheetStateValues.Hidden ? "hidden"
                : stateValue == S.SheetStateValues.VeryHidden ? "veryHidden"
                : "visible";
            var sheetNode = Node(
                $"/{title}", "sheet",
                format: new Dictionary<string, object?>
                {
                    ["max_row"] = maxRow,
                    ["max_column"] = maxColumn,
                    ["state"] = state,
                });
            if (!budget.Add(root, sheetNode))
                break;
            foreach (var cell in cells)
            {
                if (cell.CellValue is null && cell.CellFormula is null && cell.InlineString is null)
                    continue;
                var styleIndex = cell.StyleIndex?.Value ?? 0;
                var format = ExcelCellFormat(styleIndex, cellFormats, fonts, fills, numberFormats);
                var (text, value) = ExcelCellText(cell, shared, format);
                if (text is null)
                    continue;
                var node = Node(
                    $"/{title}/{cell.CellReference!.Value}", "cell", text,
                    format: format, value: value);
                if (!budget.Add(sheetNode, node))
                    break;
            }
        }
        return root;
    }

    private static (string? Text, object? Value) ExcelCellText(
        S.Cell cell, List<string> shared, Dictionary<string, object?> format)
    {
        if (cell.CellFormula?.Text is { Length: > 0 } formula)
            return ("=" + formula, "=" + formula);
        if (cell.InlineString is not null)
        {
            var inline = string.Concat(
                cell.InlineString.Descendants<S.Text>().Select(text => text.Text));
            return (inline, inline);
        }
        var raw = cell.CellValue?.Text;
        if (raw is null)
            return (null, null);
        if (cell.DataType?.Value == S.CellValues.SharedString
            && int.TryParse(raw, out var index) && index >= 0 && index < shared.Count)
            return (shared[index], shared[index]);
        if (cell.DataType?.Value == S.CellValues.Boolean)
        {
            var boolean = raw == "1";
            return (boolean ? "True" : "False", boolean);
        }
        if (double.TryParse(raw, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var number))
        {
            var numberFormat = format["number_format"]?.ToString() ?? string.Empty;
            if (IsDateFormat(numberFormat))
            {
                try
                {
                    var text = DateTime.FromOADate(number).ToString("yyyy-MM-dd HH:mm:ss");
                    return (text, text);
                }
                catch (ArgumentException) { }
            }
            return (raw, number);
        }
        return (raw, raw);
    }

    private static bool IsDateFormat(string numberFormat)
    {
        if (string.IsNullOrWhiteSpace(numberFormat)
            || numberFormat is "General" or "@")
            return false;
        return numberFormat.IndexOfAny(['y', 'Y', 'm', 'M', 'd', 'D', 'h', 'H', 's', 'S']) >= 0;
    }

    private static Dictionary<string, object?> ExcelCellFormat(
        uint styleIndex, List<S.CellFormat> cellFormats, List<S.Font> fonts,
        List<S.Fill> fills, Dictionary<uint, string?> numberFormats)
    {
        if (styleIndex >= cellFormats.Count)
            return new Dictionary<string, object?>
            {
                ["number_format"] = "General",
                ["bold"] = false,
                ["fill"] = null,
            };
        var cellFormat = cellFormats[(int)styleIndex];
        var numFmtId = cellFormat.NumberFormatId?.Value ?? 0;
        var numberFormat = numberFormats.TryGetValue(numFmtId, out var custom) && custom is not null
            ? custom
            : BuiltinNumberFormat(numFmtId);
        var fontIndex = cellFormat.FontId?.Value ?? 0;
        var bold = fontIndex < fonts.Count && fonts[(int)fontIndex].Bold is not null
            && (fonts[(int)fontIndex].Bold!.Val?.Value ?? true);
        var fillIndex = cellFormat.FillId?.Value ?? 0;
        string? fill = null;
        if (fillIndex < fills.Count)
        {
            var pattern = fills[(int)fillIndex].PatternFill;
            fill = pattern?.ForegroundColor?.Rgb?.Value
                   ?? pattern?.ForegroundColor?.Theme?.Value.ToString();
        }
        return new Dictionary<string, object?>
        {
            ["number_format"] = numberFormat,
            ["bold"] = bold,
            ["fill"] = fill,
        };
    }

    private static string BuiltinNumberFormat(uint id) => id switch
    {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        9 => "0%",
        10 => "0.00%",
        14 => "m/d/yyyy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yyyy h:mm",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mmss.0",
        49 => "@",
        _ => id.ToString(),
    };

    private static int RowIndex(string reference)
    {
        var digits = new string(reference.SkipWhile(char.IsLetter).ToArray());
        return int.TryParse(digits, out var row) ? row : 0;
    }

    private static int ColumnIndex(string reference)
    {
        var column = 0;
        foreach (var character in reference.TakeWhile(char.IsLetter))
            column = column * 26 + (char.ToUpperInvariant(character) - 'A' + 1);
        return column;
    }

    // ── PowerPoint ───────────────────────────────────────────────

    private static Dictionary<string, object?> InspectPowerpoint(string filePath, NodeBudget budget)
    {
        using var document = PresentationDocument.Open(filePath, false);
        var presentationPart = document.PresentationPart
                               ?? throw new NativeRuntimeException("演示文稿没有 PresentationPart");
        var slideIds = presentationPart.Presentation.SlideIdList?.Elements<P.SlideId>()
            .ToList() ?? [];
        var root = Node("/", "presentation");
        root["format"] = "pptx";
        for (var slideIndex = 0; slideIndex < slideIds.Count; slideIndex++)
        {
            var relationshipId = slideIds[slideIndex].RelationshipId?.Value;
            if (relationshipId is null
                || presentationPart.GetPartById(relationshipId) is not SlidePart slidePart)
                continue;
            var slide = slidePart.Slide;
            var slideNode = Node(
                $"/slide[{slideIndex + 1}]", "slide",
                format: new Dictionary<string, object?>
                {
                    ["layout"] = slidePart.SlideLayoutPart?.SlideLayout.CommonSlideData?.Name?.Value,
                });
            if (!budget.Add(root, slideNode))
                break;
            var shapeTree = slide.CommonSlideData?.ShapeTree;
            if (shapeTree is null)
                continue;
            var shapeIndex = 0;
            foreach (var shape in shapeTree.ChildElements)
            {
                if (shape is not (P.Shape or P.Picture or P.GraphicFrame or P.GroupShape or P.ConnectionShape)
                    && !shape.LocalName.Equals("contentPart", StringComparison.Ordinal))
                    continue;
                shapeIndex++;
                var node = Node(
                    $"/slide[{slideIndex + 1}]/shape[{shapeIndex}]", "shape",
                    ShapeText(shape), ShapeFormat(shape), ShapeName(shape));
                if (!budget.Add(slideNode, node))
                    break;
            }
        }
        return root;
    }

    private static string? ShapeName(DocumentFormat.OpenXml.OpenXmlElement shape) =>
        shape.Descendants<P.NonVisualDrawingProperties>().FirstOrDefault()?.Name?.Value;

    private static string? ShapeText(DocumentFormat.OpenXml.OpenXmlElement shape)
    {
        if (shape is P.Shape textShape)
            return string.Join("\n", textShape.TextBody?.Elements<A.Paragraph>()
                .Select(paragraph => string.Concat(
                    paragraph.Descendants<A.Text>().Select(text => text.Text))) ?? []);
        if (shape is P.GraphicFrame frame)
        {
            var table = frame.Descendants<A.Table>().FirstOrDefault();
            if (table is not null)
                return string.Join("\n", table.Elements<A.TableRow>().Select(row =>
                    string.Join(" | ", row.Elements<A.TableCell>().Select(cell =>
                        string.Concat(cell.Descendants<A.Text>().Select(text => text.Text))))));
        }
        return null;
    }

    private static string PptShapeType(DocumentFormat.OpenXml.OpenXmlElement shape) => shape switch
    {
        P.Picture => "PICTURE (13)",
        P.ConnectionShape => "LINE (9)",
        P.GroupShape => "GROUP (6)",
        P.GraphicFrame frame when frame.Descendants<A.Table>().Any() => "TABLE (19)",
        P.GraphicFrame frame when frame.Descendants<A.GraphicData>()
            .Any(data => (data.Uri?.Value ?? string.Empty).Contains("chart")) => "CHART (3)",
        P.Shape item when item.NonVisualShapeProperties?
            .NonVisualShapeDrawingProperties?.TextBox?.Value == true => "TEXT_BOX (17)",
        P.Shape item when item.NonVisualShapeProperties?
            .ApplicationNonVisualDrawingProperties?.PlaceholderShape is not null => "PLACEHOLDER (14)",
        P.Shape => "AUTO_SHAPE (1)",
        _ => shape.LocalName,
    };

    private static Dictionary<string, object?> ShapeFormat(DocumentFormat.OpenXml.OpenXmlElement shape)
    {
        var offset = shape.Descendants<A.Offset>().FirstOrDefault();
        var extents = shape.Descendants<A.Extents>().FirstOrDefault();
        static double Emu(long? value) => Math.Round((value ?? 0) / 914400.0, 3);
        return new Dictionary<string, object?>
        {
            ["shape_type"] = PptShapeType(shape),
            ["x"] = Emu(offset?.X?.Value),
            ["y"] = Emu(offset?.Y?.Value),
            ["width"] = Emu(extents?.Cx?.Value),
            ["height"] = Emu(extents?.Cy?.Value),
        };
    }
}
