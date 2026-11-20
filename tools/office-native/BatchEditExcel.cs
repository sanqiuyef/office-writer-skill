using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeNative;

/// <summary>Excel 批量编辑：/{sheet}/{cell}、区域、row[N]/col[A] 与工作表排序。</summary>
internal static class BatchEditExcel
{
    private static readonly Regex PathPattern = new(
        @"^/([^/]+)(?:/([A-Za-z]+\d+(?::[A-Za-z]+\d+)?|row\[\d+\]|col\[[A-Za-z]+\]))?$");
    private static readonly Regex RangePattern = new(
        @"^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$");
    private static readonly Regex RowPattern = new(@"^row\[(\d+)\]$");
    private static readonly Regex ColumnPattern = new(@"^col\[([A-Za-z]+)\]$");

    private static readonly HashSet<string> CellProps = new(StringComparer.Ordinal)
    {
        "value", "formula", "bold", "italic", "font_name", "font_size",
        "font_color", "fill", "alignment", "number_format", "values",
    };

    private static readonly HashSet<string> FontProps = new(StringComparer.Ordinal)
    {
        "bold", "italic", "font_name", "font_size", "font_color",
    };

    private static readonly string[] FontOrder =
    [
        "b", "i", "strike", "condense", "extend", "outline", "shadow",
        "u", "vertAlign", "sz", "color", "name", "family", "charset",
        "scheme",
    ];

    private static readonly Dictionary<string, uint> BuiltinNumberFormats =
        new(StringComparer.Ordinal)
        {
            ["General"] = 0,
            ["0"] = 1,
            ["0.00"] = 2,
            ["#,##0"] = 3,
            ["#,##0.00"] = 4,
            ["0%"] = 9,
            ["0.00%"] = 10,
            ["0.00E+00"] = 11,
            ["# ?/?"] = 12,
            ["# ??/??"] = 13,
            ["mm-dd-yy"] = 14,
            ["d-mmm-yy"] = 15,
            ["d-mmm"] = 16,
            ["mmm-yy"] = 17,
            ["h:mm AM/PM"] = 18,
            ["h:mm:ss AM/PM"] = 19,
            ["h:mm"] = 20,
            ["h:mm:ss"] = 21,
            ["m/d/yy h:mm"] = 22,
            ["#,##0 ;(#,##0)"] = 37,
            ["#,##0 ;[Red](#,##0)"] = 38,
            ["#,##0.00;(#,##0.00)"] = 39,
            ["#,##0.00;[Red](#,##0.00)"] = 40,
            ["mm:ss"] = 45,
            ["[h]:mm:ss"] = 46,
            ["mmss.0"] = 47,
            ["##0.0E+0"] = 48,
            ["@"] = 49,
        };

    private static readonly string[] SheetDataFollowers =
    [
        "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios",
        "autoFilter", "sortState", "dataConsolidate", "customSheetViews",
        "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations",
        "hyperlinks", "printOptions", "pageMargins", "pageSetup",
        "headerFooter", "rowBreaks", "colBreaks", "customProperties",
        "cellWatches", "ignoredErrors", "smartTags", "drawing",
        "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects",
        "controls", "webPublishItems", "tableParts", "extLst",
    ];

    private sealed class ExcelContext
    {
        public WorkbookPart WorkbookPart { get; init; } = null!;
        public S.Workbook Workbook { get; init; } = null!;
        public S.Sheets Sheets { get; init; } = null!;
        public S.Stylesheet Stylesheet { get; init; } = null!;
        public ExcelStyleBook StyleBook { get; init; } = null!;
        public List<S.SharedStringItem> SharedStrings { get; init; } = [];
    }

    public static List<string> Apply(
        string workPath, IReadOnlyList<OfficeOperation> operations)
    {
        using var document = SpreadsheetDocument.Open(workPath, true);
        var context = CreateContext(document);
        var messages = new List<string>(operations.Count);
        foreach (var operation in operations)
        {
            try
            {
                messages.Add(ApplyOperation(context, operation));
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

        context.StyleBook.RefreshCounts();
        context.Stylesheet.Save();
        foreach (var part in context.WorkbookPart.WorksheetParts)
            part.Worksheet?.Save();
        context.Workbook.Save();
        return messages;
    }

    private static ExcelContext CreateContext(SpreadsheetDocument document)
    {
        var workbookPart = document.WorkbookPart
                           ?? throw new NativeRuntimeException("工作簿没有 WorkbookPart");
        var workbook = workbookPart.Workbook
                       ?? throw new NativeRuntimeException("工作簿缺少 Workbook 根节点");
        var sheets = workbook.Sheets;
        if (sheets is null)
        {
            sheets = new S.Sheets();
            var anchor = workbook.ChildElements.FirstOrDefault(element =>
                element.LocalName is "functionGroups" or "externalReferences"
                    or "definedNames" or "calcPr" or "oleSize"
                    or "customWorkbookViews" or "pivotCaches"
                    or "webPublishing" or "extLst");
            if (anchor is not null)
                workbook.InsertBefore(sheets, anchor);
            else
                workbook.Append(sheets);
        }
        var stylesPart = workbookPart.WorkbookStylesPart;
        if (stylesPart is null)
        {
            stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = ExcelStyleBook.CreateDefaultStylesheet();
        }
        stylesPart.Stylesheet ??= ExcelStyleBook.CreateDefaultStylesheet();
        var styleBook = new ExcelStyleBook(stylesPart.Stylesheet);
        var shared = workbookPart.SharedStringTablePart?.SharedStringTable
                         ?.Elements<S.SharedStringItem>().ToList() ?? [];
        return new ExcelContext
        {
            WorkbookPart = workbookPart,
            Workbook = workbook,
            Sheets = sheets,
            Stylesheet = stylesPart.Stylesheet,
            StyleBook = styleBook,
            SharedStrings = shared,
        };
    }

    private static string ApplyOperation(
        ExcelContext context, OfficeOperation operation)
    {
        var (worksheet, selector, sheet, part) =
            Resolve(context, operation.Path);
        return operation.Op switch
        {
            "set" => Set(context, worksheet, selector, operation),
            "add" => Add(context, worksheet, selector, operation),
            "remove" => Remove(context, worksheet, selector, sheet, part, operation),
            "move" => Move(context, worksheet, selector, sheet, operation),
            "replace" => Replace(context, worksheet, selector, operation),
            _ => throw new BatchEditException(
                "unsupported_operation", $"不支持的操作: {operation.Op}",
                operation.Path),
        };
    }

    // ── 路径/工作表解析 ───────────────────────────────────────────

    private static (S.Worksheet? Worksheet, string? Selector,
        S.Sheet? Sheet, WorksheetPart? Part) Resolve(
        ExcelContext context, string path)
    {
        if (path == "/")
            return (null, null, null, null);
        var match = PathPattern.Match(path);
        if (!match.Success)
            throw new BatchEditException(
                "invalid_path", $"不支持的 Excel 路径: {path}", path,
                validValues:
                [
                    "/", "/SheetName", "/SheetName/A1", "/SheetName/A1:C3",
                    "/SheetName/row[N]", "/SheetName/col[A]",
                ]);
        var name = match.Groups[1].Value;
        var selector = match.Groups[2].Success ? match.Groups[2].Value : null;
        var (sheet, part) = FindSheet(context, name, path);
        var worksheet = part.Worksheet
                        ?? throw new NativeRuntimeException($"工作表为空: {name}");
        return (worksheet, selector, sheet, part);
    }

    private static (S.Sheet Sheet, WorksheetPart Part) FindSheet(
        ExcelContext context, string name, string path)
    {
        var sheet = context.Sheets.Elements<S.Sheet>()
            .FirstOrDefault(item => item.Name?.Value == name);
        if (sheet is null)
            throw new BatchEditException(
                "not_found", $"工作表不存在: {name}", path,
                validValues: SheetNames(context));
        var relationshipId = sheet.Id?.Value;
        if (relationshipId is null)
            throw new BatchEditException(
                "not_found", $"工作表缺少关系: {name}", path);
        try
        {
            if (context.WorkbookPart.GetPartById(relationshipId)
                is WorksheetPart worksheetPart)
                return (sheet, worksheetPart);
        }
        catch (KeyNotFoundException)
        {
            // fall through
        }
        throw new BatchEditException(
            "unsupported_operation", $"工作表不是普通工作表: {name}", path);
    }

    private static List<string> SheetNames(ExcelContext context) =>
        context.Sheets.Elements<S.Sheet>()
            .Select(sheet => sheet.Name?.Value ?? "")
            .ToList();

    private static (int MinColumn, int MinRow, int MaxColumn, int MaxRow)
        ParseRange(string selector, string path)
    {
        var match = RangePattern.Match(selector);
        if (!match.Success)
            throw new BatchEditException(
                "invalid_path", $"无效的单元格路径: {selector}", path);
        var column1 = ColumnNumber(match.Groups[1].Value, path);
        var row1 = RowNumber(match.Groups[2].Value, path);
        var column2 = match.Groups[3].Success
            ? ColumnNumber(match.Groups[3].Value, path) : column1;
        var row2 = match.Groups[4].Success
            ? RowNumber(match.Groups[4].Value, path) : row1;
        return (Math.Min(column1, column2), Math.Min(row1, row2),
            Math.Max(column1, column2), Math.Max(row1, row2));
    }

    private static int ColumnNumber(string letters, string path)
    {
        var column = 0;
        foreach (var character in letters)
        {
            column = column * 26 + (char.ToUpperInvariant(character) - 'A' + 1);
            if (column > 16384)
                throw new BatchEditException(
                    "invalid_path", $"列号超出范围: {letters}", path);
        }
        if (column < 1)
            throw new BatchEditException(
                "invalid_path", $"无效的列号: {letters}", path);
        return column;
    }

    private static int RowNumber(string digits, string path)
    {
        if (!int.TryParse(digits, NumberStyles.Integer,
                CultureInfo.InvariantCulture, out var row)
            || row < 1 || row > 1048576)
            throw new BatchEditException(
                "invalid_path", $"行号超出范围: {digits}", path);
        return row;
    }

    private static string ColumnName(int column)
    {
        var builder = new StringBuilder();
        while (column > 0)
        {
            var remainder = (column - 1) % 26;
            builder.Insert(0, (char)('A' + remainder));
            column = (column - 1) / 26;
        }
        return builder.ToString();
    }

    private static int ColumnOfReference(string? reference)
    {
        if (string.IsNullOrEmpty(reference))
            return 0;
        var column = 0;
        foreach (var character in reference)
        {
            if (!char.IsLetter(character))
                break;
            column = column * 26 + (char.ToUpperInvariant(character) - 'A' + 1);
        }
        return column;
    }

    private static int RowOfReference(string? reference)
    {
        if (string.IsNullOrEmpty(reference))
            return 0;
        var digits = new string(reference.SkipWhile(char.IsLetter).ToArray());
        return int.TryParse(digits, NumberStyles.Integer,
            CultureInfo.InvariantCulture, out var row) ? row : 0;
    }

    private static string CellReference(int column, int row) =>
        $"{ColumnName(column)}{row}";

    // ── 单元格/区域 ──────────────────────────────────────────────

    private static S.SheetData EnsureSheetData(S.Worksheet worksheet)
    {
        var sheetData = worksheet.GetFirstChild<S.SheetData>();
        if (sheetData is not null)
            return sheetData;
        sheetData = new S.SheetData();
        var anchor = worksheet.ChildElements.FirstOrDefault(element =>
            SheetDataFollowers.Contains(element.LocalName));
        if (anchor is not null)
            worksheet.InsertBefore(sheetData, anchor);
        else
            worksheet.Append(sheetData);
        return sheetData;
    }

    private static S.Cell? GetCell(
        S.Worksheet worksheet, int row, int column, bool create)
    {
        var sheetData = EnsureSheetData(worksheet);
        var rowElement = sheetData.Elements<S.Row>()
            .FirstOrDefault(item => item.RowIndex?.Value == (uint)row);
        if (rowElement is null)
        {
            if (!create)
                return null;
            rowElement = new S.Row { RowIndex = (uint)row };
            InsertRow(sheetData, rowElement, row);
        }
        var reference = CellReference(column, row);
        var cell = rowElement.Elements<S.Cell>().FirstOrDefault(item =>
            string.Equals(item.CellReference?.Value, reference,
                StringComparison.OrdinalIgnoreCase));
        if (cell is null)
        {
            if (!create)
                return null;
            cell = new S.Cell { CellReference = reference };
            InsertCell(rowElement, cell, column);
        }
        return cell;
    }

    private static void InsertRow(S.SheetData sheetData, S.Row row, int index)
    {
        var anchor = sheetData.Elements<S.Row>()
            .FirstOrDefault(item => (item.RowIndex?.Value ?? 0) > (uint)index);
        if (anchor is not null)
            anchor.InsertBeforeSelf(row);
        else
            sheetData.Append(row);
    }

    private static void InsertCell(S.Row row, S.Cell cell, int column)
    {
        var anchor = row.Elements<S.Cell>()
            .FirstOrDefault(item =>
                ColumnOfReference(item.CellReference?.Value) > column);
        if (anchor is not null)
            anchor.InsertBeforeSelf(cell);
        else
            row.Append(cell);
    }

    private static List<S.Cell> AllCells(S.Worksheet worksheet)
    {
        var sheetData = worksheet.GetFirstChild<S.SheetData>();
        if (sheetData is null)
            return [];
        return sheetData.Elements<S.Row>()
            .SelectMany(row => row.Elements<S.Cell>())
            .Where(cell => !string.IsNullOrEmpty(cell.CellReference?.Value))
            .ToList();
    }

    private static List<S.Cell> CellsForSelector(
        S.Worksheet worksheet, string? selector, string path)
    {
        if (selector is null)
            return AllCells(worksheet);
        if (selector.StartsWith("row[", StringComparison.Ordinal)
            || selector.StartsWith("col[", StringComparison.Ordinal))
            throw new BatchEditException(
                "invalid_path", "此操作需要单元格或区域路径", path);
        var (minColumn, minRow, maxColumn, maxRow) = ParseRange(selector, path);
        return GetRangeCells(worksheet, minColumn, minRow, maxColumn, maxRow,
            out _);
    }

    /// <summary>
    /// 一次性建立行/单元格索引后批量取区域单元格，避免每个单元格线性扫描全部
    /// row。索引只在单次区域操作内复用，不跨操作缓存（操作可能插删行）。
    /// </summary>
    private static List<S.Cell> GetRangeCells(
        S.Worksheet worksheet, int minColumn, int minRow, int maxColumn, int maxRow,
        out SheetCellIndex index)
    {
        index = new SheetCellIndex(worksheet);
        var cells = new List<S.Cell>(
            checked((maxRow - minRow + 1) * (maxColumn - minColumn + 1)));
        for (var row = minRow; row <= maxRow; row++)
        for (var column = minColumn; column <= maxColumn; column++)
            cells.Add(index.GetOrCreate(row, column));
        return cells;
    }

    private sealed class SheetCellIndex
    {
        private readonly S.SheetData _sheetData;
        private readonly Dictionary<uint, S.Row> _rows = [];
        private readonly Dictionary<uint, Dictionary<int, S.Cell>> _cellsByRow = [];

        public SheetCellIndex(S.Worksheet worksheet)
        {
            _sheetData = EnsureSheetData(worksheet);
            foreach (var row in _sheetData.Elements<S.Row>())
            {
                if (row.RowIndex?.Value is { } rowIndex)
                    _rows.TryAdd(rowIndex, row);
            }
        }

        public S.Cell GetOrCreate(int rowIndex, int column)
        {
            var key = (uint)rowIndex;
            if (!_rows.TryGetValue(key, out var row))
            {
                row = new S.Row { RowIndex = key };
                InsertRow(_sheetData, row, rowIndex);
                _rows[key] = row;
            }
            if (!_cellsByRow.TryGetValue(key, out var cells))
            {
                cells = [];
                foreach (var cell in row.Elements<S.Cell>())
                {
                    var cellColumn = ColumnOfReference(cell.CellReference?.Value);
                    if (cellColumn > 0)
                        cells.TryAdd(cellColumn, cell);
                }
                _cellsByRow[key] = cells;
            }
            if (cells.TryGetValue(column, out var existing))
                return existing;
            var created = new S.Cell
            {
                CellReference = CellReference(column, rowIndex),
            };
            InsertCell(row, created, column);
            cells[column] = created;
            return created;
        }
    }

    private readonly record struct MergeRange(
        int MinColumn, int MinRow, int MaxColumn, int MaxRow);

    /// <summary>读取工作表 mergeCells；引用异常时跳过该条，不阻断其余操作。</summary>
    private static List<MergeRange> MergeRanges(S.Worksheet worksheet)
    {
        var container = worksheet.GetFirstChild<S.MergeCells>();
        if (container is null)
            return [];
        var ranges = new List<MergeRange>();
        foreach (var mergeCell in container.Elements<S.MergeCell>())
        {
            var reference = mergeCell.Reference?.Value;
            if (string.IsNullOrEmpty(reference))
                continue;
            var match = RangePattern.Match(reference);
            if (!match.Success)
                continue;
            var column1 = ColumnOfReference(match.Groups[1].Value);
            var row1 = RowOfReference(match.Groups[2].Value);
            var column2 = match.Groups[3].Success
                ? ColumnOfReference(match.Groups[3].Value) : column1;
            var row2 = match.Groups[4].Success
                ? RowOfReference(match.Groups[4].Value) : row1;
            if (column1 < 1 || row1 < 1 || column2 < 1 || row2 < 1)
                continue;
            ranges.Add(new MergeRange(
                Math.Min(column1, column2), Math.Min(row1, row2),
                Math.Max(column1, column2), Math.Max(row1, row2)));
        }
        return ranges;
    }

    /// <summary>
    /// 区域写入前校验：命中合并区但不是左上锚点时明确失败（openpyxl 对
    /// MergedCell 赋值同样失败）；只覆盖锚点本身时允许。
    /// </summary>
    private static void EnsureRangeWritable(
        S.Worksheet worksheet, int minColumn, int minRow, int maxColumn, int maxRow,
        string path)
    {
        foreach (var merge in MergeRanges(worksheet))
        {
            var overlapMinColumn = Math.Max(minColumn, merge.MinColumn);
            var overlapMinRow = Math.Max(minRow, merge.MinRow);
            var overlapMaxColumn = Math.Min(maxColumn, merge.MaxColumn);
            var overlapMaxRow = Math.Min(maxRow, merge.MaxRow);
            if (overlapMinColumn > overlapMaxColumn
                || overlapMinRow > overlapMaxRow)
                continue;
            if (overlapMinColumn == merge.MinColumn
                && overlapMinRow == merge.MinRow
                && overlapMaxColumn == merge.MinColumn
                && overlapMaxRow == merge.MinRow)
                continue;
            var badColumn = overlapMinColumn;
            var badRow = overlapMinRow;
            if (badColumn == merge.MinColumn && badRow == merge.MinRow)
            {
                if (overlapMaxColumn > merge.MinColumn)
                    badColumn = merge.MinColumn + 1;
                else
                    badRow = merge.MinRow + 1;
            }
            throw MergedCellError(badColumn, badRow, merge, path);
        }
    }

    /// <summary>values 区域与任何合并区交叠都直接报错，避免写入一半。</summary>
    private static void EnsureValuesClearOfMerges(
        S.Worksheet worksheet, int minColumn, int minRow, int maxColumn, int maxRow,
        string path)
    {
        foreach (var merge in MergeRanges(worksheet))
        {
            if (maxColumn < merge.MinColumn || minColumn > merge.MaxColumn
                || maxRow < merge.MinRow || minRow > merge.MaxRow)
                continue;
            throw new BatchEditException(
                "invalid_path",
                $"values 区域 {RangeReference(minColumn, minRow, maxColumn, maxRow)} "
                + $"与合并区域 {MergeReference(merge)} 交叠，无法批量写入",
                path,
                suggestion: "拆分合并单元格，或只写入合并区左上锚点");
        }
    }

    /// <summary>对已定位单元格（单元格/区域/整表路径）做合并区写保护。</summary>
    private static void EnsureCellsWritable(
        S.Worksheet worksheet, IEnumerable<S.Cell> cells, string path)
    {
        var merges = MergeRanges(worksheet);
        if (merges.Count == 0)
            return;
        foreach (var cell in cells)
        {
            var column = ColumnOfReference(cell.CellReference?.Value);
            var row = RowOfReference(cell.CellReference?.Value);
            if (column < 1 || row < 1)
                continue;
            foreach (var merge in merges)
            {
                if (column < merge.MinColumn || column > merge.MaxColumn
                    || row < merge.MinRow || row > merge.MaxRow)
                    continue;
                if (column == merge.MinColumn && row == merge.MinRow)
                    break;
                throw MergedCellError(column, row, merge, path);
            }
        }
    }

    private static BatchEditException MergedCellError(
        int column, int row, MergeRange merge, string path) =>
        new("unsupported_operation",
            $"单元格 {CellReference(column, row)} 位于合并区域 "
            + $"{MergeReference(merge)} 内且不是左上锚点，无法写入"
            + "（与 openpyxl 对 MergedCell 赋值失败一致）",
            path,
            suggestion: $"只写入合并区左上锚点 {CellReference(merge.MinColumn, merge.MinRow)}");

    private static string RangeReference(
        int minColumn, int minRow, int maxColumn, int maxRow) =>
        minColumn == maxColumn && minRow == maxRow
            ? CellReference(minColumn, minRow)
            : $"{CellReference(minColumn, minRow)}:"
              + CellReference(maxColumn, maxRow);

    private static string MergeReference(MergeRange merge) =>
        RangeReference(merge.MinColumn, merge.MinRow,
            merge.MaxColumn, merge.MaxRow);

    private static int MaxRow(S.Worksheet worksheet)
    {
        var maximum = 0;
        foreach (var cell in AllCells(worksheet))
            maximum = Math.Max(maximum, RowOfReference(cell.CellReference?.Value));
        return Math.Max(1, maximum);
    }

    private static int MaxColumn(S.Worksheet worksheet)
    {
        var maximum = 0;
        foreach (var cell in AllCells(worksheet))
            maximum = Math.Max(maximum,
                ColumnOfReference(cell.CellReference?.Value));
        return Math.Max(1, maximum);
    }

    // ── set ──────────────────────────────────────────────────────

    private static string Set(
        ExcelContext context, S.Worksheet? worksheet, string? selector,
        OfficeOperation operation)
    {
        if (worksheet is null || selector is null
            || selector.StartsWith("row[", StringComparison.Ordinal)
            || selector.StartsWith("col[", StringComparison.Ordinal))
            throw new BatchEditException(
                "invalid_path", "Excel set 需要单元格或区域路径",
                operation.Path);

        BatchEditJson.EnsureAllowed(operation, CellProps, "Excel 单元格");
        var (minColumn, minRow, maxColumn, maxRow) =
            ParseRange(selector, operation.Path);
        var hasValues = operation.Props.TryGetValue("values", out var values)
                        && values.ValueKind != JsonValueKind.Null;
        var writesProps = operation.Props.Keys.Any(key => key != "values");
        if (hasValues)
        {
            var expectedRows = maxRow - minRow + 1;
            var expectedColumns = maxColumn - minColumn + 1;
            ValidateValues(values, expectedRows, expectedColumns,
                operation.Path);
            EnsureValuesClearOfMerges(worksheet, minColumn, minRow,
                maxColumn, maxRow, operation.Path);
        }
        else if (writesProps)
        {
            EnsureRangeWritable(worksheet, minColumn, minRow, maxColumn, maxRow,
                operation.Path);
        }

        var cells = GetRangeCells(worksheet, minColumn, minRow, maxColumn,
            maxRow, out var index);
        if (hasValues)
        {
            var rowIndex = 0;
            foreach (var row in values.EnumerateArray())
            {
                var columnIndex = 0;
                foreach (var value in row.EnumerateArray())
                {
                    var cell = index.GetOrCreate(minRow + rowIndex,
                        minColumn + columnIndex);
                    SetCellFromJson(cell, value, interpretFormula: true);
                    columnIndex++;
                }
                rowIndex++;
            }
            var remaining = operation.Props
                .Where(pair => pair.Key != "values")
                .ToDictionary(pair => pair.Key, pair => pair.Value,
                    StringComparer.Ordinal);
            foreach (var cell in cells)
                ApplyCellProps(context, cell, remaining, operation.Path);
        }
        else
        {
            foreach (var cell in cells)
                ApplyCellProps(context, cell, operation.Props, operation.Path);
        }
        return $"已更新 {cells.Count} 个单元格";
    }

    private static void ValidateValues(
        JsonElement values, int expectedRows, int expectedColumns, string path)
    {
        var valid = values.ValueKind == JsonValueKind.Array
                    && values.GetArrayLength() == expectedRows;
        if (valid)
        {
            foreach (var row in values.EnumerateArray())
            {
                if (row.ValueKind != JsonValueKind.Array
                    || row.GetArrayLength() != expectedColumns)
                {
                    valid = false;
                    break;
                }
            }
        }
        if (!valid)
            throw new BatchEditException(
                "invalid_value",
                $"values 尺寸必须为 {expectedRows}x{expectedColumns}", path);
    }

    private static void ApplyCellProps(
        ExcelContext context, S.Cell cell,
        IReadOnlyDictionary<string, JsonElement> props, string path)
    {
        if (props.Count == 0)
            return;
        if (props.TryGetValue("value", out var value))
        {
            if (value.ValueKind == JsonValueKind.String)
            {
                var text = value.GetString() ?? "";
                if (text.StartsWith("=", StringComparison.Ordinal))
                    text = "'" + text;
                SetCellString(cell, text);
            }
            else
            {
                SetCellFromJson(cell, value, interpretFormula: false);
            }
        }
        if (props.TryGetValue("formula", out var formulaValue))
        {
            var formula = BatchEditJson.PythonString(formulaValue);
            SetCellFormula(cell, formula.StartsWith("=", StringComparison.Ordinal)
                ? formula[1..] : formula);
        }
        ApplyCellStyle(context, cell, props, path);
    }

    private static void ApplyCellStyle(
        ExcelContext context, S.Cell cell,
        IReadOnlyDictionary<string, JsonElement> props, string path)
    {
        var hasFont = FontProps.Any(props.ContainsKey);
        var hasFill = props.ContainsKey("fill");
        var hasAlignment = props.ContainsKey("alignment");
        var hasNumberFormat = props.ContainsKey("number_format");
        if (!hasFont && !hasFill && !hasAlignment && !hasNumberFormat)
            return;

        var styleIndex = cell.StyleIndex?.Value ?? 0;
        var source = GetCellFormat(context, styleIndex);
        var updated = (S.CellFormat)source.CloneNode(true);

        if (hasFont)
        {
            var font = GetFont(context, source.FontId?.Value ?? 0);
            var updatedFont = (S.Font)font.CloneNode(true);
            if (props.TryGetValue("bold", out var bold))
                SetChildInOrder(updatedFont,
                    new S.Bold { Val = BatchEditJson.Truthy(bold) }, FontOrder);
            if (props.TryGetValue("italic", out var italic))
                SetChildInOrder(updatedFont,
                    new S.Italic { Val = BatchEditJson.Truthy(italic) }, FontOrder);
            if (props.TryGetValue("font_name", out var name))
                SetChildInOrder(updatedFont,
                    new S.FontName { Val = BatchEditJson.PythonString(name) },
                    FontOrder);
            if (props.TryGetValue("font_size", out var size))
                SetChildInOrder(updatedFont,
                    new S.FontSize
                    {
                        Val = BatchEditJson.ToDoubleValue(
                            size, "font_size", path),
                    }, FontOrder);
            if (props.TryGetValue("font_color", out var color))
                SetChildInOrder(updatedFont,
                    new S.Color
                    {
                        Rgb = "00" + BatchEditJson.RequireHexColorValue(
                            color, "font_color", path),
                    }, FontOrder);
            updated.FontId = (uint)GetOrAddFont(context, updatedFont);
            updated.ApplyFont = true;
        }

        if (hasFill)
        {
            var color = BatchEditJson.RequireHexColorValue(
                props["fill"], "fill", path);
            updated.FillId = (uint)GetOrAddFill(context, color);
            updated.ApplyFill = true;
        }

        if (hasAlignment)
        {
            var value = BatchEditJson.PythonString(props["alignment"])
                .ToLowerInvariant();
            var horizontal = value switch
            {
                "left" => S.HorizontalAlignmentValues.Left,
                "center" => S.HorizontalAlignmentValues.Center,
                "right" => S.HorizontalAlignmentValues.Right,
                _ => throw new BatchEditException(
                    "invalid_value", $"无效的对齐方式: {value}", path,
                    validValues: ["left", "center", "right"]),
            };
            SetCellFormatAlignment(updated, horizontal);
            updated.ApplyAlignment = true;
        }

        if (hasNumberFormat)
        {
            var code = BatchEditJson.PythonString(props["number_format"]);
            updated.NumberFormatId = GetOrAddNumberFormat(context, code);
            updated.ApplyNumberFormat = true;
        }

        cell.StyleIndex = (uint)GetOrAddCellFormat(context, updated);
    }

    private static S.CellFormat GetCellFormat(ExcelContext context, uint index)
    {
        var formats = context.Stylesheet.CellFormats
                          ?.Elements<S.CellFormat>().ToList() ?? [];
        if (index < formats.Count)
            return formats[(int)index];
        return new S.CellFormat
        {
            NumberFormatId = 0,
            FontId = 0,
            FillId = 0,
            BorderId = 0,
        };
    }

    private static S.Font GetFont(ExcelContext context, uint index)
    {
        var fonts = context.Stylesheet.Fonts?.Elements<S.Font>().ToList() ?? [];
        return index < fonts.Count ? fonts[(int)index] : new S.Font();
    }

    private static int GetOrAddFont(ExcelContext context, S.Font font)
    {
        var fonts = context.Stylesheet.Fonts
                    ?? throw new NativeRuntimeException("工作簿缺少 fonts 样式区");
        var key = font.OuterXml;
        var existing = fonts.Elements<S.Font>().ToList();
        for (var index = 0; index < existing.Count; index++)
        {
            if (existing[index].OuterXml == key)
                return index;
        }
        fonts.Append(font);
        return existing.Count;
    }

    private static int GetOrAddFill(ExcelContext context, string hex)
    {
        var fills = context.Stylesheet.Fills
                    ?? throw new NativeRuntimeException("工作簿缺少 fills 样式区");
        var fill = new S.Fill(
            new S.PatternFill
            {
                PatternType = S.PatternValues.Solid,
                ForegroundColor = new S.ForegroundColor { Rgb = "00" + hex },
                BackgroundColor = new S.BackgroundColor { Indexed = 64U },
            });
        var key = fill.OuterXml;
        var existing = fills.Elements<S.Fill>().ToList();
        for (var index = 0; index < existing.Count; index++)
        {
            if (existing[index].OuterXml == key)
                return index;
        }
        fills.Append(fill);
        return existing.Count;
    }

    private static uint GetOrAddNumberFormat(ExcelContext context, string code)
    {
        if (BuiltinNumberFormats.TryGetValue(code, out var builtin))
            return builtin;
        var container = context.Stylesheet.NumberingFormats;
        if (container is null)
        {
            container = new S.NumberingFormats();
            context.Stylesheet.PrependChild(container);
        }
        var existing = container.Elements<S.NumberingFormat>().ToList();
        foreach (var format in existing)
        {
            if (format.FormatCode?.Value == code)
                return format.NumberFormatId?.Value ?? 164U;
        }
        var next = 163U;
        foreach (var format in existing)
            next = Math.Max(next, format.NumberFormatId?.Value ?? 0U);
        next++;
        container.Append(new S.NumberingFormat
        {
            NumberFormatId = next,
            FormatCode = code,
        });
        container.Count = (uint)(existing.Count + 1);
        return next;
    }

    private static int GetOrAddCellFormat(
        ExcelContext context, S.CellFormat cellFormat)
    {
        var formats = context.Stylesheet.CellFormats
                      ?? throw new NativeRuntimeException("工作簿缺少 cellXfs 样式区");
        var key = cellFormat.OuterXml;
        var existing = formats.Elements<S.CellFormat>().ToList();
        for (var index = 0; index < existing.Count; index++)
        {
            if (existing[index].OuterXml == key)
                return index;
        }
        formats.Append(cellFormat);
        return existing.Count;
    }

    private static void SetCellFormatAlignment(
        S.CellFormat cellFormat, S.HorizontalAlignmentValues horizontal)
    {
        S.Alignment alignment;
        var current = cellFormat.Alignment;
        if (current is not null)
        {
            alignment = (S.Alignment)current.CloneNode(true);
            cellFormat.RemoveChild(current);
        }
        else
        {
            alignment = new S.Alignment();
        }
        alignment.Horizontal = horizontal;
        var protection = cellFormat.GetFirstChild<S.Protection>();
        var extension = cellFormat.GetFirstChild<S.ExtensionList>();
        if (protection is not null)
            cellFormat.InsertBefore(alignment, protection);
        else if (extension is not null)
            cellFormat.InsertBefore(alignment, extension);
        else
            cellFormat.Append(alignment);
    }

    // ── 值写入 ───────────────────────────────────────────────────

    private static void SetCellFromJson(
        S.Cell cell, JsonElement value, bool interpretFormula)
    {
        ClearCellContent(cell);
        switch (value.ValueKind)
        {
            case JsonValueKind.Null:
            case JsonValueKind.Undefined:
                return;
            case JsonValueKind.True:
            case JsonValueKind.False:
                cell.DataType = S.CellValues.Boolean;
                AppendCellChild(cell,
                    new S.CellValue(value.GetBoolean() ? "1" : "0"));
                return;
            case JsonValueKind.Number:
                AppendCellChild(cell, new S.CellValue(value.GetRawText()));
                return;
            case JsonValueKind.String:
            {
                var text = value.GetString() ?? "";
                if (interpretFormula
                    && text.StartsWith("=", StringComparison.Ordinal))
                    AppendCellChild(cell,
                        new S.CellFormula(text[1..]));
                else
                    AppendCellText(cell, text);
                return;
            }
            default:
                throw new BatchEditException(
                    "invalid_value",
                    "value 只支持字符串/数字/布尔/null");
        }
    }

    private static void SetCellString(S.Cell cell, string text)
    {
        ClearCellContent(cell);
        AppendCellText(cell, text);
    }

    private static void SetCellFormula(S.Cell cell, string formula)
    {
        ClearCellContent(cell);
        AppendCellChild(cell, new S.CellFormula(formula));
    }

    private static void ClearCellContent(S.Cell cell)
    {
        cell.CellFormula?.Remove();
        cell.CellValue?.Remove();
        cell.InlineString?.Remove();
        cell.DataType = null;
    }

    private static void AppendCellText(S.Cell cell, string text)
    {
        if (text.Length > 32767)
            throw new BatchEditException(
                "invalid_value", "字符串超过 Excel 单元格上限 32767 字符");
        cell.DataType = S.CellValues.InlineString;
        AppendCellChild(cell, new S.InlineString(
            new S.Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static void AppendCellChild(S.Cell cell, OpenXmlElement child)
    {
        var extension = cell.GetFirstChild<S.ExtensionList>();
        if (extension is not null)
            cell.InsertBefore(child, extension);
        else
            cell.Append(child);
    }

    // ── replace ──────────────────────────────────────────────────

    private static string Replace(
        ExcelContext context, S.Worksheet? worksheet, string? selector,
        OfficeOperation operation)
    {
        if (worksheet is null)
            throw new BatchEditException(
                "invalid_path", "Excel replace 需要工作表或区域路径",
                operation.Path);
        var find = BatchEditJson.PythonStringProp(operation, "find");
        var replacement = BatchEditJson.PythonStringProp(operation, "replace");
        if (find.Length == 0)
            throw new BatchEditException(
                "invalid_value", "replace 操作的 find 不能为空", operation.Path);
        var cells = CellsForSelector(worksheet, selector, operation.Path);
        EnsureCellsWritable(worksheet, cells, operation.Path);
        var count = 0;
        foreach (var cell in cells)
        {
            if (!TryGetStringCellValue(context, cell, out var text))
                continue;
            var occurrences = CountOccurrences(text, find);
            if (occurrences == 0)
                continue;
            count += occurrences;
            var changed = text.Replace(find, replacement, StringComparison.Ordinal);
            if (changed.StartsWith("=", StringComparison.Ordinal))
                SetCellFormula(cell, changed[1..]);
            else
                SetCellString(cell, changed);
        }
        return $"已替换 {count} 处文本";
    }

    private static bool TryGetStringCellValue(
        ExcelContext context, S.Cell cell, out string text)
    {
        text = "";
        if (cell.CellFormula?.Text is { } formula)
        {
            text = "=" + formula;
            return true;
        }
        if (cell.InlineString is not null)
        {
            // 富文本（多 run 或带 rPr 的单 run）在 openpyxl 中是 CellRichText，
            // isinstance(value, str) 为 False；replace 必须跳过而不是重建文本。
            if (IsRichTextContainer(cell.InlineString))
                return false;
            text = PlainTextOf(cell.InlineString);
            return true;
        }
        if (cell.CellValue is null)
            return false;
        var type = cell.DataType?.Value;
        if (type == S.CellValues.SharedString)
        {
            if (!int.TryParse(cell.CellValue.Text, NumberStyles.Integer,
                    CultureInfo.InvariantCulture, out var index)
                || index < 0 || index >= context.SharedStrings.Count)
                return false;
            var shared = context.SharedStrings[index];
            if (IsRichTextContainer(shared))
                return false;
            text = PlainTextOf(shared);
            return true;
        }
        if (type == S.CellValues.String)
        {
            text = cell.CellValue.Text;
            return true;
        }
        return false;
    }

    /// <summary>
    /// 普通字符串取值：直接 &lt;t&gt; 优先（与 openpyxl Text.from_tree 一致），
    /// 否则只拼接 run 文本，不混入 rPh 注音文本。
    /// </summary>
    private static string PlainTextOf(OpenXmlElement container)
    {
        var direct = container.GetFirstChild<S.Text>();
        if (direct is not null)
            return direct.Text ?? "";
        return string.Concat(container.Elements<S.Run>()
            .Select(run => run.GetFirstChild<S.Text>()?.Text ?? ""));
    }

    /// <summary>
    /// 与 openpyxl 富文本判定一致：直接文本节点 &lt;t&gt; → 普通字符串；只有 run
    /// 时，单 run 且无 rPr 仍是普通字符串，其余均为 CellRichText（不可按 str 编辑）。
    /// </summary>
    private static bool IsRichTextContainer(OpenXmlElement container)
    {
        if (container.GetFirstChild<S.Text>() is not null)
            return false;
        var runs = container.Elements<S.Run>().ToList();
        return runs.Count > 1
               || (runs.Count == 1 && runs[0].RunProperties is not null);
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

    // ── add ──────────────────────────────────────────────────────

    private static string Add(
        ExcelContext context, S.Worksheet? worksheet, string? selector,
        OfficeOperation operation)
    {
        if (worksheet is null)
        {
            var elementType = BatchEditJson.RequireElementType(
                operation, "Excel", ["sheet"]);
            if (elementType != "sheet")
                throw new BatchEditException(
                    "invalid_value", "工作簿根节点只能添加 sheet",
                    operation.Path, validValues: ["sheet"]);
            return AddSheet(context, operation);
        }
        if (selector is not null)
            throw new BatchEditException(
                "invalid_path", "添加行或列时路径应指向工作表",
                operation.Path);

        var type = BatchEditJson.RequireElementType(
            operation, "Excel", ["row", "column"]);
        if (type is not ("row" or "column"))
            throw new BatchEditException(
                "invalid_value", $"不支持的 Excel 元素类型: {type}",
                operation.Path, validValues: ["row", "column"]);

        var values = new List<JsonElement>();
        if (operation.Props.TryGetValue("values", out var valuesElement)
            && valuesElement.ValueKind != JsonValueKind.Null)
        {
            if (valuesElement.ValueKind != JsonValueKind.Array)
                throw new BatchEditException(
                    "invalid_value", "values 必须是数组", operation.Path);
            values.AddRange(valuesElement.EnumerateArray());
        }
        var index = BatchEditJson.TryGet(operation, "index", out var indexValue)
            ? BatchEditJson.ToIntValue(indexValue, "index", operation.Path)
            : type == "row" ? MaxRow(worksheet) + 1 : MaxColumn(worksheet) + 1;
        if (index < 1)
            throw new BatchEditException(
                "invalid_value", $"{type} 索引必须大于等于 1", operation.Path);

        if (type == "row")
        {
            ShiftRowsDown(worksheet, index);
            for (var offset = 0; offset < values.Count; offset++)
            {
                var cell = GetCell(worksheet, index, offset + 1, true)!;
                SetCellFromJson(cell, values[offset], interpretFormula: true);
            }
        }
        else
        {
            ShiftColumnsRight(worksheet, index);
            for (var offset = 0; offset < values.Count; offset++)
            {
                var cell = GetCell(worksheet, offset + 1, index, true)!;
                SetCellFromJson(cell, values[offset], interpretFormula: true);
            }
        }
        return $"已添加 {type}";
    }

    private static string AddSheet(
        ExcelContext context, OfficeOperation operation)
    {
        var name = BatchEditJson.PythonStringProp(operation, "name", "Sheet");
        if (name.Length > 31)
            name = name[..31];
        if (string.IsNullOrWhiteSpace(name))
            throw new BatchEditException(
                "invalid_value", "工作表名称不能为空", operation.Path);
        if (name.IndexOfAny(['[', ']', ':', '*', '?', '/', '\\']) >= 0)
            throw new BatchEditException(
                "invalid_value", $"工作表名称包含非法字符: {name}",
                operation.Path);
        if (context.Sheets.Elements<S.Sheet>().Any(sheet =>
                string.Equals(sheet.Name?.Value, name,
                    StringComparison.OrdinalIgnoreCase)))
            throw new BatchEditException(
                "already_exists", $"工作表已存在: {name}", operation.Path);

        var worksheetPart = context.WorkbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new S.Worksheet(new S.SheetData());
        var nextId = 0U;
        foreach (var sheet in context.Sheets.Elements<S.Sheet>())
            nextId = Math.Max(nextId, sheet.SheetId?.Value ?? 0U);
        nextId++;
        var added = new S.Sheet
        {
            Name = name,
            SheetId = nextId,
            Id = context.WorkbookPart.GetIdOfPart(worksheetPart),
        };
        context.Sheets.Append(added);
        return $"已添加工作表 {name}";
    }

    // ── remove ───────────────────────────────────────────────────

    private static string Remove(
        ExcelContext context, S.Worksheet? worksheet, string? selector,
        S.Sheet? sheet, WorksheetPart? part, OfficeOperation operation)
    {
        if (worksheet is null)
            throw new BatchEditException(
                "unsupported_operation", "不能删除整个工作簿", operation.Path);

        if (selector is null)
        {
            if (context.Sheets.Elements<S.Sheet>().Count() <= 1)
                throw new BatchEditException(
                    "invalid_value", "工作簿必须至少保留一个工作表",
                    operation.Path);
            if (sheet is not null)
                context.Sheets.RemoveChild(sheet);
            if (part is not null)
                DeletePartTree(context.WorkbookPart, part);
            return $"已删除 {operation.Path}";
        }

        var rowMatch = RowPattern.Match(selector);
        if (rowMatch.Success)
        {
            var index = int.Parse(rowMatch.Groups[1].Value,
                CultureInfo.InvariantCulture);
            if (index < 1)
                throw new BatchEditException(
                    "invalid_value", "行索引必须大于等于 1", operation.Path);
            ShiftRowsUp(worksheet, index);
            return $"已删除 {operation.Path}";
        }
        var columnMatch = ColumnPattern.Match(selector);
        if (columnMatch.Success)
        {
            var index = ColumnNumber(columnMatch.Groups[1].Value, operation.Path);
            ShiftColumnsLeft(worksheet, index);
            return $"已删除 {operation.Path}";
        }

        var cells = CellsForSelector(worksheet, selector, operation.Path);
        EnsureCellsWritable(worksheet, cells, operation.Path);
        foreach (var cell in cells)
            ClearCellContent(cell);
        return $"已删除 {operation.Path}";
    }

    // ── move ─────────────────────────────────────────────────────

    private static string Move(
        ExcelContext context, S.Worksheet? worksheet, string? selector,
        S.Sheet? sheet, OfficeOperation operation)
    {
        if (worksheet is null || selector is not null || sheet is null
            || string.IsNullOrEmpty(operation.TargetPath))
            throw new BatchEditException(
                "unsupported_operation", "Excel move 当前仅支持工作表排序",
                operation.Path);
        var (targetWorksheet, targetSelector, targetSheet, _) =
            Resolve(context, operation.TargetPath);
        if (targetWorksheet is null || targetSelector is not null
            || targetSheet is null
            || ReferenceEquals(targetWorksheet, worksheet)
            || targetSheet.Id?.Value == sheet.Id?.Value)
            throw new BatchEditException(
                "invalid_path", "target_path 必须是另一个工作表",
                operation.TargetPath);

        sheet.Remove();
        if (operation.Position == "before")
            targetSheet.InsertBeforeSelf(sheet);
        else
            targetSheet.InsertAfterSelf(sheet);
        return $"已移动工作表 {sheet.Name?.Value}";
    }

    private static void DeletePartTree(
        OpenXmlPartContainer parent, OpenXmlPart part)
    {
        foreach (var child in part.Parts
                     .Select(pair => pair.OpenXmlPart).ToList())
            DeletePartTree(part, child);
        parent.DeletePart(part);
    }

    // ── 行列位移 ─────────────────────────────────────────────────

    private static void ShiftRowsDown(S.Worksheet worksheet, int index)
    {
        if (index > 1048576)
            throw new BatchEditException(
                "invalid_value", $"行索引超出范围: {index}", "/");
        var sheetData = EnsureSheetData(worksheet);
        var rows = sheetData.Elements<S.Row>()
            .Where(row => row.RowIndex?.Value is not null)
            .OrderByDescending(row => row.RowIndex!.Value)
            .ToList();
        foreach (var row in rows)
        {
            var current = (int)row.RowIndex!.Value;
            if (current < index)
                continue;
            row.RowIndex = (uint)(current + 1);
            foreach (var cell in row.Elements<S.Cell>())
                UpdateCellRow(cell, current + 1);
        }
        SortRows(sheetData);
    }

    private static void ShiftRowsUp(S.Worksheet worksheet, int index)
    {
        var sheetData = EnsureSheetData(worksheet);
        var target = sheetData.Elements<S.Row>().FirstOrDefault(row =>
            row.RowIndex?.Value == (uint)index);
        if (target is null)
            return;
        target.Remove();
        foreach (var row in sheetData.Elements<S.Row>()
                     .Where(row => (row.RowIndex?.Value ?? 0) > (uint)index)
                     .ToList())
        {
            var current = (int)row.RowIndex!.Value;
            row.RowIndex = (uint)(current - 1);
            foreach (var cell in row.Elements<S.Cell>())
                UpdateCellRow(cell, current - 1);
        }
        SortRows(sheetData);
    }

    private static void ShiftColumnsRight(S.Worksheet worksheet, int index)
    {
        if (index > 16384)
            throw new BatchEditException(
                "invalid_value", $"列索引超出范围: {index}", "/");
        var sheetData = EnsureSheetData(worksheet);
        foreach (var row in sheetData.Elements<S.Row>())
        {
            var cells = row.Elements<S.Cell>()
                .Where(cell => ColumnOfReference(cell.CellReference?.Value) >= index)
                .OrderByDescending(cell =>
                    ColumnOfReference(cell.CellReference?.Value))
                .ToList();
            foreach (var cell in cells)
            {
                var column = ColumnOfReference(cell.CellReference?.Value);
                UpdateCellColumn(cell, column + 1);
            }
            SortCells(row);
        }
    }

    private static void ShiftColumnsLeft(S.Worksheet worksheet, int index)
    {
        var sheetData = EnsureSheetData(worksheet);
        foreach (var row in sheetData.Elements<S.Row>())
        {
            var removed = row.Elements<S.Cell>().FirstOrDefault(cell =>
                ColumnOfReference(cell.CellReference?.Value) == index);
            removed?.Remove();
            var cells = row.Elements<S.Cell>()
                .Where(cell => ColumnOfReference(cell.CellReference?.Value) > index)
                .OrderBy(cell => ColumnOfReference(cell.CellReference?.Value))
                .ToList();
            foreach (var cell in cells)
            {
                var column = ColumnOfReference(cell.CellReference?.Value);
                UpdateCellColumn(cell, column - 1);
            }
            SortCells(row);
        }
    }

    private static void UpdateCellRow(S.Cell cell, int row)
    {
        var reference = cell.CellReference?.Value;
        if (string.IsNullOrEmpty(reference))
            return;
        var column = ColumnOfReference(reference);
        if (column > 0)
            cell.CellReference = CellReference(column, row);
    }

    private static void UpdateCellColumn(S.Cell cell, int column)
    {
        var reference = cell.CellReference?.Value;
        if (string.IsNullOrEmpty(reference))
            return;
        var row = RowOfReference(reference);
        if (row > 0)
            cell.CellReference = CellReference(column, row);
    }

    private static void SortRows(S.SheetData sheetData)
    {
        var ordered = sheetData.Elements<S.Row>()
            .OrderBy(row => row.RowIndex?.Value ?? uint.MaxValue)
            .ToList();
        foreach (var row in ordered)
            row.Remove();
        foreach (var row in ordered)
            sheetData.Append(row);
    }

    private static void SortCells(S.Row row)
    {
        var ordered = row.Elements<S.Cell>()
            .OrderBy(cell => ColumnOfReference(cell.CellReference?.Value))
            .ToList();
        foreach (var cell in ordered)
            cell.Remove();
        foreach (var cell in ordered)
            row.Append(cell);
    }

    // ── 样式辅助 ─────────────────────────────────────────────────

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
