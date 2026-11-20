using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeNative;

/// <summary>Excel（.xlsx/.xlsm）文件级创建、编辑、读取与模板合并（Open XML SDK）。</summary>
internal static partial class ExcelFileOoxml
{
    private static readonly string[] OutputExtensions = [".xlsx", ".xlsm"];
    private static readonly string[] TemplateExtensions = [".xlsx", ".xlsm", ".xltx", ".xltm"];

    internal const int MaxColumns = 16384;
    internal const int MaxRows = 1048576;

    // ── 创建 ─────────────────────────────────────────────────────

    public static Dictionary<string, object?> CreateWorkbook(
        string outputPath,
        string? template,
        List<ExcelSheetDef> sheets,
        string ifExists,
        bool confirmOverwrite)
    {
        ValidateOutputPath(outputPath, OutputExtensions);
        if (sheets.Count == 0)
            throw new NativeRuntimeException("sheets 至少需要 1 个工作表");
        if (template is not null)
            NativeCommon.EnsureFile(template, TemplateExtensions);

        var outputExtension = Path.GetExtension(outputPath).ToLowerInvariant();
        var outputIsMacro = outputExtension == ".xlsm";
        var templateExtension = template is null
            ? null
            : Path.GetExtension(template).ToLowerInvariant();
        var templateIsMacro = templateExtension is ".xlsm" or ".xltm";
        if (outputIsMacro && !templateIsMacro)
            throw new NativeRuntimeException(
                "生成 .xlsm 必须提供 .xlsm 或 .xltm 宏模板，以避免扩展名与内容不匹配");
        if (templateIsMacro && !outputIsMacro)
            throw new NativeRuntimeException(
                "使用宏模板时输出文件必须为 .xlsm，以避免丢失 VBA 内容");

        var target = ResolveOutputPath(outputPath, ifExists, confirmOverwrite);
        var fullTarget = Path.GetFullPath(target);
        Directory.CreateDirectory(Path.GetDirectoryName(fullTarget)!);
        var temp = TempPath(fullTarget);
        string? backupPath = null;
        try
        {
            if (template is not null)
                BuildFromTemplate(temp, template, outputExtension, sheets);
            else
                BuildNewWorkbook(temp, outputIsMacro, sheets);

            if (File.Exists(fullTarget) && new FileInfo(fullTarget).Length > 0)
                backupPath = NativeCommon.Backup(fullTarget);
            ReplaceFile(temp, fullTarget);
            return new Dictionary<string, object?>
            {
                ["success"] = true,
                ["file_path"] = target,
                ["file_size"] = new FileInfo(fullTarget).Length,
                ["sheets_count"] = sheets.Count,
                ["message"] = $"Excel 工作簿创建成功: {target}",
                ["backup_path"] = backupPath,
            };
        }
        finally
        {
            TryDelete(temp);
        }
    }

    private static void BuildNewWorkbook(
        string tempPath, bool macroEnabled, List<ExcelSheetDef> sheets)
    {
        using var document = SpreadsheetDocument.Create(
            tempPath,
            macroEnabled
                ? SpreadsheetDocumentType.MacroEnabledWorkbook
                : SpreadsheetDocumentType.Workbook);
        document.AddWorkbookPart();
        document.WorkbookPart!.Workbook = new Workbook();
        var context = new ExcelWorkbookContext(document);
        context.AddSheets(sheets);
        context.Save();
    }

    private static void BuildFromTemplate(
        string tempPath, string templatePath, string outputExtension, List<ExcelSheetDef> sheets)
    {
        File.Copy(templatePath, tempPath, overwrite: true);
        ClearReadOnly(tempPath);
        using var document = SpreadsheetDocument.Open(tempPath, true);
        var templateExtension = Path.GetExtension(templatePath).ToLowerInvariant();
        if (templateExtension is ".xltx" or ".xltm")
        {
            document.ChangeDocumentType(outputExtension == ".xlsm"
                ? SpreadsheetDocumentType.MacroEnabledWorkbook
                : SpreadsheetDocumentType.Workbook);
        }
        var context = new ExcelWorkbookContext(document);
        context.DeleteAllSheets();
        context.AddSheets(sheets);
        context.NormalizeWorkbookViews();
        context.Save();
    }

    // ── 追加工作表 ───────────────────────────────────────────────

    public static Dictionary<string, object?> EditWorkbook(
        string filePath, List<ExcelSheetDef> appendSheets)
    {
        try
        {
            NativeCommon.EnsureFile(filePath, OutputExtensions);
        }
        catch (NativeRuntimeException)
        {
            throw;
        }
        catch (Exception exception) when (
            exception is IOException or UnauthorizedAccessException)
        {
            throw LockedFileError(filePath, exception);
        }
        var fullTarget = Path.GetFullPath(filePath);
        var temp = TempPath(fullTarget);
        string? backupPath = null;
        try
        {
            CopyForEdit(fullTarget, temp);
            using (var document = SpreadsheetDocument.Open(temp, true))
            {
                var context = new ExcelWorkbookContext(document);
                context.AddSheets(appendSheets);
                context.Save();
            }
            try
            {
                backupPath = NativeCommon.Backup(fullTarget);
                ReplaceFile(temp, fullTarget);
            }
            catch (Exception exception) when (
                exception is IOException or UnauthorizedAccessException)
            {
                throw LockedFileError(filePath, exception);
            }
            return new Dictionary<string, object?>
            {
                ["success"] = true,
                ["file_path"] = filePath,
                ["operation_count"] = appendSheets.Count,
                ["message"] = $"已追加 {appendSheets.Count} 个工作表到 Excel 文件",
                ["backup_path"] = backupPath,
            };
        }
        finally
        {
            TryDelete(temp);
        }
    }

    // ── 读取 ─────────────────────────────────────────────────────

    public static Dictionary<string, object?> ReadWorkbook(
        string filePath, int maxRowsPerSheet, int maxChars)
    {
        if (maxRowsPerSheet < 1 || maxChars < 1)
            throw new NativeUsageException("max_rows_per_sheet 和 max_chars 必须大于 0");
        NativeCommon.EnsureFile(filePath, OutputExtensions);

        try
        {
            using var document = SpreadsheetDocument.Open(filePath, false);
            var workbookPart = document.WorkbookPart
                               ?? throw new NativeRuntimeException("工作簿缺少 WorkbookPart");
            var sharedStrings = LoadSharedStringTexts(workbookPart);
            var formats = ExcelReadFormats.Load(workbookPart);
            var parts = new List<string>();
            var currentLength = 0;
            var truncated = false;
            var sheetCount = 0;
            var totalRows = 0;

            bool AddPart(string part)
            {
                var separator = parts.Count > 0 ? 1 : 0;
                if (currentLength + separator + part.Length > maxChars)
                {
                    truncated = true;
                    return false;
                }
                parts.Add(part);
                currentLength += separator + part.Length;
                return true;
            }

            var sheets = workbookPart.Workbook?.Sheets?.Elements<Sheet>().ToList() ?? [];
            foreach (var sheet in sheets)
            {
                sheetCount++;
                var sheetName = sheet.Name?.Value ?? "";
                if (!AddPart($"## Sheet: {sheetName}"))
                    break;

                var worksheetPart = GetWorksheetPart(workbookPart, sheet);
                var (rowCount, rowLimitHit, charLimitHit) = worksheetPart is null
                    ? (0, false, false)
                    : ReadSheetRows(
                        worksheetPart, maxRowsPerSheet, sharedStrings, formats, AddPart);
                totalRows += rowCount;

                if (rowLimitHit)
                {
                    // 行数上限与 Python 一致：先写截断标记，再终止整次读取。
                    AddPart($"  _(已截断，超过 {maxRowsPerSheet} 行限制)_");
                    truncated = true;
                }
                if (truncated || charLimitHit)
                    break;
                if (rowCount == 0)
                    AddPart("  _(空工作表)_");
            }

            var content = string.Join("\n", parts);
            if (truncated)
            {
                const string marker = "\n_(内容已截断)_";
                var keep = Math.Min(content.Length, Math.Max(0, maxChars - marker.Length));
                content = content[..keep] + marker;
                if (content.Length > maxChars)
                    content = content[..maxChars];
            }

            var result = new Dictionary<string, object?>
            {
                ["success"] = true,
                ["file_path"] = filePath,
                ["content"] = content,
                ["sheets_count"] = sheetCount,
                ["total_rows"] = totalRows,
                ["truncated"] = truncated,
                ["message"] = $"成功读取工作簿，共 {sheetCount} 个工作表、{totalRows} 行数据"
                              + (truncated ? $"（已截断至 {maxChars} 字符）" : ""),
            };
            return new Dictionary<string, object?>
            {
                ["markdown"] = BuildReadMarkdown(result),
                ["result"] = result,
            };
        }
        catch (NativeRuntimeException)
        {
            throw;
        }
        catch (Exception exception)
        {
            throw new NativeRuntimeException($"读取 Excel 工作簿失败: {exception.Message}");
        }
    }

    private static (int RowCount, bool RowLimitHit, bool CharLimitHit) ReadSheetRows(
        WorksheetPart worksheetPart,
        int maxRowsPerSheet,
        IReadOnlyList<string> sharedStrings,
        ExcelReadFormats formats,
        Func<string, bool> addPart)
    {
        int? dimensionMaxRow = null;
        int? dimensionMaxColumn = null;
        var rowCount = 0;
        var lastRowIndex = 0;
        var rowLimitHit = false;
        var charLimitHit = false;

        // OpenXmlReader 流式读取：只在命中 Row 时加载该行 DOM，达到上限立即停止。
        using var reader = OpenXmlReader.Create(worksheetPart);
        while (reader.Read())
        {
            if (reader.ElementType == typeof(SheetDimension))
            {
                var dimension = (SheetDimension)reader.LoadCurrentElement();
                if (ExcelA1.TryParseRange(dimension.Reference?.Value ?? "", out var range))
                {
                    dimensionMaxRow = range.MaxRow;
                    dimensionMaxColumn = range.MaxCol;
                }
                continue;
            }
            if (reader.ElementType != typeof(Row))
                continue;

            var row = (Row)reader.LoadCurrentElement();
            var rawIndex = (long)(row.RowIndex?.Value ?? 0);
            var rowIndex = rawIndex > lastRowIndex
                ? (int)Math.Min(rawIndex, int.MaxValue)
                : lastRowIndex + 1;
            lastRowIndex = rowIndex;

            // openpyxl read_only 以 dimension 的 max_row 为上限，超出部分直接忽略。
            if (dimensionMaxRow is { } dimensionMax && rowIndex > dimensionMax)
                break;
            if (rowIndex > maxRowsPerSheet)
            {
                rowLimitHit = true;
                break;
            }

            var values = RowValues(row, sharedStrings, formats, dimensionMaxColumn);
            if (values.Length == 0 || values.All(value => value.Length == 0))
                continue;
            if (!addPart("| " + string.Join(" | ", values) + " |"))
            {
                charLimitHit = true;
                break;
            }
            rowCount++;
        }

        if (!rowLimitHit && !charLimitHit
            && dimensionMaxRow is { } maxRow && maxRow > maxRowsPerSheet)
        {
            // dimension 声明的尾部空行也计入 Python 的行数上限。
            rowLimitHit = true;
        }
        return (rowCount, rowLimitHit, charLimitHit);
    }

    private static string[] RowValues(
        Row row,
        IReadOnlyList<string> sharedStrings,
        ExcelReadFormats formats,
        int? dimensionMaxColumn)
    {
        var cells = new List<(int Column, string Value)>();
        var previousColumn = 0;
        foreach (var cell in row.Elements<Cell>())
        {
            var column = CellColumn(cell);
            if (column < 1)
                column = previousColumn + 1;
            previousColumn = column;
            if (column > MaxColumns)
                continue;
            // openpyxl 以 dimension 的 max_column 作为每行宽度，超出该列的单元格被丢弃。
            if (dimensionMaxColumn is { } dimensionColumn && column > dimensionColumn)
                continue;
            cells.Add((column, CellDisplay(cell, sharedStrings, formats)));
        }
        var maxColumn = dimensionMaxColumn
                        ?? (cells.Count == 0 ? 0 : cells.Max(item => item.Column));
        if (maxColumn <= 0)
            return [];
        var values = new string[maxColumn];
        Array.Fill(values, "");
        foreach (var (column, value) in cells)
            values[column - 1] = value;
        return values;
    }

    private static List<string> LoadSharedStringTexts(WorkbookPart workbookPart)
    {
        var items = workbookPart.SharedStringTablePart?.SharedStringTable
            ?.Elements<SharedStringItem>();
        return items is null
            ? []
            : items.Select(item => RichTextContent(item, stripEscapedUnderscore: true)).ToList();
    }

    /// <summary>按 openpyxl Text.content 语义拼接 &lt;t&gt; 与 &lt;r&gt;&lt;t&gt;，忽略 rPh/extLst。</summary>
    private static string RichTextContent(OpenXmlElement? element, bool stripEscapedUnderscore)
    {
        if (element is null)
            return "";
        var builder = new StringBuilder();
        foreach (var child in element.ChildElements)
        {
            if (child is Text text)
            {
                builder.Append(text.Text);
            }
            else if (child is Run run)
            {
                foreach (var runText in run.Elements<Text>())
                    builder.Append(runText.Text);
            }
        }
        var value = builder.ToString();
        return stripEscapedUnderscore
            ? value.Replace("x005F_", "", StringComparison.Ordinal)
            : value;
    }

    private static string CellDisplay(
        Cell cell, IReadOnlyList<string> sharedStrings, ExcelReadFormats formats)
    {
        var type = cell.DataType?.Value;
        if (type == CellValues.InlineString)
            return RichTextContent(cell.InlineString, stripEscapedUnderscore: false);

        var raw = cell.CellValue?.InnerText ?? "";
        // data_only=True：公式没有缓存值时 openpyxl 得到 None，这里输出空串。
        if (raw.Length == 0)
            return "";
        if (type == CellValues.SharedString)
        {
            return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture,
                       out var index)
                   && index >= 0 && index < sharedStrings.Count
                ? sharedStrings[index]
                : "";
        }
        if (type == CellValues.Boolean)
            return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var flag)
                ? flag != 0 ? "True" : "False"
                : raw;
        if (type == CellValues.Date)
            return FormatIsoCell(raw);
        if (type is not null && type != CellValues.Number)
            return raw;

        var styleIndex = (int)(cell.StyleIndex?.Value ?? 0);
        if (formats.IsTimedeltaStyle(styleIndex))
            return formats.FormatTimedeltaValue(raw);
        if (formats.IsDateStyle(styleIndex))
            return formats.FormatDateValue(raw);
        return ExcelReadFormats.FormatNumber(raw);
    }

    /// <summary>t="d"（ISO 8601 日期/时间）单元格，对齐 openpyxl from_ISO8601 的 str() 输出。</summary>
    private static string FormatIsoCell(string raw)
    {
        if (raw.Contains('T', StringComparison.Ordinal)
            || raw.Contains(' ', StringComparison.Ordinal))
        {
            if (DateTime.TryParse(raw, CultureInfo.InvariantCulture,
                    DateTimeStyles.RoundtripKind, out var dateTime))
            {
                var text = dateTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                var micro = (dateTime.Ticks % TimeSpan.TicksPerSecond) / 10;
                if (micro != 0)
                    text += "." + micro.ToString("D6", CultureInfo.InvariantCulture);
                return text;
            }
        }
        if (raw.Contains(':', StringComparison.Ordinal)
            && TimeOnly.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.None,
                out var time))
        {
            var text = time.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
            var micro = (time.Ticks % TimeSpan.TicksPerSecond) / 10;
            if (micro != 0)
                text += "." + micro.ToString("D6", CultureInfo.InvariantCulture);
            return text;
        }
        if (DateOnly.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
            return date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        return raw;
    }

    private static string BuildReadMarkdown(Dictionary<string, object?> result)
    {
        var parts = new List<string> { result["content"] as string ?? "" };
        parts.Add($"> sheets_count: {result["sheets_count"]}；total_rows: {result["total_rows"]}");
        if (result["truncated"] is true)
            parts.Add("> 内容过长已截断");
        return string.Join("\n\n", parts.Where(part => !string.IsNullOrEmpty(part)));
    }

    private static WorksheetPart? GetWorksheetPart(WorkbookPart workbookPart, Sheet sheet)
    {
        try
        {
            return string.IsNullOrEmpty(sheet.Id?.Value)
                ? null
                : workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
        }
        catch (KeyNotFoundException)
        {
            return null;
        }
    }

    private static int CellColumn(Cell cell)
    {
        if (ExcelA1.TryParseCell(cell.CellReference?.Value ?? "", out var column, out _))
            return column;
        return 0;
    }

    // ── 模板合并（xlsx 分支）─────────────────────────────────────

    public static Dictionary<string, object?> MergeTemplate(
        string templatePath,
        string outputPath,
        JsonElement data,
        string ifExists,
        bool confirmOverwrite)
    {
        var templateExtension = Path.GetExtension(templatePath).ToLowerInvariant();
        if (templateExtension is not (".xlsx" or ".xlsm"))
            throw new NativeRuntimeException(
                $"不支持的模板格式: {templateExtension}（xlsx 分支仅支持 .xlsx/.xlsm）");
        NativeCommon.EnsureFile(templatePath, templateExtension);
        var outputExtension = Path.GetExtension(outputPath).ToLowerInvariant();
        if (outputExtension != templateExtension)
            throw new NativeRuntimeException(
                $"输出文件必须与模板扩展名一致: {templateExtension}");
        ValidateOutputPath(outputPath, [templateExtension]);

        var target = ResolveOutputPath(outputPath, ifExists, confirmOverwrite);
        var fullTarget = Path.GetFullPath(target);
        Directory.CreateDirectory(Path.GetDirectoryName(fullTarget)!);
        var temp = TempPath(fullTarget);
        var values = FlattenTemplateData(data);
        var used = new HashSet<string>(StringComparer.Ordinal);
        var unresolved = new HashSet<string>(StringComparer.Ordinal);
        var replacedCount = 0;
        try
        {
            File.Copy(templatePath, temp, overwrite: true);
            ClearReadOnly(temp);
            using (var document = SpreadsheetDocument.Open(temp, true))
            {
                var workbookPart = document.WorkbookPart
                                   ?? throw new NativeRuntimeException("模板缺少 WorkbookPart");
                var sharedPart = workbookPart.SharedStringTablePart;
                var sharedTable = sharedPart?.SharedStringTable;
                var sharedItems = sharedTable?.Elements<SharedStringItem>().ToList() ?? [];
                foreach (var worksheetPart in workbookPart.WorksheetParts)
                {
                    var worksheet = worksheetPart.Worksheet;
                    if (worksheet is null)
                        continue;
                    foreach (var cell in worksheet.Descendants<Cell>())
                    {
                        string? original = null;
                        if (cell.DataType?.Value == CellValues.SharedString)
                        {
                            if (sharedTable is null)
                                continue;
                            if (!int.TryParse(cell.CellValue?.InnerText, NumberStyles.Integer,
                                    CultureInfo.InvariantCulture, out var index)
                                || index < 0 || index >= sharedItems.Count)
                                continue;
                            original = sharedItems[index].InnerText;
                        }
                        else if (cell.InlineString is not null)
                        {
                            original = cell.InlineString.InnerText;
                        }
                        else if (cell.DataType?.Value == CellValues.String && cell.CellFormula is null)
                        {
                            original = cell.CellValue?.InnerText;
                        }
                        if (original is null)
                            continue;

                        var changed = ReplacePlaceholders(
                            original, values, used, ref replacedCount);
                        if (!string.Equals(changed, original, StringComparison.Ordinal))
                        {
                            if (changed.StartsWith('='))
                            {
                                // 与 openpyxl 一致：合并结果以 = 开头时写成公式并清掉缓存值。
                                ApplyFormula(cell, changed);
                            }
                            else if (cell.DataType?.Value == CellValues.SharedString)
                            {
                                var newItem = new SharedStringItem(
                                    new Text(changed)
                                    {
                                        Space = SpaceProcessingModeValues.Preserve,
                                    });
                                // extLst 必须是 sharedStrings 的最后一个子节点，新 si 插到它之前。
                                var extLst = sharedTable!.ChildElements
                                    .FirstOrDefault(element => element.LocalName == "extLst");
                                if (extLst is not null)
                                    sharedTable.InsertBefore(newItem, extLst);
                                else
                                    sharedTable.Append(newItem);
                                sharedItems.Add(newItem);
                                var newIndex = sharedItems.Count - 1;
                                if (cell.CellValue is null)
                                    cell.CellValue = new CellValue(
                                        newIndex.ToString(CultureInfo.InvariantCulture));
                                else
                                    cell.CellValue.Text =
                                        newIndex.ToString(CultureInfo.InvariantCulture);
                            }
                            else if (cell.InlineString is not null)
                            {
                                cell.InlineString = new InlineString(
                                    new Text(changed)
                                    {
                                        Space = SpaceProcessingModeValues.Preserve,
                                    });
                            }
                            else if (cell.CellValue is not null)
                            {
                                cell.CellValue.Text = changed;
                            }
                        }
                        foreach (Match match in PlaceholderRegex().Matches(changed))
                            unresolved.Add(match.Groups[1].Value.Trim());
                    }
                }
                if (sharedTable is not null)
                {
                    // extLst 不是 si，计数只统计 SharedStringItem。
                    var itemCount = (uint)sharedTable.Elements<SharedStringItem>().Count();
                    sharedTable.UniqueCount = itemCount;
                    sharedTable.Count = itemCount;
                }
                if (workbookPart.CalculationChainPart is { } chain)
                    workbookPart.DeletePart(chain);
            }

            if (File.Exists(fullTarget) && new FileInfo(fullTarget).Length > 0)
                NativeCommon.Backup(fullTarget);
            ReplaceFile(temp, fullTarget);

            var usedKeys = used.ToList();
            usedKeys.Sort(StringComparer.Ordinal);
            var unresolvedKeys = unresolved.ToList();
            unresolvedKeys.Sort(StringComparer.Ordinal);
            return new Dictionary<string, object?>
            {
                ["success"] = true,
                ["template_path"] = templatePath,
                ["file_path"] = target,
                ["file_size"] = new FileInfo(fullTarget).Length,
                ["replaced_count"] = replacedCount,
                ["used_keys"] = usedKeys,
                ["unresolved_placeholders"] = unresolvedKeys,
                ["message"] = $"模板合并完成，替换 {replacedCount} 处，占位符未解析 {unresolvedKeys.Count} 个",
            };
        }
        finally
        {
            TryDelete(temp);
        }
    }

    private static void ApplyFormula(Cell cell, string text)
    {
        cell.CellValue = null;
        cell.InlineString = null;
        cell.DataType = null;
        cell.CellFormula = new CellFormula(text[1..]);
    }

    private static string ReplacePlaceholders(
        string text,
        Dictionary<string, string> values,
        HashSet<string> used,
        ref int replacedCount)
    {
        var count = 0;
        var changed = PlaceholderRegex().Replace(text, match =>
        {
            var key = match.Groups[1].Value.Trim();
            if (!values.TryGetValue(key, out var value))
                return match.Value;
            used.Add(key);
            count++;
            return value;
        });
        replacedCount += count;
        return changed;
    }

    [GeneratedRegex(@"\{\{\s*([^{}]+?)\s*\}\}", RegexOptions.CultureInvariant)]
    private static partial Regex PlaceholderRegex();

    private static Dictionary<string, string> FlattenTemplateData(JsonElement data)
    {
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var property in data.EnumerateObject())
            values[property.Name] = StringifyTemplateValue(property.Value);
        foreach (var property in data.EnumerateObject())
            VisitTemplateValue(property.Value, property.Name, values);
        return values;
    }

    private static void VisitTemplateValue(
        JsonElement value, string prefix, Dictionary<string, string> values)
    {
        if (value.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in value.EnumerateObject())
            {
                var path = $"{prefix}.{property.Name}";
                values.TryAdd(path, StringifyTemplateValue(property.Value));
                VisitTemplateValue(property.Value, path, values);
            }
        }
        else if (value.ValueKind == JsonValueKind.Array)
        {
            var index = 0;
            foreach (var child in value.EnumerateArray())
            {
                var path = $"{prefix}[{index}]";
                values.TryAdd(path, StringifyTemplateValue(child));
                VisitTemplateValue(child, path, values);
                index++;
            }
        }
    }

    private static string StringifyTemplateValue(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.Null or JsonValueKind.Undefined => "",
        JsonValueKind.String => value.GetString() ?? "",
        JsonValueKind.Number => value.GetRawText(),
        JsonValueKind.True => "True",
        JsonValueKind.False => "False",
        _ => PythonStyleJson(value),
    };

    private static string PythonStyleJson(JsonElement value)
    {
        switch (value.ValueKind)
        {
            case JsonValueKind.Object:
                return "{" + string.Join(", ", value.EnumerateObject().Select(property =>
                    JsonSerializer.Serialize(property.Name,
                        new JsonSerializerOptions
                        {
                            Encoder = System.Text.Encodings.Web.JavaScriptEncoder
                                .UnsafeRelaxedJsonEscaping,
                        })
                    + ": " + PythonStyleJson(property.Value))) + "}";
            case JsonValueKind.Array:
                return "[" + string.Join(", ", value.EnumerateArray()
                    .Select(PythonStyleJson)) + "]";
            case JsonValueKind.String:
                return JsonSerializer.Serialize(value.GetString(),
                    new JsonSerializerOptions
                    {
                        Encoder = System.Text.Encodings.Web.JavaScriptEncoder
                            .UnsafeRelaxedJsonEscaping,
                    });
            case JsonValueKind.Number:
                return value.GetRawText();
            case JsonValueKind.True:
                return "true";
            case JsonValueKind.False:
                return "false";
            default:
                return "null";
        }
    }

    // ── 路径与文件 ───────────────────────────────────────────────

    private static void ValidateOutputPath(string path, string[] extensions)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new NativeRuntimeException("输出路径不能为空");
        if (path.IndexOfAny(['<', '>', '"', '|', '?', '*']) >= 0)
            throw new NativeRuntimeException($"路径包含非法字符: {path}");
        var extension = Path.GetExtension(path).ToLowerInvariant();
        if (!extensions.Contains(extension))
            throw new NativeRuntimeException(
                $"输出文件必须使用以下扩展名之一: {string.Join("、", extensions)}");
    }

    internal static string ResolveOutputPath(string path, string ifExists, bool confirmOverwrite)
    {
        if (ifExists is not ("error" or "overwrite" or "rename"))
            throw new NativeUsageException(
                $"if_exists 必须是 error/overwrite/rename 之一，收到 {ifExists}");
        if (ifExists == "rename" && File.Exists(path) && new FileInfo(path).Length > 0)
        {
            var directory = Path.GetDirectoryName(Path.GetFullPath(path))!;
            var stem = Path.GetFileNameWithoutExtension(path);
            var extension = Path.GetExtension(path);
            for (var counter = 1; ; counter++)
            {
                var candidate = Path.Combine(directory, $"{stem} ({counter}){extension}");
                if (!File.Exists(candidate))
                    return candidate;
            }
        }
        if (File.Exists(path) && new FileInfo(path).Length > 0
            && !(ifExists == "overwrite" && confirmOverwrite))
        {
            throw new NativeRuntimeException(
                $"拒绝复写：目标已存在且非空 {path}；确需覆盖请传 overwrite=true 与 confirm_overwrite=true");
        }
        return path;
    }

    private static string TempPath(string target)
    {
        var directory = Path.GetDirectoryName(Path.GetFullPath(target))!;
        return Path.Combine(directory,
            $".{Path.GetFileNameWithoutExtension(target)}_{Guid.NewGuid():N}"
            + Path.GetExtension(target));
    }

    /// <summary>复制待编辑文件；被 Excel/WPS 占用时给出明确的替代入口提示。</summary>
    private static void CopyForEdit(string source, string temp)
    {
        try
        {
            File.Copy(source, temp, overwrite: true);
        }
        catch (Exception exception) when (
            exception is IOException or UnauthorizedAccessException)
        {
            throw LockedFileError(source, exception);
        }
    }

    private static NativeRuntimeException LockedFileError(string path, Exception exception) =>
        new($"无法写入 Excel 文件（可能已在 Excel/WPS 中打开）: {path}；"
            + "请改用 edit-open-excel-workbook 编辑活动工作簿，或关闭文件后重试。"
            + exception.Message);

    private static void ReplaceFile(string temp, string target)
    {
        if (File.Exists(target))
        {
            try
            {
                File.Replace(temp, target, null, true);
                return;
            }
            catch (PlatformNotSupportedException)
            {
                // 某些文件系统不支持原子替换，退回覆盖移动。
            }
        }
        File.Move(temp, target, true);
    }

    private static void ClearReadOnly(string path)
    {
        var attributes = File.GetAttributes(path);
        if ((attributes & FileAttributes.ReadOnly) != 0)
            File.SetAttributes(path, attributes & ~FileAttributes.ReadOnly);
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        catch (Exception exception) when (
            exception is IOException or UnauthorizedAccessException)
        {
            // 临时文件清理失败不影响结果。
        }
    }

    // ── 图表命令 ─────────────────────────────────────────────────

    public static Dictionary<string, object?> AddChart(
        string filePath, string? sheetName, ExcelChartDef chart)
    {
        NativeCommon.EnsureFile(filePath, OutputExtensions);
        var fullTarget = Path.GetFullPath(filePath);
        var temp = TempPath(fullTarget);
        string? backupPath = null;
        try
        {
            File.Copy(fullTarget, temp, overwrite: true);
            string targetSheetName;
            List<string> warnings;
            using (var document = SpreadsheetDocument.Open(temp, true))
            {
                var workbookPart = document.WorkbookPart
                                   ?? throw new NativeRuntimeException("工作簿缺少 WorkbookPart");
                var sheets = workbookPart.Workbook?.Sheets?.Elements<Sheet>().ToList() ?? [];
                if (sheets.Count == 0)
                    throw new NativeRuntimeException("工作簿没有可用的工作表");
                var names = new HashSet<string>(
                    sheets.Select(sheet => sheet.Name?.Value ?? ""), StringComparer.Ordinal);
                Sheet? target;
                if (sheetName is null)
                {
                    var activeTab = workbookPart.Workbook?.BookViews?
                        .Elements<WorkbookView>().FirstOrDefault()?.ActiveTab?.Value ?? 0;
                    target = activeTab >= 0 && activeTab < sheets.Count
                        ? sheets[(int)activeTab]
                        : sheets[0];
                }
                else
                {
                    target = sheets.FirstOrDefault(sheet =>
                        string.Equals(sheet.Name?.Value, sheetName, StringComparison.Ordinal));
                    if (target is null)
                        throw new NativeRuntimeException($"工作表不存在: {sheetName}");
                }
                targetSheetName = target.Name?.Value ?? "";
                var worksheetPart = GetWorksheetPart(workbookPart, target)
                                    ?? throw new NativeRuntimeException(
                                        $"找不到工作表部件: {targetSheetName}");
                warnings = ExcelFileCharts.AddChartToSheet(
                    worksheetPart, chart, targetSheetName, names, true);
            }

            backupPath = NativeCommon.Backup(fullTarget);
            ReplaceFile(temp, fullTarget);
            var data = new Dictionary<string, object?>
            {
                ["success"] = true,
                ["file_path"] = filePath,
                ["sheet"] = targetSheetName,
                ["chart_type"] = chart.ChartType,
                ["anchor"] = chart.Anchor,
                ["backup_path"] = backupPath,
            };
            if (warnings.Count > 0)
                data["warnings"] = warnings;
            return data;
        }
        finally
        {
            TryDelete(temp);
        }
    }
}

// ── 工作簿上下文 ─────────────────────────────────────────────────

internal sealed class ExcelWorkbookContext
{
    public SpreadsheetDocument Document { get; }
    public WorkbookPart WorkbookPart { get; }
    public Workbook Workbook { get; }
    public Sheets Sheets { get; }
    public ExcelStyleBook Styles { get; }
    public HashSet<string> SheetNames { get; } = new(StringComparer.Ordinal);
    public Dictionary<string, int> SheetIndexByName { get; } = new(StringComparer.Ordinal);
    public Dictionary<string, WorksheetPart> WorksheetPartsByName { get; } =
        new(StringComparer.Ordinal);
    public HashSet<string> TableNames { get; } = new(StringComparer.OrdinalIgnoreCase);
    public uint NextSheetId { get; private set; }
    public uint NextTableId { get; private set; }

    public uint AllocateTableId() => NextTableId++;

    public ExcelWorkbookContext(SpreadsheetDocument document)
    {
        Document = document;
        WorkbookPart = document.WorkbookPart
                       ?? throw new NativeRuntimeException("工作簿缺少 WorkbookPart");
        Workbook = WorkbookPart.Workbook
                   ?? throw new NativeRuntimeException("工作簿缺少 Workbook 根节点");
        Sheets = EnsureSheetsElement(Workbook);
        var stylesPart = WorkbookPart.WorkbookStylesPart
                         ?? WorkbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet ??= ExcelStyleBook.CreateDefaultStylesheet();
        Styles = new ExcelStyleBook(stylesPart.Stylesheet);
        RefreshSheetIndex();
        NextSheetId = Sheets.Elements<Sheet>()
            .Select(sheet => sheet.SheetId?.Value ?? 0U)
            .DefaultIfEmpty(0U)
            .Max() + 1;
        LoadTableState();
    }

    private static Sheets EnsureSheetsElement(Workbook workbook)
    {
        if (workbook.Sheets is { } existing)
            return existing;
        var sheets = new Sheets();
        var anchor = workbook.ChildElements.FirstOrDefault(element =>
            element.LocalName is "functionGroups" or "externalReferences" or "definedNames"
                or "calcPr" or "oleSize" or "customWorkbookViews" or "pivotCaches"
                or "webPublishing" or "extLst");
        if (anchor is not null)
            workbook.InsertBefore(sheets, anchor);
        else
            workbook.Append(sheets);
        return sheets;
    }

    public void RefreshSheetIndex()
    {
        SheetNames.Clear();
        SheetIndexByName.Clear();
        WorksheetPartsByName.Clear();
        var index = 0;
        foreach (var sheet in Sheets.Elements<Sheet>())
        {
            var name = sheet.Name?.Value ?? "";
            SheetNames.Add(name);
            SheetIndexByName[name] = index;
            if (sheet.Id?.Value is { } id)
            {
                try
                {
                    if (WorkbookPart.GetPartById(id) is WorksheetPart worksheetPart)
                        WorksheetPartsByName[name] = worksheetPart;
                }
                catch (KeyNotFoundException)
                {
                    // 损坏的工作簿中可能残留无效关系；跳过即可。
                }
            }
            index++;
        }
    }

    private void LoadTableState()
    {
        TableNames.Clear();
        var maxId = 0U;
        foreach (var worksheetPart in WorkbookPart.WorksheetParts)
        {
            foreach (var tablePart in worksheetPart.TableDefinitionParts)
            {
                var table = tablePart.Table;
                if (table?.Name?.Value is { } name)
                    TableNames.Add(name);
                if (table?.Id?.Value is { } id && id > maxId)
                    maxId = id;
            }
        }
        NextTableId = maxId + 1;
    }

    public List<(ExcelSheetDef Def, string Name, WorksheetPart Part, Sheet Sheet)> AddSheets(
        List<ExcelSheetDef> definitions)
    {
        var created = new List<(ExcelSheetDef, string, WorksheetPart, Sheet)>();
        foreach (var definition in definitions)
        {
            var name = NormalizeSheetName(definition.Name);
            var part = WorkbookPart.AddNewPart<WorksheetPart>();
            part.Worksheet = new Worksheet();
            var sheet = new Sheet
            {
                Name = name,
                SheetId = NextSheetId++,
                Id = WorkbookPart.GetIdOfPart(part),
            };
            sheet.State = definition.SheetState switch
            {
                "hidden" => SheetStateValues.Hidden,
                "veryHidden" => SheetStateValues.VeryHidden,
                _ => SheetStateValues.Visible,
            };
            Sheets.Append(sheet);
            SheetNames.Add(name);
            created.Add((definition, name, part, sheet));
        }

        RefreshSheetIndex();
        foreach (var (definition, name, part, _) in created)
        {
            var writer = new ExcelSheetWriter(this, definition, name, part);
            writer.Write();
        }
        EnsureVisibleSheet();
        return created;
    }

    private string NormalizeSheetName(string raw)
    {
        var name = raw.Length > 31 ? raw[..31] : raw;
        if (string.IsNullOrWhiteSpace(name))
            throw new NativeRuntimeException("工作表名称不能为空");
        if (name.IndexOfAny(['[', ']', ':', '*', '?', '/', '\\']) >= 0)
            throw new NativeRuntimeException($"工作表名称包含非法字符: {raw}");
        if (SheetNames.Contains(name)
            || SheetNames.Any(existing =>
                string.Equals(existing, name, StringComparison.OrdinalIgnoreCase)))
            throw new NativeRuntimeException($"工作表已存在: {name}");
        return name;
    }

    private void EnsureVisibleSheet()
    {
        var hasVisible = Sheets.Elements<Sheet>().Any(sheet =>
            sheet.State is null || sheet.State.Value == SheetStateValues.Visible);
        if (!hasVisible)
            throw new NativeRuntimeException("工作簿至少需要一个可见工作表");
    }

    public void DeleteAllSheets()
    {
        var visited = new HashSet<OpenXmlPart>();
        foreach (var sheet in Sheets.Elements<Sheet>().ToList())
        {
            if (sheet.Id?.Value is { } id)
            {
                try
                {
                    var part = WorkbookPart.GetPartById(id);
                    DeletePartTree(WorkbookPart, part, visited);
                }
                catch (KeyNotFoundException)
                {
                    // 关系已丢失，直接移除 sheet 记录。
                }
            }
            sheet.Remove();
        }
        NextSheetId = 1;
        NextTableId = 1;
        TableNames.Clear();

        if (WorkbookPart.SharedStringTablePart is { } sharedPart)
        {
            sharedPart.SharedStringTable = new SharedStringTable
            {
                Count = 0,
                UniqueCount = 0,
            };
        }
        if (WorkbookPart.CalculationChainPart is { } chain)
            WorkbookPart.DeletePart(chain);

        RemoveDefinedNamesForSheets();
        RefreshSheetIndex();
    }

    private static void DeletePartTree(
        OpenXmlPartContainer parent, OpenXmlPart part, HashSet<OpenXmlPart> visited)
    {
        if (!visited.Add(part))
        {
            try
            {
                parent.DeletePart(part);
            }
            catch (Exception exception) when (
                exception is ArgumentOutOfRangeException or KeyNotFoundException
                    or InvalidOperationException)
            {
                // 关系已在另一棵子树中移除。
            }
            return;
        }
        foreach (var child in part.Parts.Select(pair => pair.OpenXmlPart).ToList())
            DeletePartTree(part, child, visited);
        parent.DeletePart(part);
    }

    /// <summary>
    /// 删除所有工作表后，本地定义名的 LocalSheetId 已全部失效；按 LocalSheetId 清理，
    /// 不做名称文本子串匹配（工作簿级名称保持原样，与 openpyxl 一致）。
    /// </summary>
    private void RemoveDefinedNamesForSheets()
    {
        if (Workbook.DefinedNames is not { } definedNames)
            return;
        foreach (var definedName in definedNames.Elements<DefinedName>().ToList())
        {
            if (definedName.LocalSheetId is not null)
                definedName.Remove();
        }
        if (!definedNames.Elements<DefinedName>().Any())
            definedNames.Remove();
    }

    /// <summary>把 bookViews 的 activeTab / firstSheet 重算到第一个可见工作表。</summary>
    public void NormalizeWorkbookViews()
    {
        if (Workbook.BookViews is not { } bookViews)
            return;
        var sheets = Sheets.Elements<Sheet>().ToList();
        if (sheets.Count == 0)
            return;
        var firstVisible = sheets.FindIndex(sheet =>
            sheet.State is null || sheet.State.Value == SheetStateValues.Visible);
        if (firstVisible < 0)
            firstVisible = 0;
        foreach (var view in bookViews.Elements<WorkbookView>())
        {
            view.ActiveTab = (uint)firstVisible;
            view.FirstSheet = (uint)firstVisible;
        }
    }

    public DefinedNames EnsureDefinedNames()
    {
        if (Workbook.DefinedNames is { } existing)
            return existing;
        var definedNames = new DefinedNames();
        var anchor = Workbook.ChildElements.FirstOrDefault(element =>
            element.LocalName is "calcPr" or "oleSize" or "customWorkbookViews"
                or "pivotCaches" or "webPublishing" or "extLst");
        if (anchor is not null)
            Workbook.InsertBefore(definedNames, anchor);
        else
            Workbook.Append(definedNames);
        return definedNames;
    }

    public bool HasDefinedName(string name) =>
        Workbook.DefinedNames?.Elements<DefinedName>().Any(definedName =>
            string.Equals(definedName.Name?.Value, name, StringComparison.OrdinalIgnoreCase))
        == true;

    public void AddDefinedName(
        string name, string refersTo, int? localSheetId, bool replace)
    {
        var text = refersTo.Trim();
        if (text.StartsWith('='))
            text = text[1..].Trim();
        if (text.Length == 0)
            throw new NativeRuntimeException($"定义名称 {name} 的 refers_to 不能为空");
        if (TableNames.Contains(name))
            throw new NativeRuntimeException($"定义名称 {name} 与超级表名称冲突");

        var container = EnsureDefinedNames();
        var existing = container.Elements<DefinedName>().FirstOrDefault(definedName =>
            string.Equals(definedName.Name?.Value, name, StringComparison.OrdinalIgnoreCase)
            && definedName.LocalSheetId?.Value == (uint?)localSheetId);
        var scope = localSheetId is null
            ? "工作簿"
            : $"工作表 {Sheets.Elements<Sheet>().ElementAtOrDefault(localSheetId.Value)?.Name?.Value}";
        if (existing is not null)
        {
            if (!replace)
                throw new NativeRuntimeException($"{scope}中已存在定义名称: {name}");
            existing.Text = text;
            return;
        }
        var definedName = new DefinedName { Name = name, Text = text };
        if (localSheetId is not null)
            definedName.LocalSheetId = (uint)localSheetId.Value;
        container.Append(definedName);
    }

    public void Save()
    {
        Styles.RefreshCounts();
        EnsureCalculationProperties();
        if (WorkbookPart.CalculationChainPart is { } chain)
            WorkbookPart.DeletePart(chain);
        Workbook.Save();
    }

    private void EnsureCalculationProperties()
    {
        var calculation = Workbook.CalculationProperties;
        if (calculation is null)
        {
            calculation = new CalculationProperties
            {
                CalculationId = 124519,
                FullCalculationOnLoad = true,
            };
            var anchor = Workbook.ChildElements.FirstOrDefault(element =>
                element.LocalName is "oleSize" or "customWorkbookViews"
                    or "pivotCaches" or "webPublishing" or "extLst");
            if (anchor is not null)
                Workbook.InsertBefore(calculation, anchor);
            else
                Workbook.Append(calculation);
        }
        else
        {
            calculation.FullCalculationOnLoad = true;
            calculation.CalculationId ??= 124519;
        }
    }
}

// ── 工作表写入 ───────────────────────────────────────────────────

internal sealed class ExcelSheetWriter
{
    private const int MaxColumns = ExcelFileOoxml.MaxColumns;
    private const int MaxRows = ExcelFileOoxml.MaxRows;

    private readonly ExcelWorkbookContext _context;
    private readonly ExcelSheetDef _def;
    private readonly string _name;
    private readonly WorksheetPart _part;
    private readonly ExcelStyleBook _styles;
    private ExcelSortPlan? _sortPlan;
    private List<List<JsonElement>> _dataRows = [];

    public ExcelSheetWriter(
        ExcelWorkbookContext context, ExcelSheetDef definition, string name, WorksheetPart part)
    {
        _context = context;
        _def = definition;
        _name = name;
        _part = part;
        _styles = context.Styles;
    }

    public void Write()
    {
        ValidateGeometry();
        // 排序在写行之前完成：计划失败时整张表都不写入，避免留下半成品。
        _sortPlan = ExcelSortPlanner.Prepare(_def);
        _dataRows = _sortPlan is null ? _def.Data : _sortPlan.Rows;
        var worksheet = _part.Worksheet ??= new Worksheet();
        worksheet.RemoveAllChildren();

        var rows = new SortedDictionary<int, Row>();
        var cells = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
        var textValues = new Dictionary<(int Row, int Col), string>();

        Row GetRow(int rowIndex)
        {
            if (!rows.TryGetValue(rowIndex, out var row))
            {
                row = new Row { RowIndex = (uint)rowIndex };
                rows[rowIndex] = row;
            }
            return row;
        }

        Cell GetCell(int column, int rowIndex, int styleIndex)
        {
            var reference = ExcelA1.CellRef(column, rowIndex);
            if (cells.TryGetValue(reference, out var existing))
                return existing;
            var cell = new Cell { CellReference = reference };
            if (styleIndex > 0)
                cell.StyleIndex = (uint)styleIndex;
            GetRow(rowIndex).Append(cell);
            cells[reference] = cell;
            return cell;
        }

        var columnStyles = new Dictionary<int, int>();
        if (_def.ColumnFormats is not null)
        {
            for (var index = 0; index < _def.ColumnFormats.Count; index++)
                columnStyles[index + 1] = _styles.GetStyle(_def.ColumnFormats[index].ToCellStyle());
        }

        if (_def.Headers is { Count: > 0 })
        {
            var headerStyle = _def.HeaderStyle is not null
                ? _styles.GetStyle(_def.HeaderStyle)
                : _styles.GetBoldHeaderStyle();
            for (var index = 0; index < _def.Headers.Count; index++)
            {
                var column = index + 1;
                var cell = GetCell(column, 1, headerStyle);
                SetInlineString(cell, _def.Headers[index]);
                textValues[(1, column)] = _def.Headers[index];
            }
        }

        var startRow = _def.Headers is { Count: > 0 } ? 2 : 1;
        for (var rowOffset = 0; rowOffset < _dataRows.Count; rowOffset++)
        {
            var rowIndex = startRow + rowOffset;
            var rowData = _dataRows[rowOffset];
            for (var columnOffset = 0; columnOffset < rowData.Count; columnOffset++)
            {
                var column = columnOffset + 1;
                var styleIndex = columnStyles.TryGetValue(column, out var style) ? style : 0;
                var cell = GetCell(column, rowIndex, styleIndex);
                WriteCellValue(cell, rowData[columnOffset], _def.AllowFormulas);
                if (PlainText(rowData[columnOffset]) is { } text)
                    textValues[(rowIndex, column)] = text;
            }
        }

        if (_def.RowHeights is not null)
        {
            foreach (var (rowIndex, height) in _def.RowHeights)
            {
                if (rowIndex < 1)
                    continue;
                var row = GetRow(rowIndex);
                row.Height = height;
                row.CustomHeight = true;
            }
        }

        var columns = BuildColumns();
        var hyperlinks = BuildHyperlinks(GetCell, cells);
        var comments = BuildComments(GetCell);
        var conditionalFormats = BuildConditionalFormats();
        var dataValidations = BuildDataValidations();
        var tableParts = BuildTables(textValues);
        var pageSetup = BuildPageSetup();
        BuildDefinedNames();
        var sheetProperties = BuildSheetProperties();
        var sheetViews = BuildSheetViews();
        var mergedCells = BuildMergedCells();

        string? drawingRelId = null;
        if (_def.Images is not null)
        {
            foreach (var image in _def.Images)
            {
                drawingRelId ??= ExcelFileDrawing.EnsureDrawingsPart(_part);
                ExcelFileDrawing.AddImage(_part, image);
            }
        }
        if (_def.Charts is not null)
        {
            foreach (var chart in _def.Charts)
            {
                drawingRelId ??= ExcelFileDrawing.EnsureDrawingsPart(_part);
                ExcelFileCharts.AddChartToSheet(
                    _part, chart, _name, _context.SheetNames, false);
            }
        }
        string? legacyDrawingRelId = null;
        if (comments is not null)
            legacyDrawingRelId = ExcelFileDrawing.AddCommentsVml(_part, _def.Comments!);

        // 把可能乱序插入的单元格按列排序，Excel 的 sheetData 惯例按列递增。
        foreach (var row in rows.Values)
        {
            var ordered = row.Elements<Cell>()
                .OrderBy(cell => ExcelA1.TryParseCell(
                    cell.CellReference?.Value ?? "", out var column, out _) ? column : 0)
                .ToList();
            row.RemoveAllChildren<Cell>();
            foreach (var cell in ordered)
                row.Append(cell);
        }

        var dimension = BuildDimension(cells);
        var sheetData = new SheetData(rows.Values);

        if (sheetProperties is not null)
            worksheet.Append(sheetProperties);
        worksheet.Append(dimension);
        if (sheetViews is not null)
            worksheet.Append(sheetViews);
        worksheet.Append(new SheetFormatProperties
        {
            DefaultRowHeight = 15,
            BaseColumnWidth = 8,
        });
        if (columns is not null)
            worksheet.Append(columns);
        worksheet.Append(sheetData);
        // CT_Worksheet 顺序：sheetProtection 位于 sheetData 之后、autoFilter/mergeCells 之前。
        if (_def.Protection is { } protection)
            worksheet.Append(ExcelSheetProtectionBuilder.Build(protection));
        if (BuildAutoFilter() is { } autoFilter)
            worksheet.Append(autoFilter);
        if (mergedCells is not null)
            worksheet.Append(mergedCells);
        foreach (var conditionalFormat in conditionalFormats)
            worksheet.Append(conditionalFormat);
        if (dataValidations is not null)
            worksheet.Append(dataValidations);
        if (hyperlinks is not null)
            worksheet.Append(hyperlinks);
        if (pageSetup.PrintOptions is not null)
            worksheet.Append(pageSetup.PrintOptions);
        if (pageSetup.Margins is not null)
            worksheet.Append(pageSetup.Margins);
        if (pageSetup.PageSetupElement is not null)
            worksheet.Append(pageSetup.PageSetupElement);
        if (drawingRelId is not null)
            worksheet.Append(new Drawing { Id = drawingRelId });
        if (legacyDrawingRelId is not null)
            worksheet.Append(new LegacyDrawing { Id = legacyDrawingRelId });
        if (tableParts is not null)
            worksheet.Append(tableParts);
    }

    // ── 单元格 ───────────────────────────────────────────────────

    /// <summary>Excel 工作表几何上限：1..16384 列、1..1048576 行。</summary>
    private void ValidateGeometry()
    {
        var sheetLabel = $"工作表 {_name}";
        if (_def.Headers is { Count: > MaxColumns } headers)
            throw new NativeRuntimeException(
                $"{sheetLabel} 表头超过 {MaxColumns} 列上限: {headers.Count} 列");
        if (_def.ColumnWidths is { Count: > MaxColumns } widths)
            throw new NativeRuntimeException(
                $"{sheetLabel} column_widths 超过 {MaxColumns} 列上限: {widths.Count} 项");
        if (_def.ColumnFormats is { Count: > MaxColumns } formats)
            throw new NativeRuntimeException(
                $"{sheetLabel} column_formats 超过 {MaxColumns} 列上限: {formats.Count} 项");
        if (_def.RowHeights is not null)
        {
            foreach (var row in _def.RowHeights.Keys)
            {
                if (row > MaxRows)
                    throw new NativeRuntimeException(
                        $"{sheetLabel} row_heights 行号超过 {MaxRows} 行上限: {row}");
            }
        }

        var startRow = _def.Headers is { Count: > 0 } ? 2 : 1;
        if (_def.Data.Count > 0 && startRow + _def.Data.Count - 1 > MaxRows)
        {
            throw new NativeRuntimeException(
                $"{sheetLabel} 数据行超出 {MaxRows} 行上限: "
                + $"写入到第 {startRow + _def.Data.Count - 1} 行");
        }
        for (var index = 0; index < _def.Data.Count; index++)
        {
            if (_def.Data[index].Count > MaxColumns)
                throw new NativeRuntimeException(
                    $"{sheetLabel} 第 {startRow + index} 行数据超过 {MaxColumns} 列上限: "
                    + $"{_def.Data[index].Count} 列");
        }
    }

    private static void WriteCellValue(Cell cell, JsonElement value, bool allowFormulas)
    {
        switch (value.ValueKind)
        {
            case JsonValueKind.Null:
            case JsonValueKind.Undefined:
                return;
            case JsonValueKind.True:
            case JsonValueKind.False:
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new CellValue(value.GetBoolean() ? "1" : "0");
                return;
            case JsonValueKind.Number:
                cell.CellValue = new CellValue(value.GetRawText());
                return;
            case JsonValueKind.String:
                var text = value.GetString() ?? "";
                if (allowFormulas && text.StartsWith('='))
                {
                    cell.CellFormula = new CellFormula(text[1..]);
                    return;
                }
                SetInlineString(cell, text.StartsWith('=') ? "'" + text : text);
                return;
            default:
                throw new NativeUsageException("单元格数据只支持字符串、数字、布尔值与 null");
        }
    }

    private static void SetInlineString(Cell cell, string text)
    {
        cell.DataType = CellValues.InlineString;
        cell.InlineString = new InlineString(new Text(text)
        {
            Space = SpaceProcessingModeValues.Preserve,
        });
    }

    private static string? PlainText(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.Null or JsonValueKind.Undefined => null,
        JsonValueKind.String => value.GetString(),
        JsonValueKind.Number => value.GetRawText(),
        JsonValueKind.True => "True",
        JsonValueKind.False => "False",
        _ => null,
    };

    // ── 列 / 视图 / 合并 ─────────────────────────────────────────

    private Columns? BuildColumns()
    {
        var widths = new SortedDictionary<int, double>();
        if (_def.ColumnWidths is not null)
        {
            for (var index = 0; index < _def.ColumnWidths.Count; index++)
                widths[index + 1] = _def.ColumnWidths[index];
        }
        if (_def.ColumnFormats is not null)
        {
            for (var index = 0; index < _def.ColumnFormats.Count; index++)
            {
                if (_def.ColumnFormats[index].Width is { } width)
                    widths[index + 1] = width;
            }
        }
        if (widths.Count == 0)
            return null;
        var columns = new Columns();
        foreach (var (index, width) in widths)
        {
            if (width < 0)
                throw new NativeRuntimeException($"列宽不能为负: 第 {index} 列");
            columns.Append(new Column
            {
                Min = (uint)index,
                Max = (uint)index,
                Width = width,
                CustomWidth = true,
            });
        }
        return columns;
    }

    private SheetViews? BuildSheetViews()
    {
        var sheetView = new SheetView { WorkbookViewId = 0 };
        if (!string.IsNullOrWhiteSpace(_def.Freeze))
        {
            if (!ExcelA1.TryParseCell(_def.Freeze, out var column, out var row))
                throw new NativeUsageException($"无效的冻结单元格: {_def.Freeze}");
            var xSplit = column - 1;
            var ySplit = row - 1;
            if (xSplit > 0 || ySplit > 0)
            {
                var topLeft = ExcelA1.CellRef(column, row);
                var pane = new Pane
                {
                    TopLeftCell = topLeft,
                    State = PaneStateValues.Frozen,
                };
                if (xSplit > 0)
                    pane.HorizontalSplit = xSplit;
                if (ySplit > 0)
                    pane.VerticalSplit = ySplit;
                var paneValue = xSplit > 0 && ySplit > 0
                    ? PaneValues.BottomRight
                    : ySplit > 0
                        ? PaneValues.BottomLeft
                        : PaneValues.TopRight;
                pane.ActivePane = paneValue;
                sheetView.Append(pane);
                sheetView.Append(new Selection
                {
                    Pane = paneValue,
                    ActiveCell = topLeft,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = topLeft },
                });
            }
        }
        if (sheetView.ChildElements.Count == 0)
        {
            sheetView.Append(new Selection
            {
                ActiveCell = "A1",
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" },
            });
        }
        return new SheetViews(sheetView);
    }

    private MergeCells? BuildMergedCells()
    {
        if (_def.MergedCells is null || _def.MergedCells.Count == 0)
            return null;
        var merges = new MergeCells { Count = (uint)_def.MergedCells.Count };
        foreach (var reference in _def.MergedCells)
            merges.Append(new MergeCell { Reference = reference });
        return merges;
    }

    private AutoFilter? BuildAutoFilter()
    {
        if (!_def.AutoFilter || _def.Headers is not { Count: > 0 } || _dataRows.Count == 0)
            return null;
        var lastRow = _dataRows.Count + 1;
        var lastColumn = _def.Headers.Count;
        // Excel 不允许工作表级自动筛选与超级表重叠；表自身已带筛选按钮。
        if (_def.Tables is not null)
        {
            foreach (var table in _def.Tables)
            {
                if (!ExcelA1.TryParseRange(table.Range, out var tableRange))
                    continue;
                if (tableRange.MinCol <= lastColumn && tableRange.MaxCol >= 1
                    && tableRange.MinRow <= lastRow && tableRange.MaxRow >= 1)
                    return null;
            }
        }
        var autoFilter = new AutoFilter
        {
            Reference = $"A1:{ExcelA1.ColumnName(lastColumn)}{lastRow}",
        };
        // 排序只写进 autoFilter 的 sortState，不额外造 autoFilter 元素。
        if (_sortPlan is not null)
            autoFilter.Append(ExcelSortPlanner.BuildSortState(_sortPlan, _def, autoFilter.Reference!));
        return autoFilter;
    }

    private SheetDimension BuildDimension(Dictionary<string, Cell> cells)
    {
        var maxColumn = 0;
        var maxRow = 0;
        foreach (var reference in cells.Keys)
        {
            if (!ExcelA1.TryParseCell(reference, out var column, out var row))
                continue;
            if (column > maxColumn)
                maxColumn = column;
            if (row > maxRow)
                maxRow = row;
        }
        return new SheetDimension
        {
            Reference = maxColumn > 0 && maxRow > 0
                ? ExcelA1.FormatRange(1, 1, maxColumn, maxRow)
                : "A1:A1",
        };
    }

    private SheetProperties? BuildSheetProperties()
    {
        var properties = new SheetProperties();
        if (_def.TabColor is not null)
            properties.Append(new TabColor { Rgb = _def.TabColor });
        if (_def.PageSetup is { } pageSetup
            && (pageSetup.FitToWidth is not null || pageSetup.FitToHeight is not null))
        {
            properties.Append(new PageSetupProperties { FitToPage = true });
        }
        return properties.ChildElements.Count == 0 ? null : properties;
    }

    // ── 超链接 / 批注 ────────────────────────────────────────────

    private Hyperlinks? BuildHyperlinks(
        Func<int, int, int, Cell> getCell, Dictionary<string, Cell> cells)
    {
        if (_def.Hyperlinks is null || _def.Hyperlinks.Count == 0)
            return null;
        var hyperlinks = new Hyperlinks();
        foreach (var definition in _def.Hyperlinks)
        {
            if (!ExcelA1.TryParseCell(definition.Cell, out var column, out var row))
                throw new NativeUsageException($"无效的超链接单元格: {definition.Cell}");
            var reference = ExcelA1.CellRef(column, row);
            var cell = cells.TryGetValue(reference, out var existing)
                ? existing
                : getCell(column, row, 0);
            var uri = MakeHyperlinkUri(definition.Url);
            var relationship = _part.AddHyperlinkRelationship(uri, true);
            hyperlinks.Append(new Hyperlink
            {
                Reference = reference,
                Id = relationship.Id,
                Tooltip = definition.Tooltip,
            });
            var styleIndex = _styles.GetHyperlinkStyle((int)(cell.StyleIndex?.Value ?? 0));
            cell.StyleIndex = (uint)styleIndex;
        }
        return hyperlinks;
    }

    private static Uri MakeHyperlinkUri(string url)
    {
        if (Uri.TryCreate(url, UriKind.RelativeOrAbsolute, out var uri))
            return uri;
        if (Uri.TryCreate(url.Replace('\\', '/'), UriKind.RelativeOrAbsolute, out uri))
            return uri;
        throw new NativeUsageException($"无效的超链接地址: {url}");
    }

    private Comments? BuildComments(Func<int, int, int, Cell> getCell)
    {
        if (_def.Comments is null || _def.Comments.Count == 0)
            return null;
        var authors = new List<string>();
        var authorIds = new Dictionary<string, int>(StringComparer.Ordinal);
        var commentList = new CommentList();
        foreach (var definition in _def.Comments)
        {
            if (!ExcelA1.TryParseCell(definition.Cell, out var column, out var row))
                throw new NativeUsageException($"无效的批注单元格: {definition.Cell}");
            if (!authorIds.TryGetValue(definition.Author, out var authorId))
            {
                authorId = authors.Count;
                authors.Add(definition.Author);
                authorIds[definition.Author] = authorId;
            }
            getCell(column, row, 0);
            commentList.Append(new Comment
            {
                Reference = ExcelA1.CellRef(column, row),
                AuthorId = (uint)authorId,
                CommentText = new CommentText(new Text(definition.Text)
                {
                    Space = SpaceProcessingModeValues.Preserve,
                }),
            });
        }
        var commentsPart = _part.AddNewPart<WorksheetCommentsPart>();
        var comments = new Comments(
            new Authors(authors.Select(author => new Author { Text = author })),
            commentList);
        commentsPart.Comments = comments;
        return comments;
    }

    // ── 条件格式 / 数据验证 ──────────────────────────────────────

    private List<ConditionalFormatting> BuildConditionalFormats()
    {
        var result = new List<ConditionalFormatting>();
        if (_def.ConditionalFormats is null)
            return result;
        var priority = 1;
        foreach (var definition in _def.ConditionalFormats)
        {
            if (!ExcelA1.TryParseRange(definition.Range, out var range))
                throw new NativeUsageException($"无效的条件格式范围: {definition.Range}");
            var rangeText = ExcelA1.FormatRange(
                range.MinCol, range.MinRow, range.MaxCol, range.MaxRow);
            var rule = new ConditionalFormattingRule { Priority = priority++ };
            if (definition.FillColor is not null || definition.FontColor is not null)
            {
                rule.FormatId = (uint)_styles.AddDifferentialFormat(
                    definition.FillColor, definition.FontColor);
            }

            switch (definition.Type)
            {
                case "cell_is":
                    rule.Type = ConditionalFormatValues.CellIs;
                    rule.Operator = definition.Operator switch
                    {
                        "greaterThan" => ConditionalFormattingOperatorValues.GreaterThan,
                        "lessThan" => ConditionalFormattingOperatorValues.LessThan,
                        "between" => ConditionalFormattingOperatorValues.Between,
                        "equal" => ConditionalFormattingOperatorValues.Equal,
                        "notEqual" => ConditionalFormattingOperatorValues.NotEqual,
                        "greaterThanOrEqual" =>
                            ConditionalFormattingOperatorValues.GreaterThanOrEqual,
                        "lessThanOrEqual" =>
                            ConditionalFormattingOperatorValues.LessThanOrEqual,
                        _ => throw new NativeUsageException(
                            $"不支持的条件格式运算符: {definition.Operator}"),
                    };
                    foreach (var formula in definition.Formula ?? [])
                        rule.Append(new Formula(formula));
                    break;
                case "formula":
                    rule.Type = ConditionalFormatValues.Expression;
                    foreach (var formula in definition.Formula ?? [])
                        rule.Append(new Formula(formula));
                    break;
                case "color_scale":
                    rule.Type = ConditionalFormatValues.ColorScale;
                    var colorScale = new ColorScale();
                    colorScale.Append(new ConditionalFormatValueObject
                    {
                        Type = ConditionalFormatValueObjectValues.Min,
                    });
                    if (definition.MidColor is not null)
                    {
                        colorScale.Append(new ConditionalFormatValueObject
                        {
                            Type = ConditionalFormatValueObjectValues.Percentile,
                            Val = "50",
                        });
                    }
                    colorScale.Append(new ConditionalFormatValueObject
                    {
                        Type = ConditionalFormatValueObjectValues.Max,
                    });
                    colorScale.Append(new Color
                    {
                        Rgb = definition.MinColor ?? "FFF8696B",
                    });
                    if (definition.MidColor is not null)
                        colorScale.Append(new Color { Rgb = definition.MidColor });
                    colorScale.Append(new Color
                    {
                        Rgb = definition.MaxColor ?? "FF63BE7B",
                    });
                    rule.Append(colorScale);
                    break;
                case "data_bar":
                    rule.Type = ConditionalFormatValues.DataBar;
                    rule.Append(new DataBar(
                        new ConditionalFormatValueObject
                        {
                            Type = ConditionalFormatValueObjectValues.Min,
                        },
                        new ConditionalFormatValueObject
                        {
                            Type = ConditionalFormatValueObjectValues.Max,
                        },
                        new Color { Rgb = definition.BarColor ?? "FF638EC6" })
                    {
                        ShowValue = true,
                    });
                    break;
                case "icon_set":
                    rule.Type = ConditionalFormatValues.IconSet;
                    var (iconSet, thresholds) = IconSetFor(
                        definition.IconStyle ?? "3TrafficLights1");
                    var iconElement = new IconSet
                    {
                        IconSetValue = iconSet,
                        ShowValue = true,
                    };
                    foreach (var threshold in thresholds)
                    {
                        iconElement.Append(new ConditionalFormatValueObject
                        {
                            Type = ConditionalFormatValueObjectValues.Percent,
                            Val = threshold.ToString(CultureInfo.InvariantCulture),
                        });
                    }
                    rule.Append(iconElement);
                    break;
                default:
                    throw new NativeUsageException(
                        $"不支持的条件格式类型: {definition.Type}");
            }

            var conditionalFormatting = new ConditionalFormatting
            {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = rangeText },
            };
            conditionalFormatting.Append(rule);
            result.Add(conditionalFormatting);
        }
        return result;
    }

    private static (IconSetValues Value, int[] Thresholds) IconSetFor(string name) =>
        name switch
        {
            "3Arrows" => (IconSetValues.ThreeArrows, [0, 33, 67]),
            "3ArrowsGray" => (IconSetValues.ThreeArrowsGray, [0, 33, 67]),
            "3Flags" => (IconSetValues.ThreeFlags, [0, 33, 67]),
            "3TrafficLights1" => (IconSetValues.ThreeTrafficLights1, [0, 33, 67]),
            "3TrafficLights2" => (IconSetValues.ThreeTrafficLights2, [0, 33, 67]),
            "3Signs" => (IconSetValues.ThreeSigns, [0, 33, 67]),
            "3Symbols" => (IconSetValues.ThreeSymbols, [0, 33, 67]),
            "3Symbols2" => (IconSetValues.ThreeSymbols2, [0, 33, 67]),
            "4Arrows" => (IconSetValues.FourArrows, [0, 25, 50, 75]),
            "4ArrowsGray" => (IconSetValues.FourArrowsGray, [0, 25, 50, 75]),
            "4RedToBlack" => (IconSetValues.FourRedToBlack, [0, 25, 50, 75]),
            "4Rating" => (IconSetValues.FourRating, [0, 25, 50, 75]),
            "4TrafficLights" => (IconSetValues.FourTrafficLights, [0, 25, 50, 75]),
            "5Arrows" => (IconSetValues.FiveArrows, [0, 20, 40, 60, 80]),
            "5ArrowsGray" => (IconSetValues.FiveArrowsGray, [0, 20, 40, 60, 80]),
            "5Rating" => (IconSetValues.FiveRating, [0, 20, 40, 60, 80]),
            "5Quarters" => (IconSetValues.FiveQuarters, [0, 20, 40, 60, 80]),
            _ => throw new NativeUsageException($"不支持的图标集: {name}"),
        };

    private DataValidations? BuildDataValidations()
    {
        if (_def.DataValidations is null || _def.DataValidations.Count == 0)
            return null;
        var validations = new DataValidations
        {
            Count = (uint)_def.DataValidations.Count,
        };
        foreach (var definition in _def.DataValidations)
        {
            if (!ExcelA1.TryParseRange(definition.Range, out var range))
                throw new NativeUsageException($"无效的数据验证范围: {definition.Range}");
            var validation = new DataValidation
            {
                Type = definition.Type switch
                {
                    "list" => DataValidationValues.List,
                    "whole" => DataValidationValues.Whole,
                    "decimal" => DataValidationValues.Decimal,
                    "date" => DataValidationValues.Date,
                    "time" => DataValidationValues.Time,
                    "textLength" => DataValidationValues.TextLength,
                    "custom" => DataValidationValues.Custom,
                    _ => throw new NativeUsageException(
                        $"不支持的数据验证类型: {definition.Type}"),
                },
                AllowBlank = definition.AllowBlank,
                ShowDropDown = !definition.ShowDropdown,
                ShowErrorMessage = definition.Error is not null
                                   || definition.ErrorTitle is not null,
                ShowInputMessage = definition.Prompt is not null
                                   || definition.PromptTitle is not null,
                PromptTitle = definition.PromptTitle,
                Prompt = definition.Prompt,
                ErrorTitle = definition.ErrorTitle,
                Error = definition.Error,
                SequenceOfReferences = new ListValue<StringValue>
                {
                    InnerText = ExcelA1.FormatRange(
                        range.MinCol, range.MinRow, range.MaxCol, range.MaxRow),
                },
            };
            validation.Append(new Formula1(definition.Formula1));
            if (definition.Formula2 is not null)
                validation.Append(new Formula2(definition.Formula2));
            validations.Append(validation);
        }
        return validations;
    }

    // ── 超级表 ───────────────────────────────────────────────────

    private TableParts? BuildTables(Dictionary<(int Row, int Col), string> textValues)
    {
        if (_def.Tables is null || _def.Tables.Count == 0)
            return null;
        var tableParts = new TableParts();
        var tableCount = 0;
        var occupied = new List<(string Name, ExcelRange Range)>();
        foreach (var tablePart in _part.TableDefinitionParts)
        {
            if (tablePart.Table?.Name?.Value is { } existingName
                && ExcelA1.TryParseRange(tablePart.Table.Reference?.Value ?? "", out var existingRange))
                occupied.Add((existingName, existingRange));
        }
        foreach (var definition in _def.Tables)
        {
            if (!ExcelA1.TryParseRange(definition.Range, out var range))
                throw new NativeRuntimeException($"无效的表格范围: {definition.Range}");
            if (range.MaxCol > MaxColumns || range.MaxRow > MaxRows)
                throw new NativeRuntimeException(
                    $"表格范围超出工作表上限（{MaxColumns} 列 × {MaxRows} 行）: {definition.Range}");
            if (range.MaxRow <= range.MinRow)
                throw new NativeRuntimeException("表格区域必须包含表头行和至少一行数据");
            if (string.IsNullOrWhiteSpace(definition.Name)
                || definition.Name.Any(char.IsWhiteSpace))
                throw new NativeRuntimeException($"表格名称无效: {definition.Name}");
            if (definition.Name.Length > 255)
                throw new NativeRuntimeException($"表格名称不能超过 255 个字符: {definition.Name}");
            if (!char.IsLetter(definition.Name[0])
                && definition.Name[0] != '_' && definition.Name[0] != '\\')
            {
                throw new NativeRuntimeException(
                    $"表格名称必须以字母、下划线或反斜杠开头: {definition.Name}");
            }
            if (ExcelA1.TryParseCell(definition.Name, out _, out _)
                || Regex.IsMatch(definition.Name, @"^R\d+C\d+$",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                throw new NativeRuntimeException(
                    $"表格名称不能与单元格引用相同: {definition.Name}");
            }
            if (definition.Name.IndexOfAny(['[', ']', ':', '*', '?', '/', '\\']) >= 0)
                throw new NativeRuntimeException($"表格名称包含非法字符: {definition.Name}");
            if (!_context.TableNames.Add(definition.Name))
                throw new NativeRuntimeException($"表格名称重复: {definition.Name}");
            if (_context.HasDefinedName(definition.Name)
                || (_def.DefinedNames?.Any(definedName => string.Equals(
                        definedName.Name, definition.Name, StringComparison.OrdinalIgnoreCase))
                    ?? false))
            {
                throw new NativeRuntimeException(
                    $"表格名称与定义名称冲突: {definition.Name}");
            }
            if (definition.ShowTotalsRow && range.MaxRow >= MaxRows)
            {
                throw new NativeRuntimeException(
                    $"表格 {definition.Name} 添加总计行后超出 {MaxRows} 行上限");
            }

            var fullRange = definition.ShowTotalsRow
                ? range with { MaxRow = range.MaxRow + 1 }
                : range;
            if (_def.MergedCells is not null)
            {
                foreach (var mergedText in _def.MergedCells)
                {
                    if (ExcelA1.TryParseRange(mergedText, out var merged)
                        && RangesOverlap(fullRange, merged))
                    {
                        throw new NativeRuntimeException(
                            $"表格 {definition.Name} 的区域与合并单元格 {mergedText} 重叠");
                    }
                }
            }
            foreach (var (otherName, otherRange) in occupied)
            {
                if (RangesOverlap(fullRange, otherRange))
                {
                    throw new NativeRuntimeException(
                        $"表格 {definition.Name} 的区域 {ExcelA1.FormatRange(
                            fullRange.MinCol, fullRange.MinRow, fullRange.MaxCol, fullRange.MaxRow)} "
                        + $"与表格 {otherName} 的区域重叠");
                }
            }
            occupied.Add((definition.Name, fullRange));

            var dataReference = ExcelA1.FormatRange(
                range.MinCol, range.MinRow, range.MaxCol, range.MaxRow);
            var tableReference = dataReference;
            var table = new Table
            {
                Id = _context.AllocateTableId(),
                Name = definition.Name,
                DisplayName = definition.Name,
                Reference = tableReference,
            };
            if (definition.ShowTotalsRow)
            {
                tableReference = ExcelA1.FormatRange(
                    range.MinCol, range.MinRow, range.MaxCol, range.MaxRow + 1);
                table.Reference = tableReference;
                table.TotalsRowCount = 1;
                table.TotalsRowShown = true;
            }
            table.Append(new AutoFilter { Reference = dataReference });

            var columns = new TableColumns
            {
                Count = (uint)(range.MaxCol - range.MinCol + 1),
            };
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var columnId = 1;
            for (var column = range.MinCol; column <= range.MaxCol; column++)
            {
                var header = textValues.TryGetValue((range.MinRow, column), out var text)
                             && !string.IsNullOrWhiteSpace(text)
                    ? text
                    : $"Column{columnId}";
                var name = header;
                var suffix = 2;
                while (!usedNames.Add(name))
                    name = $"{header}{suffix++}";
                columns.Append(new TableColumn
                {
                    Id = (uint)columnId++,
                    Name = name,
                });
            }
            table.Append(columns);
            table.Append(new TableStyleInfo
            {
                Name = definition.Style,
                ShowFirstColumn = false,
                ShowLastColumn = false,
                ShowRowStripes = definition.ShowRowStripes,
                ShowColumnStripes = false,
            });

            var tablePart = _part.AddNewPart<TableDefinitionPart>();
            tablePart.Table = table;
            tableParts.Append(new TablePart { Id = _part.GetIdOfPart(tablePart) });
            tableCount++;
        }
        tableParts.Count = (uint)tableCount;
        return tableParts;
    }

    // ── 定义名称 ─────────────────────────────────────────────────

    private static bool RangesOverlap(ExcelRange left, ExcelRange right) =>
        left.MinCol <= right.MaxCol && right.MinCol <= left.MaxCol
        && left.MinRow <= right.MaxRow && right.MinRow <= left.MaxRow;

    private void BuildDefinedNames()
    {
        if (_def.DefinedNames is null || _def.DefinedNames.Count == 0)
            return;
        var localSheetId = _context.SheetIndexByName.TryGetValue(_name, out var sheetIndex)
            ? sheetIndex
            : (int?)null;
        foreach (var definition in _def.DefinedNames)
        {
            var scope = definition.Local ? localSheetId : null;
            if (definition.Local && scope is null)
                throw new NativeRuntimeException(
                    $"工作表级定义名称 {definition.Name} 必须指定所在工作表");
            _context.AddDefinedName(
                definition.Name, definition.RefersTo, scope, replace: false);
        }
    }

    // ── 页面设置 ─────────────────────────────────────────────────

    private (PrintOptions? PrintOptions, PageMargins? Margins, S.PageSetup? PageSetupElement)
        BuildPageSetup()
    {
        if (_def.PageSetup is not { } definition)
            return (null, null, null);

        PrintOptions? printOptions = null;
        if (definition.CenterHorizontally is not null
            || definition.CenterVertically is not null)
        {
            printOptions = new PrintOptions
            {
                HorizontalCentered = definition.CenterHorizontally,
                VerticalCentered = definition.CenterVertically,
            };
        }

        var margins = new PageMargins
        {
            Left = definition.MarginLeft ?? 0.75,
            Right = definition.MarginRight ?? 0.75,
            Top = definition.MarginTop ?? 1,
            Bottom = definition.MarginBottom ?? 1,
            Header = 0.5,
            Footer = 0.5,
        };

        S.PageSetup? pageSetup = null;
        if (definition.Orientation is not null
            || definition.PaperSize is not null
            || definition.FitToWidth is not null
            || definition.FitToHeight is not null)
        {
            pageSetup = new S.PageSetup();
            if (definition.Orientation is not null)
            {
                pageSetup.Orientation = definition.Orientation == "landscape"
                    ? OrientationValues.Landscape
                    : OrientationValues.Portrait;
            }
            if (definition.PaperSize is not null)
                pageSetup.PaperSize = PaperSizeValue(definition.PaperSize);
            if (definition.FitToWidth is { } fitWidth)
                pageSetup.FitToWidth = (uint)fitWidth;
            if (definition.FitToHeight is { } fitHeight)
                pageSetup.FitToHeight = (uint)fitHeight;
        }

        var sheetIndex = _context.SheetIndexByName.TryGetValue(_name, out var index)
            ? index
            : (int?)null;
        if (sheetIndex is not null)
        {
            if (!string.IsNullOrWhiteSpace(definition.PrintArea))
            {
                var pieces = new List<string>();
                foreach (var piece in definition.PrintArea.Split(','))
                {
                    if (!ExcelA1.TryParseRange(piece, out var range))
                        throw new NativeUsageException($"无效的打印区域: {definition.PrintArea}");
                    var sheet = range.Sheet ?? _name;
                    pieces.Add(ExcelA1.AbsoluteRangeFormula(
                        sheet, range.MinCol, range.MinRow, range.MaxCol, range.MaxRow));
                }
                _context.AddDefinedName("_xlnm.Print_Area",
                    string.Join(",", pieces), sheetIndex, replace: true);
            }

            var titles = new List<string>();
            if (!string.IsNullOrWhiteSpace(definition.PrintTitleRows))
            {
                var parts = definition.PrintTitleRows.Split(':');
                var min = int.Parse(parts[0], CultureInfo.InvariantCulture);
                var max = parts.Length > 1
                    ? int.Parse(parts[1], CultureInfo.InvariantCulture)
                    : min;
                titles.Add(ExcelA1.AbsoluteRowRangeFormula(_name, min, max));
            }
            if (!string.IsNullOrWhiteSpace(definition.PrintTitleCols))
            {
                var parts = definition.PrintTitleCols.Split(':');
                var min = ExcelA1.ColumnIndex(parts[0]);
                var max = parts.Length > 1 ? ExcelA1.ColumnIndex(parts[1]) : min;
                titles.Add(ExcelA1.AbsoluteColumnRangeFormula(_name, min, max));
            }
            if (titles.Count > 0)
            {
                _context.AddDefinedName("_xlnm.Print_Titles",
                    string.Join(",", titles), sheetIndex, replace: true);
            }
        }

        return (printOptions, margins, pageSetup);
    }

    private static uint PaperSizeValue(string paperSize) => paperSize switch
    {
        "Letter" => 1,
        "A3" => 8,
        "A4" => 9,
        "B5" => 13,
        _ => 9,
    };
}

// ── 工作表保护 ───────────────────────────────────────────────────

internal static class ExcelSheetProtectionBuilder
{
    /// <summary>构造 &lt;sheetProtection&gt;：allow_* 为 true 时写 "0"（允许），否则 "1"。</summary>
    public static SheetProtection Build(ExcelSheetProtectionDef protection)
    {
        var element = new SheetProtection
        {
            Sheet = true,
            // objects/scenarios/pivotTables 模型未暴露，取值与 openpyxl 默认保持一致。
            Objects = false,
            Scenarios = false,
            PivotTables = true,
            FormatCells = !protection.AllowFormatCells,
            FormatColumns = !protection.AllowFormatColumns,
            FormatRows = !protection.AllowFormatRows,
            InsertRows = !protection.AllowInsertRows,
            InsertColumns = !protection.AllowInsertColumns,
            // 模型没有 allow_insert_hyperlinks，保持 openpyxl 默认（禁止插入超链接）。
            InsertHyperlinks = true,
            DeleteRows = !protection.AllowDeleteRows,
            DeleteColumns = !protection.AllowDeleteColumns,
            Sort = !protection.AllowSort,
            AutoFilter = !protection.AllowAutoFilter,
            SelectLockedCells = !protection.AllowSelectLockedCells,
            SelectUnlockedCells = !protection.AllowSelectUnlockedCells,
        };
        // 空密码不写 password 属性（openpyxl 同样省略），否则会额外写出 password=""。
        if (!string.IsNullOrEmpty(protection.Password))
            element.Password = ExcelLegacyPassword.Hash(protection.Password);
        return element;
    }
}

// ── OOXML legacy 密码 hash ───────────────────────────────────────

internal static class ExcelLegacyPassword
{
    /// <summary>
    /// OOXML 工作表保护的 legacy hash，逐位对齐
    /// openpyxl.utils.protection.hash_password（Python 后端就用它），
    /// 保证同一密码在 Python 与 C# 两个后端写出相同的 password 属性值。
    /// 按 Unicode 码点迭代（等价 Python 的字符串迭代），非 BMP 字符计 1 个长度。
    /// </summary>
    public static string Hash(string password)
    {
        long hash = 0;
        var length = 0;
        foreach (var rune in password.EnumerateRunes())
        {
            length++;
            long value = (long)rune.Value << length;
            var rotated = value >> 15;
            value &= 0x7FFF;
            hash ^= value | rotated;
        }
        hash ^= length;
        hash ^= 0xCE4B;
        return hash.ToString("X", CultureInfo.InvariantCulture);
    }
}

// ── 按行排序 ─────────────────────────────────────────────────────

internal sealed class ExcelSortPlan
{
    public List<List<JsonElement>> Rows { get; init; } = [];
    public List<int> Columns { get; init; } = [];
    public List<bool> Descending { get; init; } = [];
    /// <summary>sort.range 显式给出时的归一化区域；默认排序时为 null。</summary>
    public string? Range { get; init; }
}

internal static class ExcelSortPlanner
{
    private const int MaxColumns = ExcelFileOoxml.MaxColumns;
    // 排序值类别：数字 < 字符串 < 空（空白永远排最后）
    private const int RankNumber = 0;
    private const int RankString = 1;
    private const int RankEmpty = 2;

    public static ExcelSortPlan? Prepare(ExcelSheetDef sheet)
    {
        if (sheet.Sort is not { } sort)
            return null;
        if (sheet.MergedCells is { Count: > 0 })
        {
            throw new NativeRuntimeException(
                "工作表同时配置了 merged_cells 与 sort：排序会移动合并区域并破坏合并，"
                + "请去掉 merged_cells 或改为写入已排好序的数据");
        }
        if (sort.HasHeader && sheet.Headers is not { Count: > 0 })
        {
            throw new NativeRuntimeException(
                "排序配置 has_header=true，但工作表未提供 headers，无法按表头名定位排序列");
        }

        var columnsWritten = WrittenColumns(sheet);
        if (columnsWritten == 0)
            throw new NativeRuntimeException("排序配置了 sort，但工作表没有任何表头或数据列");
        var columns = sort.Keys
            .Select(key => ResolveColumn(sheet, key.Column, columnsWritten, sort.HasHeader))
            .ToList();
        var descending = sort.Keys.Select(key => !key.Ascending).ToList();

        ExcelRange? bounds = null;
        if (!string.IsNullOrWhiteSpace(sort.Range))
        {
            if (!ExcelA1.TryParseRange(sort.Range, out var parsed))
            {
                throw new NativeRuntimeException(
                    $"无效的排序区域: {sort.Range}（必须是含具体行列的 A1 范围，如 A1:C10）");
            }
            bounds = parsed;
            for (var index = 0; index < sort.Keys.Count; index++)
            {
                if (columns[index] < parsed.MinCol || columns[index] > parsed.MaxCol)
                {
                    throw new NativeRuntimeException(
                        $"排序键 {sort.Keys[index].Column}（第 {columns[index]} 列）"
                        + $"不在排序区域 {sort.Range} 的列范围内");
                }
            }
        }

        var dataStartRow = sheet.Headers is { Count: > 0 } ? 2 : 1;
        var selected = new List<int>();
        for (var index = 0; index < sheet.Data.Count; index++)
        {
            if (bounds is { } range)
            {
                var rowIndex = dataStartRow + index;
                if (rowIndex < range.MinRow || rowIndex > range.MaxRow)
                    continue;
            }
            selected.Add(index);
        }
        if (selected.Count == 0)
        {
            throw new NativeRuntimeException(bounds is null
                ? "排序配置了 sort，但没有可排序的数据行"
                : $"排序区域 {sort.Range} 未覆盖任何数据行（数据从第 {dataStartRow} 行开始）");
        }

        var rows = new List<List<JsonElement>>(sheet.Data);
        var keyed = selected
            .Select(index => (
                Values: columns.Select(column => Normalize(CellValue(rows[index], column))).ToArray(),
                Index: index))
            .ToList();
        var orderedIndices = Enumerable.Range(0, keyed.Count).ToList();
        orderedIndices.Sort((left, right) =>
        {
            var result = CompareRows(keyed[left].Values, keyed[right].Values, descending);
            return result != 0 ? result : left.CompareTo(right); // 显式保持稳定
        });
        var ordered = orderedIndices.Select(position => rows[keyed[position].Index]).ToList();
        for (var position = 0; position < selected.Count; position++)
            rows[selected[position]] = ordered[position];

        return new ExcelSortPlan
        {
            Rows = rows,
            Columns = columns,
            Descending = descending,
            Range = bounds is { } explicitRange
                ? ExcelA1.FormatRange(
                    explicitRange.MinCol, explicitRange.MinRow,
                    explicitRange.MaxCol, explicitRange.MaxRow)
                : null,
        };
    }

    /// <summary>按 sortState 规则在 AutoFilter 内追加 &lt;sortState&gt;/&lt;sortCondition&gt;。</summary>
    public static SortState BuildSortState(
        ExcelSortPlan plan, ExcelSheetDef sheet, string filterReference)
    {
        // 与 Excel/openpyxl 一致：sortState 区域含表头，sortCondition 只覆盖数据行。
        var range = plan.Range ?? filterReference;
        if (!ExcelA1.TryParseRange(range, out var parsed))
            throw new NativeRuntimeException($"排序区域无效: {range}");
        var dataStartRow = sheet.Headers is { Count: > 0 } ? 2 : 1;
        var dataMinRow = parsed.MinRow < dataStartRow && parsed.MaxRow >= dataStartRow
            ? dataStartRow
            : parsed.MinRow;
        var sortState = new SortState { Reference = range };
        foreach (var (column, descending) in plan.Columns.Zip(plan.Descending))
        {
            var letter = ExcelA1.ColumnName(column);
            var condition = new SortCondition
            {
                Reference = $"{letter}{dataMinRow}:{letter}{parsed.MaxRow}",
            };
            if (descending)
                condition.Descending = true;
            sortState.Append(condition);
        }
        return sortState;
    }

    private static int WrittenColumns(ExcelSheetDef sheet)
    {
        var count = sheet.Headers?.Count ?? 0;
        foreach (var row in sheet.Data)
        {
            if (row.Count > count)
                count = row.Count;
        }
        return count;
    }

    private static int ResolveColumn(
        ExcelSheetDef sheet, string keyColumn, int columnsWritten, bool hasHeader)
    {
        var text = keyColumn.Trim();
        if (hasHeader)
        {
            var headers = sheet.Headers!;
            var exact = headers.IndexOf(text);
            if (exact >= 0)
                return exact + 1;
            for (var index = 0; index < headers.Count; index++)
            {
                if (string.Equals(headers[index], text, StringComparison.OrdinalIgnoreCase))
                    return index + 1;
            }
            throw new NativeRuntimeException(
                $"排序键列 '{keyColumn}' 不在表头中（可用表头: {string.Join(", ", headers)}）");
        }

        int column;
        if (text.Length is >= 1 and <= 3 && text.All(char.IsAsciiLetter))
        {
            column = ExcelA1.ColumnIndex(text);
        }
        else if (text.Length > 0 && text.All(char.IsAsciiDigit))
        {
            // 超长数字串按越界处理，与 Python int() 后再校验范围的报错一致。
            column = long.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out var parsed)
                ? (parsed > int.MaxValue ? -1 : (int)parsed)
                : -1;
        }
        else
        {
            throw new NativeRuntimeException(
                $"排序列必须是列字母（如 'B'）或 1 基数列号: '{keyColumn}'");
        }
        if (column < 1 || column > MaxColumns)
            throw new NativeRuntimeException($"排序列号超出工作表列范围 1-{MaxColumns}: '{keyColumn}'");
        if (column > columnsWritten)
        {
            throw new NativeRuntimeException(
                $"排序列 '{text}' 超出数据范围（本次仅写入 {columnsWritten} 列）");
        }
        return column;
    }

    private static JsonElement CellValue(List<JsonElement> row, int column)
    {
        var index = column - 1;
        return index >= 0 && index < row.Count ? row[index] : default;
    }

    /// <summary>归一后的排序值：类别 + 数值或（小写, 原值）文本。</summary>
    private readonly record struct SortValue(int Rank, double Number, string Lower, string Original);

    /// <summary>与 Python _sort_value 对齐：数字 < 字符串 < 空，bool 当 1/0。</summary>
    private static SortValue Normalize(JsonElement value)
    {
        switch (value.ValueKind)
        {
            case JsonValueKind.Number when value.TryGetDouble(out var number):
                return double.IsFinite(number)
                    ? new SortValue(RankNumber, number, "", "")
                    : new SortValue(RankEmpty, 0, "", "");
            case JsonValueKind.True:
                return new SortValue(RankNumber, 1, "", "");
            case JsonValueKind.False:
                return new SortValue(RankNumber, 0, "", "");
            case JsonValueKind.String:
                var text = value.GetString() ?? "";
                return text.Length == 0
                    ? new SortValue(RankEmpty, 0, "", "")
                    : new SortValue(RankString, 0, text.ToLowerInvariant(), text);
            default:
                return new SortValue(RankEmpty, 0, "", "");
        }
    }

    private static int CompareRows(SortValue[] left, SortValue[] right, List<bool> descending)
    {
        for (var position = 0; position < left.Length; position++)
        {
            // Excel 升序/降序都保持空白单元格在末尾，所以空白判定不受方向影响。
            var leftBlank = left[position].Rank == RankEmpty;
            var rightBlank = right[position].Rank == RankEmpty;
            if (leftBlank || rightBlank)
            {
                if (leftBlank && rightBlank)
                    continue;
                return leftBlank ? 1 : -1;
            }

            int result;
            if (left[position].Rank != right[position].Rank)
            {
                result = left[position].Rank.CompareTo(right[position].Rank);
            }
            else if (left[position].Rank == RankNumber)
            {
                result = left[position].Number.CompareTo(right[position].Number);
            }
            else
            {
                result = string.CompareOrdinal(left[position].Lower, right[position].Lower);
                if (result == 0)
                    result = string.CompareOrdinal(left[position].Original, right[position].Original);
            }
            if (result != 0)
                return descending[position] ? -result : result;
        }
        return 0;
    }
}
