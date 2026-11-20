using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeNative;

/// <summary>Appends WordContent SectionItems to a main document part.</summary>
internal sealed class WordFileBuilder
{
    private readonly MainDocumentPart _main;
    private readonly Body _body;
    private readonly FontConfig? _defaultFont;
    private readonly SceneProfile? _scene;
    private readonly WordListNumbering _numbering;
    private int _docPrId;
    private int? _nextBookmarkId;

    public WordFileBuilder(
        MainDocumentPart main, FontConfig? defaultFont, SceneProfile? scene = null)
    {
        _main = main;
        _body = main.Document?.Body ?? throw new NativeRuntimeException("文档没有正文");
        _defaultFont = defaultFont;
        _scene = scene;
        _numbering = new WordListNumbering(main);
        _docPrId = WordFileOoxml.MaxDocPrId(main);
    }

    public void Add(SectionItem item)
    {
        switch (item.Type)
        {
            case "heading":
                AddHeading(item);
                break;
            case "paragraph":
                AddParagraph(item);
                break;
            case "table":
                AddTable(item);
                break;
            case "image":
                AddImage(item);
                break;
            case "page_break":
                WordFileOoxml.AppendBodyElement(_body,
                    new Paragraph(new Run(new Break { Type = BreakValues.Page })));
                break;
            case "bullet_list":
                AddList(item, bullet: true);
                break;
            case "numbered_list":
                AddList(item, bullet: false);
                break;
            case "code":
                AddCode(item);
                break;
            case "toc":
                AddToc(item);
                break;
            case "section_break":
                AddSectionBreak();
                break;
            case "hyperlink":
                AddHyperlink(item);
                break;
            case "field":
                AddField(item);
                break;
            case "bookmark":
                AddBookmark(item);
                break;
            default:
                throw new NativeUsageException($"未知的内容块类型: {item.Type}");
        }
    }

    private void AddHeading(SectionItem item)
    {
        var level = Math.Clamp(item.Level ?? 1, 1, 9);
        var format = new WordParagraphFormat { StyleId = $"Heading{level}" };
        if (item.Style is not null)
            ApplyParagraphStyle(format, item.Style);
        else if (_scene is not null)
            ApplySceneParagraphFormat(format, _scene);
        var paragraph = new Paragraph();
        var properties = WordFileOoxml.BuildParagraphProperties(format);
        if (properties.HasChildren)
            paragraph.Append(properties);
        if (!string.IsNullOrEmpty(item.Text))
        {
            paragraph.Append(WordFileOoxml.NewTextRun(item.Text, new WordRunFormat
            {
                FontName = _scene?.HeadingFont ?? _defaultFont?.Name,
                SizePt = _scene?.HeadingSizePt(level) ?? _defaultFont?.SizePt,
                Color = "000000",
            }));
        }
        WordFileOoxml.AppendBodyElement(_body, paragraph);
    }

    private void AddParagraph(SectionItem item)
    {
        var style = item.Style;
        var format = new WordParagraphFormat();
        if (style is not null || _scene is not null)
        {
            format.Alignment = FirstNonEmpty(style?.Alignment, _scene?.Alignment);
            format.LineSpacing = OrNonZero(style?.LineSpacing, _scene?.LineSpacing);
            format.FirstLineIndentPt =
                OrNonZero(style?.FirstLineIndentPt, _scene?.FirstLineIndent);
            format.SpaceBeforePt = style?.SpaceBeforePt;
            format.SpaceAfterPt = style?.SpaceAfterPt;
        }
        var font = new WordRunFormat
        {
            FontName = FirstNonEmpty(style?.FontName, _scene?.FontName, _defaultFont?.Name),
            SizePt = style?.SizePt ?? _scene?.BodySizePt ?? _defaultFont?.SizePt,
            Bold = style?.Bold,
            Italic = style?.Italic,
            Color = style?.Color,
        };
        var paragraph = new Paragraph();
        var properties = WordFileOoxml.BuildParagraphProperties(format);
        if (properties.HasChildren)
            paragraph.Append(properties);
        foreach (var chunk in WordFileOoxml.SplitLongText(item.Text))
            paragraph.Append(WordFileOoxml.NewTextRun(chunk, font));
        WordFileOoxml.AppendBodyElement(_body, paragraph);
    }

    private void AddTable(SectionItem item)
    {
        var definition = item.Table;
        if (definition is null)
            return;
        var headerCount = definition.Headers.Count;
        var rowCount = definition.Rows.Count + (headerCount > 0 ? 1 : 0);
        var columnCount = Math.Max(
            headerCount,
            definition.Rows.Count == 0 ? 0 : definition.Rows.Max(row => row.Count));
        if (rowCount == 0 || columnCount == 0)
            return;

        var usableWidth = UsableWidthTwips();
        var defaultColumnWidth = Math.Max(1, usableWidth / columnCount);
        var gridColumns = new List<GridColumn>();
        for (var column = 0; column < columnCount; column++)
            gridColumns.Add(new GridColumn { Width = defaultColumnWidth.ToString() });

        var tableProperties = new TableProperties(
            new TableStyle { Val = "TableGrid" },
            new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto });
        if (definition.ColumnWidths is { Count: > 0 })
            tableProperties.Append(new TableLayout { Type = TableLayoutValues.Fixed });
        tableProperties.Append(new TableLook
        {
            Val = "04A0",
            FirstRow = true,
            LastRow = false,
            FirstColumn = true,
            LastColumn = false,
            NoHorizontalBand = false,
            NoVerticalBand = true,
        });

        var table = new Table(tableProperties, new TableGrid(gridColumns));
        for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
        {
            var isHeader = headerCount > 0 && rowIndex == 0;
            var values = isHeader
                ? definition.Headers
                : definition.Rows[rowIndex - (headerCount > 0 ? 1 : 0)];
            var row = new TableRow();
            for (var column = 0; column < columnCount; column++)
            {
                var text = column < values.Count ? values[column] : "";
                if (isHeader && definition.HeaderStyle is { } headerStyle)
                {
                    var bold = headerStyle.Bold ?? true;
                    var cell = NewCell(text, WordFileOoxml.BuildRunProperties(new WordRunFormat
                    {
                        FontName = headerStyle.FontName,
                        SizePt = headerStyle.SizePt,
                        Bold = bold,
                        Italic = false,
                        Color = headerStyle.FontColor,
                    }), defaultColumnWidth);
                    if (!string.IsNullOrEmpty(headerStyle.BgColor))
                    {
                        var cellProperties = EnsureCellProperties(cell);
                        cellProperties.Shading = new Shading { Fill = headerStyle.BgColor };
                    }
                    row.Append(cell);
                }
                else
                {
                    row.Append(NewCell(text, null, defaultColumnWidth));
                }
            }
            table.Append(row);
        }

        if (definition.ColumnWidths is { Count: > 0 } widths)
        {
            var rows = table.Elements<TableRow>().ToList();
            for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                for (var column = 0; column < widths.Count && column < columnCount; column++)
                {
                    var cell = GetCellAt(rows[rowIndex], column);
                    if (cell is null)
                        continue;
                    var properties = EnsureCellProperties(cell);
                    properties.TableCellWidth = new TableCellWidth
                    {
                        Width = WordFileOoxml.InchesToTwips(widths[column]),
                        Type = TableWidthUnitValues.Dxa,
                    };
                }
            }
            for (var column = 0; column < widths.Count && column < gridColumns.Count; column++)
                gridColumns[column].Width = WordFileOoxml.InchesToTwips(widths[column]);
        }

        foreach (var merge in definition.MergedCells)
        {
            if (!TryParseRange(merge.Range, out var row0, out var column0,
                    out var row1, out var column1))
                continue;
            try
            {
                MergeCells(table, row0, column0, row1, column1);
            }
            catch (Exception)
            {
                // python-docx logs and skips ranges it cannot apply; do the same.
            }
        }

        WordFileOoxml.AppendBodyElement(_body, table);
    }

    private static TableCell NewCell(
        string text, RunProperties? runProperties, int widthTwips)
    {
        var paragraph = new Paragraph();
        if (!string.IsNullOrEmpty(text))
        {
            var run = new Run();
            if (runProperties is not null && runProperties.HasChildren)
                run.Append(runProperties.CloneNode(true));
            run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            paragraph.Append(run);
        }
        return new TableCell(
            new TableCellProperties(new TableCellWidth
            {
                Width = widthTwips.ToString(System.Globalization.CultureInfo.InvariantCulture),
                Type = TableWidthUnitValues.Dxa,
            }),
            paragraph);
    }

    private int UsableWidthTwips()
    {
        var section = _body.Elements<SectionProperties>().LastOrDefault();
        var width = section?.GetFirstChild<PageSize>()?.Width?.Value ?? 12240U;
        var margin = section?.GetFirstChild<PageMargin>();
        var left = margin?.Left?.Value ?? 1800U;
        var right = margin?.Right?.Value ?? 1800U;
        var usable = (long)width - left - right;
        return usable <= 0 ? 8640 : (int)usable;
    }

    private static TableCellProperties EnsureCellProperties(TableCell cell)
    {
        var properties = cell.TableCellProperties;
        if (properties is not null)
            return properties;
        properties = new TableCellProperties();
        cell.PrependChild(properties);
        return properties;
    }

    private static int CellSpan(TableCell cell) =>
        cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1;

    private static TableCell? GetCellAt(TableRow row, int columnIndex)
    {
        var offset = row.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0;
        var cursor = offset;
        foreach (var cell in row.Elements<TableCell>())
        {
            var span = Math.Max(1, CellSpan(cell));
            if (columnIndex >= cursor && columnIndex < cursor + span)
                return cell;
            cursor += span;
        }
        return null;
    }

    private static bool TryParseRange(
        string range, out int row0, out int column0, out int row1, out int column1)
    {
        row0 = column0 = row1 = column1 = 0;
        var match = Regex.Match(range, @"^([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)$");
        if (!match.Success)
            return false;
        column0 = ColumnIndex(match.Groups[1].Value);
        row0 = int.Parse(match.Groups[2].Value) - 1;
        column1 = ColumnIndex(match.Groups[3].Value);
        row1 = int.Parse(match.Groups[4].Value) - 1;
        return row0 <= row1 && column0 <= column1 && row0 >= 0 && column0 >= 0;
    }

    private static int ColumnIndex(string letters)
    {
        var result = 0;
        foreach (var character in letters.ToUpperInvariant())
            result = result * 26 + (character - 'A' + 1);
        return result - 1;
    }

    private static void MergeCells(
        Table table, int row0, int column0, int row1, int column1)
    {
        var rows = table.Elements<TableRow>().ToList();
        if (row1 >= rows.Count)
            throw new InvalidOperationException("合并范围超出表格行数");
        var top = GetCellAt(rows[row0], column0)
                  ?? throw new InvalidOperationException("合并范围起点没有单元格");
        GrowTo(rows, row0, column0, column1 - column0 + 1, row1 - row0 + 1, top);
    }

    private static void GrowTo(
        List<TableRow> rows, int rowIndex, int columnIndex, int width, int height,
        TableCell topCell)
    {
        var cell = GetCellAt(rows[rowIndex], columnIndex)
                   ?? throw new InvalidOperationException("合并范围没有连续单元格");
        SpanToWidth(cell, topCell, width);
        var properties = EnsureCellProperties(cell);
        properties.VerticalMerge = ReferenceEquals(cell, topCell)
            ? (height > 1
                ? new VerticalMerge { Val = MergedCellValues.Restart }
                : null)
            : new VerticalMerge();
        if (height > 1)
            GrowTo(rows, rowIndex + 1, columnIndex, width, height - 1, topCell);
    }

    private static void SpanToWidth(TableCell cell, TableCell topCell, int gridWidth)
    {
        MoveContentTo(cell, topCell);
        while (CellSpan(cell) < gridWidth)
        {
            var next = cell.NextSibling<TableCell>()
                       ?? throw new InvalidOperationException("合并范围宽度不足");
            var span = CellSpan(cell) + CellSpan(next);
            if (span > gridWidth)
                throw new InvalidOperationException("合并范围跨越了已合并单元格");
            MoveContentTo(next, topCell);
            var mergedWidth = CellWidthTwips(cell) is { } firstWidth
                              && CellWidthTwips(next) is { } secondWidth
                ? firstWidth + secondWidth
                : (int?)null;
            next.Remove();
            var properties = EnsureCellProperties(cell);
            properties.GridSpan = new GridSpan { Val = span };
            if (mergedWidth is { } width)
            {
                properties.TableCellWidth = new TableCellWidth
                {
                    Width = width.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    Type = TableWidthUnitValues.Dxa,
                };
            }
        }
    }

    private static int? CellWidthTwips(TableCell cell)
    {
        var width = cell.TableCellProperties?.TableCellWidth;
        if (width?.Type?.Value != TableWidthUnitValues.Dxa)
            return null;
        return int.TryParse(width.Width?.Value, out var twips) ? twips : null;
    }

    private static void MoveContentTo(TableCell source, TableCell target)
    {
        if (ReferenceEquals(source, target) || IsEmptyCell(source))
            return;
        RemoveTrailingEmptyParagraph(target);
        foreach (var child in source.ChildElements.ToList())
        {
            if (child is TableCellProperties)
                continue;
            child.Remove();
            target.AppendChild(child);
        }
        source.AppendChild(new Paragraph());
    }

    private static bool IsEmptyCell(TableCell cell)
    {
        var blocks = cell.ChildElements
            .Where(child => child is Paragraph or Table)
            .ToList();
        if (blocks.Count > 1)
            return false;
        if (blocks.Count == 0)
            return true;
        return blocks[0] is Paragraph paragraph
               && !paragraph.Elements<Run>().Any();
    }

    private static void RemoveTrailingEmptyParagraph(TableCell cell)
    {
        var blocks = cell.ChildElements
            .Where(child => child is Paragraph or Table)
            .ToList();
        if (blocks.Count == 0)
            return;
        var last = blocks[^1];
        if (last is Paragraph paragraph && !paragraph.Elements<Run>().Any())
            last.Remove();
    }

    private void AddImage(SectionItem item)
    {
        var image = item.Image;
        if (image is null || string.IsNullOrWhiteSpace(image.Path) || !File.Exists(image.Path))
            return;
        var dimensions = WordFileOoxml.TryReadImageSize(image.Path);
        double widthInches;
        double heightInches;
        if (image.Width is null && image.Height is null)
        {
            widthInches = 4.0;
            heightInches = dimensions is { Width: > 0 } size
                ? widthInches * size.Height / size.Width
                : widthInches * 0.75;
        }
        else
        {
            widthInches = image.Width ?? 4.0;
            heightInches = image.Height
                           ?? (dimensions is { Width: > 0 } size
                               ? widthInches * size.Height / size.Width
                               : widthInches * 0.75);
        }
        var cx = (long)Math.Round(widthInches * WordFileOoxml.EmuPerInch);
        var cy = (long)Math.Round(heightInches * WordFileOoxml.EmuPerInch);
        var imagePart = _main.AddImagePart(WordFileOoxml.ImagePartTypeFor(image.Path));
        using (var stream = File.OpenRead(image.Path))
            imagePart.FeedData(stream);
        var relationshipId = _main.GetIdOfPart(imagePart);
        var docPrId = ++_docPrId;
        var drawing = WordFileOoxml.BuildInlineImage(
            relationshipId, cx, cy, docPrId, $"Picture {docPrId}");
        WordFileOoxml.AppendBodyElement(_body,
            new Paragraph(new Run(drawing)));
    }

    private void AddList(SectionItem item, bool bullet)
    {
        var numberId = _numbering.GetNumberId(bullet);
        foreach (var text in item.Items)
        {
            var format = new WordParagraphFormat
            {
                StyleId = bullet ? "ListBullet" : "ListNumber",
                NumberingId = numberId,
            };
            if (_scene is not null)
            {
                format.Alignment = _scene.Alignment;
                format.LineSpacing = _scene.LineSpacing;
            }
            var paragraph = new Paragraph();
            var properties = WordFileOoxml.BuildParagraphProperties(format);
            if (properties.HasChildren)
                paragraph.Append(properties);
            if (!string.IsNullOrEmpty(text))
            {
                var font = _scene is null
                    ? new WordRunFormat
                    {
                        FontName = _defaultFont?.Name,
                        SizePt = _defaultFont?.SizePt,
                    }
                    : null;
                paragraph.Append(WordFileOoxml.NewTextRun(text, font));
            }
            WordFileOoxml.AppendBodyElement(_body, paragraph);
        }
    }

    private void AddCode(SectionItem item)
    {
        var format = new WordParagraphFormat
        {
            ShadingFill = "F5F5F5",
            ShadingIsNilPattern = true,
        };
        if (item.Style is not null)
            ApplyParagraphStyle(format, item.Style);
        else if (_scene is not null)
        {
            format.Alignment = _scene.Alignment;
            format.LineSpacing = _scene.LineSpacing;
        }
        var paragraph = new Paragraph();
        var properties = WordFileOoxml.BuildParagraphProperties(format);
        if (properties.HasChildren)
            paragraph.Append(properties);
        paragraph.Append(WordFileOoxml.NewTextRun(item.Text ?? "", new WordRunFormat
        {
            FontName = "Consolas",
            SizePt = 10,
            Bold = false,
            Italic = false,
            Color = "1E1E1E",
        }));
        WordFileOoxml.AppendBodyElement(_body, paragraph);
    }

    private void AddToc(SectionItem item)
    {
        var title = string.IsNullOrEmpty(item.Text) ? "目录" : item.Text;
        var headingFormat = new WordParagraphFormat { Alignment = "center" };
        var heading = new Paragraph();
        var headingProperties = WordFileOoxml.BuildParagraphProperties(headingFormat);
        if (headingProperties.HasChildren)
            heading.Append(headingProperties);
        heading.Append(WordFileOoxml.NewTextRun(title, new WordRunFormat
        {
            FontName = _scene?.FontName,
            SizePt = 16,
            Bold = true,
        }));
        WordFileOoxml.AppendBodyElement(_body, heading);

        var levels = string.IsNullOrEmpty(item.TocLevels) ? "1-3" : item.TocLevels;
        WordFileOoxml.AppendBodyElement(
            _body, WordFileOoxml.BuildTocFieldParagraph(levels));

        var placeholder = new Paragraph(WordFileOoxml.NewTextRun(
            "（请在 Word 中右键点击目录区域，选择「更新域」以生成完整目录）",
            new WordRunFormat
            {
                SizePt = 9,
                Color = "999999",
                Italic = true,
            }));
        WordFileOoxml.AppendBodyElement(_body, placeholder);
        WordFileOoxml.AppendBodyElement(_body,
            new Paragraph(new Run(new Break { Type = BreakValues.Page })));
    }

    private void AddSectionBreak()
    {
        var section = WordFileOoxml.EnsureBodySection(_body);
        var clone = (SectionProperties)section.CloneNode(true);
        WordFileOoxml.AppendBodyElement(_body,
            new Paragraph(new ParagraphProperties(clone)));
        WordFileOoxml.PrepareNewLastSection(section);
    }

    private void AddHyperlink(SectionItem item)
    {
        var link = item.Hyperlink
                   ?? throw new NativeUsageException("type=hyperlink 时必须提供 hyperlink 定义");
        WordFileOoxml.AppendBodyElement(_body,
            WordFileOoxml.BuildHyperlinkParagraph(_main, link, _defaultFont,
                FormatFromStyle(item.Style)));
    }

    private void AddField(SectionItem item)
    {
        var field = item.Field
                    ?? throw new NativeUsageException("type=field 时必须提供 field 定义");
        WordFileOoxml.AppendBodyElement(_body,
            WordFileOoxml.BuildFieldParagraph(field, FormatFromStyle(item.Style)));
    }

    private void AddBookmark(SectionItem item)
    {
        var definition = item.Bookmark
                         ?? throw new NativeUsageException(
                             "type=bookmark 时必须提供 bookmark 定义");
        if (_body.Descendants<BookmarkStart>()
            .Any(start => start.Name?.Value == definition.Name))
            throw new NativeUsageException($"书签名已存在: '{definition.Name}'");
        var paragraph = new Paragraph();
        if (item.Style is not null)
        {
            var properties = WordFileOoxml.BuildParagraphProperties(
                FormatFromStyle(item.Style));
            if (properties.HasChildren)
                paragraph.Append(properties);
        }
        var startId = AllocateBookmarkId();
        paragraph.Append(new BookmarkStart
        {
            Id = startId.ToString(System.Globalization.CultureInfo.InvariantCulture),
            Name = definition.Name,
        });
        if (!string.IsNullOrEmpty(item.Text))
        {
            var style = item.Style;
            var format = new WordRunFormat
            {
                FontName = FirstNonEmpty(style?.FontName, _scene?.FontName, _defaultFont?.Name),
                SizePt = style?.SizePt ?? _scene?.BodySizePt ?? _defaultFont?.SizePt,
                Bold = style?.Bold,
                Italic = style?.Italic,
                Color = style?.Color,
            };
            foreach (var chunk in WordFileOoxml.SplitLongText(item.Text))
                paragraph.Append(WordFileOoxml.NewTextRun(chunk, format));
        }
        paragraph.Append(new BookmarkEnd
        {
            Id = startId.ToString(System.Globalization.CultureInfo.InvariantCulture),
        });
        WordFileOoxml.AppendBodyElement(_body, paragraph);
    }

    /// <summary>Bookmark ids are document-unique: max existing id + 1, then sequential.</summary>
    private int AllocateBookmarkId()
    {
        if (_nextBookmarkId is null)
        {
            var maxId = -1;
            foreach (var start in _body.Descendants<BookmarkStart>())
            {
                if (int.TryParse(start.Id?.Value,
                        System.Globalization.NumberStyles.Integer,
                        System.Globalization.CultureInfo.InvariantCulture, out var parsed))
                    maxId = Math.Max(maxId, parsed);
            }
            _nextBookmarkId = maxId + 1;
        }
        var id = _nextBookmarkId!.Value;
        _nextBookmarkId = id + 1;
        return id;
    }

    private static WordParagraphFormat? FormatFromStyle(ParagraphStyle? style)
    {
        if (style is null)
            return null;
        var format = new WordParagraphFormat();
        ApplyParagraphStyle(format, style);
        return format;
    }

    private static void ApplyParagraphStyle(WordParagraphFormat format, ParagraphStyle? style)
    {
        if (style is null)
            return;
        format.Alignment = style.Alignment;
        format.LineSpacing = style.LineSpacing;
        format.FirstLineIndentPt = style.FirstLineIndentPt;
        format.SpaceBeforePt = style.SpaceBeforePt;
        format.SpaceAfterPt = style.SpaceAfterPt;
    }

    private static void ApplySceneParagraphFormat(WordParagraphFormat format, SceneProfile scene)
    {
        format.Alignment = scene.Alignment;
        format.LineSpacing = scene.LineSpacing;
        format.FirstLineIndentPt = scene.FirstLineIndent;
    }

    /// <summary>Python `a or b or c` for optional strings.</summary>
    private static string? FirstNonEmpty(params string?[] values) =>
        values.FirstOrDefault(value => !string.IsNullOrEmpty(value));

    private static double? OrNonZero(double? explicitValue, double? fallback) =>
        explicitValue is { } value && value != 0 ? value : fallback;

    private static int? OrNonZero(int? explicitValue, int? fallback) =>
        explicitValue is { } value && value != 0 ? value : fallback;
}

/// <summary>Orchestrates creating/appending whole Word documents.</summary>
internal static class WordFileWriter
{
    public static void PrepareNewDocument(
        WordprocessingDocument document, bool pythonDocxThemeFonts = false)
    {
        var main = document.AddMainDocumentPart();
        main.Document = new Document(new Body(WordFileOoxml.DefaultSection()));
        var settingsPart = main.DocumentSettingsPart
                           ?? main.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings ??= new Settings();
        WordFileOoxml.EnsureRequiredStyles(main, pythonDocxThemeFonts);
    }

    public static void CreateDocument(
        string outputPath,
        WordContent content,
        string? template,
        PageSetup? pageSetup,
        HeaderFooter? headerFooter,
        FontConfig? defaultFont,
        SceneProfile? scene,
        CoverConfig? cover,
        bool generateCover)
    {
        var tempPath = WordFileOoxml.TempSiblingPath(outputPath);
        try
        {
            if (!string.IsNullOrWhiteSpace(template))
            {
                NativeCommon.EnsureFile(template, ".docx", ".dotx");
                File.Copy(Path.GetFullPath(template), tempPath, overwrite: true);
                using var document = WordprocessingDocument.Open(tempPath, true);
                var main = document.MainDocumentPart
                           ?? throw new NativeRuntimeException("模板缺少主文档部件");
                if (main.Document is null)
                    main.Document = new Document(new Body(WordFileOoxml.DefaultSection()));
                else
                    main.Document.Body ??= new Body(WordFileOoxml.DefaultSection());
                WordFileOoxml.EnsureRequiredStyles(main);
                Build(main, document, content, pageSetup, headerFooter, defaultFont, scene,
                    cover, generateCover);
            }
            else
            {
                using var document = WordprocessingDocument.Create(
                    tempPath, WordprocessingDocumentType.Document);
                PrepareNewDocument(document, defaultFont is null && scene is null);
                var main = document.MainDocumentPart!;
                Build(main, document, content, pageSetup, headerFooter, defaultFont, scene,
                    cover, generateCover);
            }
            WordFileOoxml.ReplaceFile(tempPath, outputPath);
        }
        finally
        {
            if (File.Exists(tempPath))
            {
                try { File.Delete(tempPath); } catch { /* best effort */ }
            }
        }
    }

    private static void Build(
        MainDocumentPart main,
        WordprocessingDocument document,
        WordContent content,
        PageSetup? pageSetup,
        HeaderFooter? headerFooter,
        FontConfig? defaultFont,
        SceneProfile? scene,
        CoverConfig? cover,
        bool generateCover)
    {
        var body = main.Document?.Body
                   ?? throw new NativeRuntimeException("文档没有正文");
        if (generateCover)
        {
            var recipe = WordFileCovers.Select(cover?.DocType ?? "general", scene?.Name);
            if (cover?.RecipeId is { } recipeId
                && WordFileCovers.Recipes.TryGetValue(recipeId, out var selected))
                recipe = selected;
            var fallbackTitle = content.Sections.FirstOrDefault()?.Text;
            var title = !string.IsNullOrEmpty(cover?.Title)
                ? cover!.Title!
                : (string.IsNullOrEmpty(fallbackTitle) ? "文档标题" : fallbackTitle!);
            WordFileCovers.Generate(
                body, recipe, title, cover?.Subtitle, cover?.Author, cover?.Date,
                cover?.Organization, cover?.ExtraLines);
        }
        if (pageSetup is not null)
            WordFileOoxml.ApplyPageSetup(body, pageSetup);
        if (headerFooter is not null)
            WordFileOoxml.ApplyHeaderFooter(main, headerFooter);
        if (pageSetup is { WatermarkText: not null } || pageSetup?.WatermarkColor is not null)
            WordFileOoxml.ApplyWatermark(
                main, pageSetup.WatermarkText, pageSetup.WatermarkColor);
        var builder = new WordFileBuilder(main, defaultFont, scene);
        foreach (var section in content.Sections)
            builder.Add(section);
        if (content.Properties is not null)
            WordFileOoxml.ApplyCoreProperties(document, content.Properties);
        main.Document.Save();
    }

    public static void AppendToDocument(
        string filePath, List<SectionItem> sections, FontConfig? defaultFont)
    {
        WordFileOoxml.SaveThroughTemp(filePath, document =>
        {
            var main = document.MainDocumentPart
                       ?? throw new NativeRuntimeException("文档没有主部件");
            if (main.Document is null)
                main.Document = new Document(new Body(WordFileOoxml.DefaultSection()));
            else
                main.Document.Body ??= new Body(WordFileOoxml.DefaultSection());
            WordFileOoxml.EnsureRequiredStyles(main);
            var builder = new WordFileBuilder(main, defaultFont);
            foreach (var section in sections)
                builder.Add(section);
            main.Document.Save();
        });
    }
}

