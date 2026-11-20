using System.Globalization;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XCharts = DocumentFormat.OpenXml.Drawing.Charts;
using XPresentation = DocumentFormat.OpenXml.Presentation;
using XSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeNative;

internal readonly record struct PptDefaults(double Left, double Top, double Width, double Height)
{
    public static readonly PptDefaults Body = new(0.5, 1.5, 8.0, 4.5);
    public static readonly PptDefaults LeftColumn = new(0.5, 1.5, 4.0, 4.5);
    public static readonly PptDefaults RightColumn = new(5.0, 1.5, 4.0, 4.5);
}

/// <summary>幻灯片内容写入器：占位符、文本、列表、表格、图表、图片、形状、连接线。</summary>
internal sealed class PptSlideWriter
{
    public SlidePart Part { get; }
    public XDocument Document { get; }
    private readonly XElement _tree;
    private uint _nextId;

    public PptSlideWriter(SlidePart part, XDocument document)
    {
        Part = part;
        Document = document;
        _tree = PptFileOoxml.SlideTree(document);
        _nextId = PptFileContent.MaxShapeId(document) + 1;
        if (_nextId < 2)
            _nextId = 2;
    }

    public void Save() => PptFileOoxml.SaveSlide(Part, Document);

    public uint TakeId() => NextId();

    public XElement Tree => _tree;

    public void Append(XElement element) => _tree.Add(element);

    public void SetBackground(string color)
    {
        var common = Document.Root!.Element(PptFileOoxml.P + "cSld");
        common?.Element(PptFileOoxml.P + "bg")?.Remove();
        var background = new XElement(PptFileOoxml.P + "bg",
            new XElement(PptFileOoxml.P + "bgPr",
                PptFileOoxml.SolidFill(color),
                new XElement(PptFileOoxml.A + "effectLst")));
        common?.AddFirst(background);
    }

    public void SetPlaceholderText(string text, params string[] types)
    {
        if (string.IsNullOrEmpty(text))
            return;
        var shape = FindPlaceholderByType(types);
        if (shape is null)
            return;
        SetTextFrameText(shape, text);
    }

    public XElement? FindPlaceholderByType(params string[] types)
    {
        return PptFileOoxml.ShapeElements(_tree).FirstOrDefault(shape =>
        {
            var info = PptFileOoxml.PlaceholderInfo(shape);
            return info is not null && info.Value.Type is not null
                   && types.Any(type => string.Equals(type, info.Value.Type,
                       StringComparison.OrdinalIgnoreCase));
        });
    }

    public void WriteContent(IEnumerable<PptSlideItem>? items, PptDefaults defaults,
        int placeholderIndex)
    {
        if (items is null)
            return;
        foreach (var item in items)
            WriteItem(item, defaults, placeholderIndex);
    }

    private void WriteItem(PptSlideItem item, PptDefaults defaults, int placeholderIndex)
    {
        switch (item.Type)
        {
            case "paragraph":
                if (string.IsNullOrEmpty(item.Text))
                    return;
                var placeholder = FindBodyPlaceholder(placeholderIndex);
                if (placeholder is not null)
                {
                    AppendParagraph(PptFileOoxml.TextBodyOf(placeholder)
                                    ?? placeholder.Element(PptFileOoxml.P + "txBody"),
                        PptFileContent.BuildParagraph(item.Text!, item.Level, item.Style,
                            bullet: false, numbered: false));
                }
                else
                {
                    AddTextBox(defaults.Left, defaults.Top, defaults.Width, defaults.Height,
                        [PptFileContent.BuildPlainParagraph(item.Text!)]);
                }
                break;
            case "table" when item.Table is not null:
                AddTable(item.Table, defaults);
                break;
            case "image" when item.Image is not null:
                AddImage(item.Image!);
                break;
            case "shape" when item.Shape is not null:
                AddShape(item.Shape!);
                break;
            case "connector" when item.Connector is not null:
                AddConnector(item.Connector!, "straight", tailArrow: item.Connector!.Arrowhead,
                    headArrow: item.Connector.BeginArrowhead);
                break;
            case "bullet_list":
            case "numbered_list":
                AddList(item, defaults);
                break;
            case "chart" when item.Chart is not null:
                AddChart(item.Chart!, defaults);
                break;
        }
    }

    private XElement? FindBodyPlaceholder(int index)
    {
        var placeholders = PptFileOoxml.ShapeElements(_tree)
            .Where(PptFileOoxml.IsPlaceholder).ToList();
        var match = placeholders.FirstOrDefault(shape =>
            PptFileOoxml.PlaceholderInfo(shape)?.Idx == index);
        if (match is not null)
            return match;
        var body = placeholders.FirstOrDefault(shape =>
            PptFileOoxml.PlaceholderInfo(shape)?.Idx == 1);
        if (body is not null)
            return body;
        return placeholders.Count > 1 ? placeholders[1] : null;
    }

    public static void SetTextFrameText(XElement shape, string text)
    {
        var textBody = shape.Element(PptFileOoxml.P + "txBody");
        if (textBody is null)
        {
            textBody = new XElement(PptFileOoxml.P + "txBody",
                new XElement(PptFileOoxml.A + "bodyPr"),
                new XElement(PptFileOoxml.A + "lstStyle"));
            shape.Add(textBody);
        }
        foreach (var paragraph in textBody.Elements(PptFileOoxml.A + "p").ToList())
            paragraph.Remove();
        textBody.Add(PptFileContent.BuildPlainParagraph(text));
    }

    public static void AppendParagraph(XElement? textBody, XElement paragraph)
    {
        if (textBody is null)
            return;
        var paragraphs = textBody.Elements(PptFileOoxml.A + "p").ToList();
        if (paragraphs.Count == 1 && IsEmptyParagraph(paragraphs[0]))
            paragraphs[0].Remove();
        textBody.Add(paragraph);
    }

    private static bool IsEmptyParagraph(XElement paragraph) =>
        !paragraph.HasElements && paragraph.Attributes().All(attribute =>
            attribute.Name.LocalName is "lvl" or "algn" or "endParaRPr");

    public XElement AddTextBox(double x, double y, double width, double height,
        IEnumerable<XElement> paragraphs)
    {
        var id = NextId();
        var shape = new XElement(PptFileOoxml.P + "sp",
            new XElement(PptFileOoxml.P + "nvSpPr",
                new XElement(PptFileOoxml.P + "cNvPr",
                    new XAttribute("id", id), new XAttribute("name", $"TextBox {id}")),
                new XElement(PptFileOoxml.P + "cNvSpPr",
                    new XAttribute("txBox", "1")),
                new XElement(PptFileOoxml.P + "nvPr")),
            new XElement(PptFileOoxml.P + "spPr",
                PptFileOoxml.El(PptFileOoxml.A, "xfrm",
                    new XElement(PptFileOoxml.A + "off",
                        new XAttribute("x", PptFileOoxml.Inches(x)),
                        new XAttribute("y", PptFileOoxml.Inches(y))),
                    new XElement(PptFileOoxml.A + "ext",
                        new XAttribute("cx", PptFileOoxml.Inches(width)),
                        new XAttribute("cy", PptFileOoxml.Inches(height)))),
                new XElement(PptFileOoxml.A + "prstGeom",
                    new XAttribute("prst", "rect"),
                    new XElement(PptFileOoxml.A + "avLst")),
                new XElement(PptFileOoxml.A + "noFill")),
            new XElement(PptFileOoxml.P + "txBody",
                new XElement(PptFileOoxml.A + "bodyPr",
                    new XAttribute("wrap", "square"),
                    new XElement(PptFileOoxml.A + "spAutoFit")),
                new XElement(PptFileOoxml.A + "lstStyle"),
                paragraphs));
        _tree.Add(shape);
        return shape;
    }

    private void AddList(PptSlideItem item, PptDefaults defaults)
    {
        var numbered = item.Type == "numbered_list";
        var paragraphs = new List<XElement>();
        foreach (var text in item.Items ?? [])
        {
            var level = numbered ? item.Level : 0;
            paragraphs.Add(PptFileContent.BuildParagraph(text, level, item.Style,
                bullet: !numbered, numbered: numbered));
        }
        AddTextBox(defaults.Left, 1.8, defaults.Width, defaults.Height, paragraphs);
    }

    private void AddTable(PptSlideTable table, PptDefaults defaults)
    {
        var rowCount = table.Rows.Count + (table.Headers.Count > 0 ? 1 : 0);
        var columnCount = Math.Max(table.Headers.Count,
            table.Rows.Count == 0 ? 0 : table.Rows.Max(row => row.Count));
        if (rowCount == 0 || columnCount == 0)
            return;

        var left = table.X ?? defaults.Left;
        var top = table.Y ?? 2.0;
        var width = table.Width ?? defaults.Width;
        var height = table.Height ?? 2.5;

        var widths = table.ColumnWidths is { Count: > 0 }
            ? PptFileContent.ComputeColumnWidths(
                table.ColumnWidths, columnCount, table.Width ?? 7.0)
            : Enumerable.Repeat(width / columnCount, columnCount).ToList();
        var rowHeight = height / rowCount;
        var rows = new List<XElement>();
        if (table.Headers.Count > 0)
        {
            var cells = new List<XElement>();
            for (var column = 0; column < columnCount; column++)
            {
                var text = column < table.Headers.Count ? table.Headers[column] : "";
                cells.Add(PptFileContent.BuildTableCell(text, table.HeaderBold));
            }
            rows.Add(PptFileContent.BuildTableRow(cells, rowHeight));
        }
        foreach (var row in table.Rows)
        {
            var cells = new List<XElement>();
            for (var column = 0; column < columnCount; column++)
            {
                var text = column < row.Count ? row[column] : "";
                cells.Add(PptFileContent.BuildTableCell(text, false));
            }
            rows.Add(PptFileContent.BuildTableRow(cells, rowHeight));
        }

        var id = NextId();
        var tableElement = new XElement(PptFileOoxml.A + "tbl",
            new XElement(PptFileOoxml.A + "tblPr",
                new XAttribute("firstRow", table.Headers.Count > 0 ? "1" : "0"),
                new XAttribute("bandRow", "1")),
            new XElement(PptFileOoxml.A + "tblGrid",
                widths.Select(columnWidth => new XElement(PptFileOoxml.A + "gridCol",
                    new XAttribute("w", PptFileOoxml.Inches(columnWidth))))),
            rows);
        var frame = new XElement(PptFileOoxml.P + "graphicFrame",
            new XElement(PptFileOoxml.P + "nvGraphicFramePr",
                new XElement(PptFileOoxml.P + "cNvPr",
                    new XAttribute("id", id), new XAttribute("name", $"Table {id}")),
                new XElement(PptFileOoxml.P + "cNvGraphicFramePr",
                    new XElement(PptFileOoxml.A + "graphicFrameLocks",
                        new XAttribute("noGrp", "1"))),
                new XElement(PptFileOoxml.P + "nvPr")),
            new XElement(PptFileOoxml.P + "xfrm",
                new XElement(PptFileOoxml.A + "off",
                    new XAttribute("x", PptFileOoxml.Inches(left)),
                    new XAttribute("y", PptFileOoxml.Inches(top))),
                new XElement(PptFileOoxml.A + "ext",
                    new XAttribute("cx", PptFileOoxml.Inches(width)),
                    new XAttribute("cy", PptFileOoxml.Inches(height)))),
            new XElement(PptFileOoxml.A + "graphic",
                new XElement(PptFileOoxml.A + "graphicData",
                    new XAttribute("uri",
                        "http://schemas.openxmlformats.org/drawingml/2006/table"),
                    tableElement)));
        _tree.Add(frame);
    }

    private void AddImage(PptSlideImage image)
    {
        byte[]? bytes = null;
        if (!string.IsNullOrWhiteSpace(image.Path))
        {
            if (!File.Exists(image.Path))
                throw new NativeRuntimeException($"图片文件不存在: {image.Path}");
            try { bytes = File.ReadAllBytes(image.Path); }
            catch (Exception error)
            {
                throw new NativeRuntimeException(
                    $"读取图片失败: {image.Path}: {error.Message}");
            }
        }
        else if (!string.IsNullOrWhiteSpace(image.Url))
        {
            bytes = PptFileContent.DownloadImage(image.Url!, out _);
            if (bytes is null)
                throw new NativeRuntimeException($"图片下载失败: {image.Url}");
        }
        if (bytes is null || bytes.Length == 0)
            throw new NativeRuntimeException("图片内容为空，无法插入");
        if (!PptFileContent.TryDetectImageType(bytes, out var imageType))
        {
            if (PptFileContent.LooksLikeSvg(bytes))
                throw new NativeRuntimeException(
                    "SVG 图片不受支持，请转换为 PNG/JPEG 后重试");
            throw new NativeRuntimeException(
                "无法识别的图片格式（仅支持 PNG/JPEG/GIF/BMP/TIFF/EMF/WMF）");
        }

        var width = image.Width ?? 4.0;
        var height = image.Height
                     ?? (PptFileContent.TryImageSize(bytes, out var pixelWidth, out var pixelHeight)
                         && pixelWidth > 0
                         ? width * pixelHeight / pixelWidth
                         : width * 0.75);
        var left = image.X ?? 0.5;
        var top = image.Y ?? 1.5;

        var imagePart = Part.AddImagePart(imageType);
        using (var stream = new MemoryStream(bytes))
            imagePart.FeedData(stream);
        var embedId = Part.GetIdOfPart(imagePart);

        var id = NextId();
        var cnvPr = new XElement(PptFileOoxml.P + "cNvPr",
            new XAttribute("id", id), new XAttribute("name", $"Picture {id}"));
        if (!string.IsNullOrWhiteSpace(image.HyperlinkUrl))
        {
            var hyperlinkId = PptFileContent.AddHyperlinkRelationship(Part, image.HyperlinkUrl!);
            cnvPr.Add(new XElement(PptFileOoxml.A + "hlinkClick",
                new XAttribute(PptFileOoxml.R + "id", hyperlinkId)));
        }
        var sourceRect = new XElement(PptFileOoxml.A + "srcRect");
        var hasCrop = false;
        if (image.CropLeft is not null)
        {
            sourceRect.SetAttributeValue("l", PptFileContent.Percent(image.CropLeft.Value));
            hasCrop = true;
        }
        if (image.CropTop is not null)
        {
            sourceRect.SetAttributeValue("t", PptFileContent.Percent(image.CropTop.Value));
            hasCrop = true;
        }
        if (image.CropRight is not null)
        {
            sourceRect.SetAttributeValue("r", PptFileContent.Percent(image.CropRight.Value));
            hasCrop = true;
        }
        if (image.CropBottom is not null)
        {
            sourceRect.SetAttributeValue("b", PptFileContent.Percent(image.CropBottom.Value));
            hasCrop = true;
        }

        var picture = new XElement(PptFileOoxml.P + "pic",
            new XElement(PptFileOoxml.P + "nvPicPr",
                cnvPr,
                new XElement(PptFileOoxml.P + "cNvPicPr",
                    new XElement(PptFileOoxml.A + "picLocks",
                        new XAttribute("noChangeAspect", "1"))),
                new XElement(PptFileOoxml.P + "nvPr")),
            new XElement(PptFileOoxml.P + "blipFill",
                new XElement(PptFileOoxml.A + "blip",
                    new XAttribute(PptFileOoxml.R + "embed", embedId)),
                hasCrop ? sourceRect : null,
                new XElement(PptFileOoxml.A + "stretch",
                    new XElement(PptFileOoxml.A + "fillRect"))),
            new XElement(PptFileOoxml.P + "spPr",
                new XElement(PptFileOoxml.A + "xfrm",
                    new XElement(PptFileOoxml.A + "off",
                        new XAttribute("x", PptFileOoxml.Inches(left)),
                        new XAttribute("y", PptFileOoxml.Inches(top))),
                    new XElement(PptFileOoxml.A + "ext",
                        new XAttribute("cx", PptFileOoxml.Inches(width)),
                        new XAttribute("cy", PptFileOoxml.Inches(height)))),
                new XElement(PptFileOoxml.A + "prstGeom",
                    new XAttribute("prst", "rect"),
                    new XElement(PptFileOoxml.A + "avLst"))));
        _tree.Add(picture);
    }

    private void AddShape(PptSlideShape shape)
    {
        var preset = shape.ShapeType switch
        {
            "rounded_rect" => "roundRect",
            "oval" => "ellipse",
            "arrow" => "rightArrow",
            "cylinder" => "can",
            "diamond" => "diamond",
            _ => "rect",
        };
        var id = NextId();
        var cnvPr = new XElement(PptFileOoxml.P + "cNvPr",
            new XAttribute("id", id), new XAttribute("name", $"Shape {id}"));
        if (!string.IsNullOrWhiteSpace(shape.HyperlinkUrl))
        {
            var hyperlinkId = PptFileContent.AddHyperlinkRelationship(Part, shape.HyperlinkUrl!);
            cnvPr.Add(new XElement(PptFileOoxml.A + "hlinkClick",
                new XAttribute(PptFileOoxml.R + "id", hyperlinkId)));
        }

        var shapeProperties = new XElement(PptFileOoxml.P + "spPr",
            new XElement(PptFileOoxml.A + "xfrm",
                new XElement(PptFileOoxml.A + "off",
                    new XAttribute("x", PptFileOoxml.Inches(shape.X ?? 0.5)),
                    new XAttribute("y", PptFileOoxml.Inches(shape.Y ?? 1.5))),
                new XElement(PptFileOoxml.A + "ext",
                    new XAttribute("cx", PptFileOoxml.Inches(shape.Width)),
                    new XAttribute("cy", PptFileOoxml.Inches(shape.Height)))),
            new XElement(PptFileOoxml.A + "prstGeom",
                new XAttribute("prst", preset),
                new XElement(PptFileOoxml.A + "avLst")));

        if (shape.NoFill)
            shapeProperties.Add(new XElement(PptFileOoxml.A + "noFill"));
        else if (!string.IsNullOrWhiteSpace(shape.BgColor))
            shapeProperties.Add(PptFileOoxml.SolidFill(shape.BgColor!));

        if (!string.IsNullOrWhiteSpace(shape.LineColor) || shape.LineWidth is not null
            || shape.LineDash)
        {
            var line = new XElement(PptFileOoxml.A + "ln");
            if (shape.LineWidth is not null)
                line.SetAttributeValue("w", PptFileOoxml.Points(shape.LineWidth.Value));
            if (!string.IsNullOrWhiteSpace(shape.LineColor))
                line.Add(PptFileOoxml.SolidFill(shape.LineColor!));
            if (shape.LineDash)
                line.Add(new XElement(PptFileOoxml.A + "prstDash",
                    new XAttribute("val", "dash")));
            shapeProperties.Add(line);
        }

        var textBody = new XElement(PptFileOoxml.P + "txBody");
        var bodyProperties = new XElement(PptFileOoxml.A + "bodyPr");
        if (shape.VerticalText)
        {
            bodyProperties.SetAttributeValue("vert", "vert270");
            bodyProperties.SetAttributeValue("wrap", "none");
        }
        else
        {
            bodyProperties.SetAttributeValue("wrap", "square");
        }
        textBody.Add(bodyProperties);
        textBody.Add(new XElement(PptFileOoxml.A + "lstStyle"));
        if (!string.IsNullOrEmpty(shape.Text))
        {
            var runProperties = new XElement(PptFileOoxml.A + "rPr",
                new XAttribute("sz", (shape.FontSize is > 0 ? shape.FontSize.Value : 12) * 100));
            if (shape.Bold == true)
                runProperties.SetAttributeValue("b", "1");
            if (!string.IsNullOrWhiteSpace(shape.FontColor))
                runProperties.Add(PptFileOoxml.SolidFill(shape.FontColor!));
            var paragraph = new XElement(PptFileOoxml.A + "p",
                new XElement(PptFileOoxml.A + "pPr",
                    new XAttribute("algn", "ctr")));
            PptFileContent.AppendTextContent(paragraph, shape.Text!, runProperties);
            textBody.Add(paragraph);
        }
        else
        {
            textBody.Add(new XElement(PptFileOoxml.A + "p"));
        }

        var element = new XElement(PptFileOoxml.P + "sp",
            new XElement(PptFileOoxml.P + "nvSpPr",
                cnvPr,
                new XElement(PptFileOoxml.P + "cNvSpPr",
                    new XElement(PptFileOoxml.A + "spLocks",
                        new XAttribute("noGrp", "1"))),
                new XElement(PptFileOoxml.P + "nvPr")),
            shapeProperties,
            PptFileContent.ShapeStyle(),
            textBody);
        _tree.Add(element);
    }

    private void AddConnector(PptSlideConnector connector, string style,
        bool tailArrow, bool headArrow)
    {
        var id = NextId();
        var element = PptFileContent.BuildConnectorElement(id, connector,
            style, tailArrow, headArrow);
        _tree.Add(element);
    }

    private void AddChart(PptChartDef chart, PptDefaults defaults)
    {
        var left = chart.X ?? defaults.Left;
        var top = chart.Y ?? defaults.Top;
        var width = chart.Width ?? defaults.Width;
        var height = chart.Height ?? defaults.Height;

        var chartPart = Part.AddNewPart<ChartPart>();
        // 图表工作簿：PowerPoint「编辑数据」依赖 chart 的 externalData 关系 + 内嵌 xlsx。
        chartPart.OpenXmlPackage.Features
            .Get<DocumentFormat.OpenXml.Features.IPartExtensionFeature>()
            ?.Register(PptFileContent.SpreadsheetContentType, ".xlsx");
        using var workbookStream = new MemoryStream(PptFileContent.BuildChartWorkbook(chart));
        var embeddedPart = chartPart.AddNewPart<EmbeddedPackagePart>(
            PptFileContent.SpreadsheetContentType);
        embeddedPart.FeedData(workbookStream);
        var externalDataId = chartPart.GetIdOfPart(embeddedPart);
        chartPart.ChartSpace = new XCharts.ChartSpace(
            PptFileContent.BuildChartSpace(chart, externalDataId)
                .ToString(SaveOptions.DisableFormatting));
        var relationshipId = Part.GetIdOfPart(chartPart);

        var id = NextId();
        var frame = new XElement(PptFileOoxml.P + "graphicFrame",
            new XElement(PptFileOoxml.P + "nvGraphicFramePr",
                new XElement(PptFileOoxml.P + "cNvPr",
                    new XAttribute("id", id), new XAttribute("name", $"Chart {id}")),
                new XElement(PptFileOoxml.P + "cNvGraphicFramePr",
                    new XElement(PptFileOoxml.A + "graphicFrameLocks",
                        new XAttribute("noGrp", "1"))),
                new XElement(PptFileOoxml.P + "nvPr")),
            new XElement(PptFileOoxml.P + "xfrm",
                new XElement(PptFileOoxml.A + "off",
                    new XAttribute("x", PptFileOoxml.Inches(left)),
                    new XAttribute("y", PptFileOoxml.Inches(top))),
                new XElement(PptFileOoxml.A + "ext",
                    new XAttribute("cx", PptFileOoxml.Inches(width)),
                    new XAttribute("cy", PptFileOoxml.Inches(height)))),
            new XElement(PptFileOoxml.A + "graphic",
                new XElement(PptFileOoxml.A + "graphicData",
                    new XAttribute("uri",
                        "http://schemas.openxmlformats.org/drawingml/2006/chart"),
                    new XElement(PptFileOoxml.C + "chart",
                        new XAttribute(XNamespace.Xmlns + "c", PptFileOoxml.C),
                        new XAttribute(PptFileOoxml.R + "id", relationshipId)))));
        _tree.Add(frame);
    }

    private uint NextId() => _nextId++;
}

internal static class PptFileContent
{
    // ── 新建幻灯片 ───────────────────────────────────────────────

    public static PptSlideWriter AddSlide(PresentationPart presentationPart,
        int layoutIndex, string? background)
    {
        var layouts = PptFileOoxml.LayoutParts(presentationPart);
        if (layouts.Count == 0)
            throw new NativeRuntimeException("演示文稿没有可用版式");
        if (layoutIndex < 0 || layoutIndex >= layouts.Count)
            throw new NativeRuntimeException(
                $"模板只有 {layouts.Count} 种版式，缺少索引 {layoutIndex} 的版式；"
                + "请改用包含该版式的模板");
        var layoutPart = layouts[layoutIndex];

        var slidePart = presentationPart.AddNewPart<SlidePart>();
        slidePart.AddPart(layoutPart);

        var root = new XElement(PptFileOoxml.P + "sld",
            new XAttribute(XNamespace.Xmlns + "a", PptFileOoxml.A),
            new XAttribute(XNamespace.Xmlns + "r", PptFileOoxml.R),
            new XAttribute(XNamespace.Xmlns + "p", PptFileOoxml.P),
            new XElement(PptFileOoxml.P + "cSld",
                new XElement(PptFileOoxml.P + "spTree",
                    new XElement(PptFileOoxml.P + "nvGrpSpPr",
                        new XElement(PptFileOoxml.P + "cNvPr",
                            new XAttribute("id", "1"), new XAttribute("name", "")),
                        new XElement(PptFileOoxml.P + "cNvGrpSpPr"),
                        new XElement(PptFileOoxml.P + "nvPr")),
                    new XElement(PptFileOoxml.P + "grpSpPr"))),
            new XElement(PptFileOoxml.P + "clrMapOvr",
                new XElement(PptFileOoxml.A + "masterClrMapping")));
        var document = new XDocument(root);
        ClonePlaceholders(layoutPart, document);

        var writer = new PptSlideWriter(slidePart, document);
        if (!string.IsNullOrWhiteSpace(background))
            writer.SetBackground(background!);
        writer.Save();

        RegisterSlideId(presentationPart, slidePart);
        return writer;
    }

    private static void RegisterSlideId(PresentationPart presentationPart,
        SlidePart slidePart)
    {
        var presentation = presentationPart.Presentation;
        var list = presentation.SlideIdList;
        if (list is null)
        {
            list = new XPresentation.SlideIdList();
            var slideSize = presentation.SlideSize;
            if (slideSize is not null)
                presentation.InsertBefore(list, slideSize);
            else
                presentation.AppendChild(list);
        }
        var nextId = 255U;
        foreach (var existing in list.Elements<XPresentation.SlideId>())
        {
            if (existing.Id?.Value > nextId)
                nextId = existing.Id.Value;
        }
        nextId++;
        list.Append(new XPresentation.SlideId
        {
            Id = nextId,
            RelationshipId = presentationPart.GetIdOfPart(slidePart),
        });
    }

    private static void ClonePlaceholders(SlideLayoutPart layoutPart, XDocument slide)
    {
        var layoutDocument = XDocument.Parse(layoutPart.SlideLayout.OuterXml);
        var masterDocument = layoutPart.SlideMasterPart?.SlideMaster is { } master
            ? XDocument.Parse(master.OuterXml)
            : null;
        var tree = PptFileOoxml.SlideTree(slide);
        uint id = 2;
        foreach (var source in PptFileOoxml.ShapeElements(layoutDocument).ToList())
        {
            var info = PptFileOoxml.PlaceholderInfo(source);
            if (info is null)
                continue;
            var type = info.Value.Type ?? "obj";
            if (type is "dt" or "ftr" or "sldNum")
                continue;

            var sourceXfrm = source.Element(PptFileOoxml.P + "spPr")
                ?.Element(PptFileOoxml.A + "xfrm");
            if (sourceXfrm is null)
            {
                var inherited = PptFileOoxml.FindPlaceholderShape(
                    masterDocument, info.Value.Type, info.Value.Idx);
                sourceXfrm = inherited?.Element(PptFileOoxml.P + "spPr")
                    ?.Element(PptFileOoxml.A + "xfrm");
            }
            var sourceShapeProperties = source.Element(PptFileOoxml.P + "spPr");
            var shapeProperties = sourceShapeProperties is null
                ? new XElement(PptFileOoxml.P + "spPr")
                : new XElement(sourceShapeProperties);
            shapeProperties.Element(PptFileOoxml.A + "xfrm")?.Remove();
            if (sourceXfrm is not null)
                shapeProperties.AddFirst(new XElement(sourceXfrm));

            var sourceCnvPr = source.Element(PptFileOoxml.P + "nvSpPr")
                ?.Element(PptFileOoxml.P + "cNvPr");
            var name = (string?)sourceCnvPr?.Attribute("name") ?? $"Placeholder {id}";
            var nvSpPr = new XElement(PptFileOoxml.P + "nvSpPr",
                new XElement(PptFileOoxml.P + "cNvPr",
                    new XAttribute("id", id), new XAttribute("name", name)),
                new XElement(PptFileOoxml.P + "cNvSpPr",
                    new XElement(PptFileOoxml.A + "spLocks",
                        new XAttribute("noGrp", "1"))),
                new XElement(PptFileOoxml.P + "nvPr",
                    new XElement(PptFileOoxml.P + "ph",
                        info.Value.Type is null
                            ? null
                            : new XAttribute("type", info.Value.Type),
                        new XAttribute("idx", info.Value.Idx))));

            var sourceTextBody = source.Element(PptFileOoxml.P + "txBody");
            XElement? textBody = null;
            if (sourceTextBody is not null)
            {
                textBody = new XElement(PptFileOoxml.P + "txBody");
                var bodyProperties = sourceTextBody.Element(PptFileOoxml.A + "bodyPr");
                textBody.Add(bodyProperties is null
                    ? new XElement(PptFileOoxml.A + "bodyPr")
                    : new XElement(bodyProperties));
                var listStyle = sourceTextBody.Element(PptFileOoxml.A + "lstStyle");
                if (listStyle is not null)
                    textBody.Add(new XElement(listStyle));
                textBody.Add(new XElement(PptFileOoxml.A + "p"));
            }

            tree.Add(new XElement(PptFileOoxml.P + "sp",
                nvSpPr, shapeProperties, textBody));
            id++;
        }
    }

    // ── 段落/文字 ────────────────────────────────────────────────

    public static XElement TextElement(string text) =>
        new(PptFileOoxml.A + "t",
            new XAttribute(XNamespace.Xml + "space", "preserve"), text);

    /// <summary>把含 \n 的文本写成 run + a:br 序列（同一段落内换行）。</summary>
    public static void AppendTextContent(XElement paragraph, string? text,
        XElement? runProperties)
    {
        var lines = (text ?? "").Split('\n');
        for (var index = 0; index < lines.Length; index++)
        {
            if (index > 0)
                paragraph.Add(new XElement(PptFileOoxml.A + "br"));
            if (lines[index].Length == 0)
                continue;
            var run = new XElement(PptFileOoxml.A + "r");
            if (runProperties is not null)
                run.Add(new XElement(runProperties));
            run.Add(TextElement(lines[index]));
            paragraph.Add(run);
        }
    }

    public static XElement BuildPlainParagraph(string text)
    {
        var paragraph = new XElement(PptFileOoxml.A + "p");
        AppendTextContent(paragraph, text, null);
        return paragraph;
    }

    public static XElement BuildParagraph(string text, int level, PptTextStyle? style,
        bool bullet, bool numbered)
    {
        level = Math.Clamp(level, 0, 8);
        var paragraph = new XElement(PptFileOoxml.A + "p");
        var properties = new XElement(PptFileOoxml.A + "pPr");
        if (level > 0)
            properties.SetAttributeValue("lvl", level);
        if (!string.IsNullOrEmpty(style?.Alignment))
        {
            properties.SetAttributeValue("algn", style!.Alignment switch
            {
                "center" => "ctr",
                "right" => "r",
                _ => "l",
            });
        }
        if (bullet || numbered)
        {
            properties.SetAttributeValue("marL", 342900L * (level + 1));
            properties.SetAttributeValue("indent", -342900);
            if (numbered)
                properties.Add(new XElement(PptFileOoxml.A + "buAutoNum",
                    new XAttribute("type", "arabicPeriod")));
            else
                properties.Add(new XElement(PptFileOoxml.A + "buChar",
                    new XAttribute("char", "•")));
        }
        if (properties.HasAttributes || properties.HasElements)
            paragraph.Add(properties);
        AppendTextContent(paragraph, text, BuildRunProperties(style));
        return paragraph;
    }

    public static XElement? BuildRunProperties(PptTextStyle? style)
    {
        if (style is null)
            return null;
        var properties = new XElement(PptFileOoxml.A + "rPr");
        if (style.FontSize is > 0)
            properties.SetAttributeValue("sz", style.FontSize.Value * 100);
        if (style.Bold is not null)
            properties.SetAttributeValue("b", style.Bold.Value ? "1" : "0");
        if (style.Italic is not null)
            properties.SetAttributeValue("i", style.Italic.Value ? "1" : "0");
        if (!string.IsNullOrWhiteSpace(style.Color))
            properties.Add(PptFileOoxml.SolidFill(style.Color!));
        if (!string.IsNullOrWhiteSpace(style.FontName))
        {
            properties.Add(new XElement(PptFileOoxml.A + "latin",
                new XAttribute("typeface", style.FontName)));
            properties.Add(new XElement(PptFileOoxml.A + "ea",
                new XAttribute("typeface", style.FontName)));
        }
        return properties.HasAttributes || properties.HasElements ? properties : null;
    }

    // ── 表格 ─────────────────────────────────────────────────────

    public static List<double> ComputeColumnWidths(List<double>? declared, int columns,
        double targetWidth)
    {
        if (declared is null || declared.Count == 0)
            return Enumerable.Repeat(targetWidth / columns, columns).ToList();
        var widths = declared.Take(columns).Select(width => Math.Max(width, 0.1)).ToList();
        if (widths.Count < columns)
        {
            var missing = columns - widths.Count;
            var remaining = Math.Max(targetWidth - widths.Sum(), 0.0);
            var fill = Math.Max(remaining / missing, 0.1);
            widths.AddRange(Enumerable.Repeat(fill, missing));
        }
        var total = widths.Sum();
        if (total > 0 && Math.Abs(total - targetWidth) > 0.01)
        {
            var factor = targetWidth / total;
            widths = widths.Select(width => width * factor).ToList();
        }
        return widths;
    }

    public static XElement BuildTableCell(string text, bool bold)
    {
        var paragraph = new XElement(PptFileOoxml.A + "p");
        AppendTextContent(paragraph, text,
            bold ? new XElement(PptFileOoxml.A + "rPr", new XAttribute("b", "1")) : null);
        return new XElement(PptFileOoxml.A + "tc",
            new XElement(PptFileOoxml.A + "txBody",
                new XElement(PptFileOoxml.A + "bodyPr"),
                new XElement(PptFileOoxml.A + "lstStyle"),
                paragraph),
            new XElement(PptFileOoxml.A + "tcPr"));
    }

    public static XElement BuildTableRow(IEnumerable<XElement> cells, double height) =>
        new(PptFileOoxml.A + "tr",
            new XAttribute("h", PptFileOoxml.Inches(height)), cells);

    // ── 形状/连接线公共部件 ─────────────────────────────────────

    public static XElement ShapeStyle() =>
        new(PptFileOoxml.P + "style",
            new XElement(PptFileOoxml.A + "lnRef", new XAttribute("idx", "2"),
                new XElement(PptFileOoxml.A + "schemeClr",
                    new XAttribute("val", "accent1"),
                    new XElement(PptFileOoxml.A + "shade", new XAttribute("val", "50000")))),
            new XElement(PptFileOoxml.A + "fillRef", new XAttribute("idx", "1"),
                new XElement(PptFileOoxml.A + "schemeClr",
                    new XAttribute("val", "accent1"))),
            new XElement(PptFileOoxml.A + "effectRef", new XAttribute("idx", "0"),
                new XElement(PptFileOoxml.A + "schemeClr",
                    new XAttribute("val", "accent1"))),
            new XElement(PptFileOoxml.A + "fontRef", new XAttribute("idx", "minor"),
                new XElement(PptFileOoxml.A + "schemeClr",
                    new XAttribute("val", "lt1"))));

    public static XElement BuildConnectorElement(uint id, PptSlideConnector connector,
        string style, bool tailArrow, bool headArrow)
    {
        var x1 = PptFileOoxml.Inches(connector.X1);
        var y1 = PptFileOoxml.Inches(connector.Y1);
        var x2 = PptFileOoxml.Inches(connector.X2);
        var y2 = PptFileOoxml.Inches(connector.Y2);
        var left = Math.Min(x1, x2);
        var top = Math.Min(y1, y2);
        var width = Math.Abs(x2 - x1);
        var height = Math.Abs(y2 - y1);
        var preset = style switch
        {
            "elbow" => "bentConnector3",
            "curve" => "curvedConnector3",
            _ => "line",
        };
        var transform = new XElement(PptFileOoxml.A + "xfrm",
            new XAttribute("flipH", x2 < x1 ? "1" : "0"),
            new XAttribute("flipV", y2 < y1 ? "1" : "0"),
            new XElement(PptFileOoxml.A + "off",
                new XAttribute("x", left), new XAttribute("y", top)),
            new XElement(PptFileOoxml.A + "ext",
                new XAttribute("cx", width), new XAttribute("cy", height)));

        var line = new XElement(PptFileOoxml.A + "ln",
            new XAttribute("w", PptFileOoxml.Points(connector.Width)));
        if (!string.IsNullOrWhiteSpace(connector.Color))
            line.Add(PptFileOoxml.SolidFill(connector.Color!));
        if (connector.Dash)
            line.Add(new XElement(PptFileOoxml.A + "prstDash",
                new XAttribute("val", "dash")));
        if (headArrow)
            line.Add(new XElement(PptFileOoxml.A + "headEnd",
                new XAttribute("type", "triangle"),
                new XAttribute("w", "med"), new XAttribute("len", "med")));
        if (tailArrow)
        {
            var arrowType = connector.ArrowheadStyle switch
            {
                "open" => "oval",
                "stealth" => "stealth",
                _ => "triangle",
            };
            line.Add(new XElement(PptFileOoxml.A + "tailEnd",
                new XAttribute("type", arrowType),
                new XAttribute("w", "med"), new XAttribute("len", "med")));
        }

        return new XElement(PptFileOoxml.P + "cxnSp",
            new XElement(PptFileOoxml.P + "nvCxnSpPr",
                new XElement(PptFileOoxml.P + "cNvPr",
                    new XAttribute("id", id), new XAttribute("name", $"Connector {id}")),
                new XElement(PptFileOoxml.P + "cNvCxnSpPr"),
                new XElement(PptFileOoxml.P + "nvPr")),
            new XElement(PptFileOoxml.P + "spPr",
                transform,
                new XElement(PptFileOoxml.A + "prstGeom",
                    new XAttribute("prst", preset),
                    new XElement(PptFileOoxml.A + "avLst")),
                line),
            ShapeStyle());
    }

    // ── 图表 ────────────────────────────────────────────────────

    public const string SpreadsheetContentType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    public static XElement BuildChartSpace(PptChartDef chart,
        string? externalDataRelationshipId = null)
    {
        var type = chart.ChartType is "column" or "bar" or "line" or "pie"
            or "doughnut" or "area" or "radar"
            ? chart.ChartType
            : "column";
        var series = chart.Series.Select((item, index) =>
            BuildSeries(item, chart.Categories, index,
                type == "line" || type == "radar")).ToList();
        const string categoryAxisId = "500001";
        const string valueAxisId = "500002";

        XElement plot;
        if (type is "pie" or "doughnut")
        {
            plot = type == "pie"
                ? new XElement(PptFileOoxml.C + "plotArea",
                    new XElement(PptFileOoxml.C + "layout"),
                    new XElement(PptFileOoxml.C + "pieChart",
                        new XElement(PptFileOoxml.C + "varyColors",
                            new XAttribute("val", "1")),
                        series,
                        new XElement(PptFileOoxml.C + "firstSliceAng",
                            new XAttribute("val", "0"))))
                : new XElement(PptFileOoxml.C + "plotArea",
                    new XElement(PptFileOoxml.C + "layout"),
                    new XElement(PptFileOoxml.C + "doughnutChart",
                        new XElement(PptFileOoxml.C + "varyColors",
                            new XAttribute("val", "1")),
                        series,
                        new XElement(PptFileOoxml.C + "holeSize",
                            new XAttribute("val", "50"))));
        }
        else
        {
            var chartElement = type switch
            {
                "bar" => new XElement(PptFileOoxml.C + "barChart",
                    new XElement(PptFileOoxml.C + "barDir",
                        new XAttribute("val", "bar")),
                    new XElement(PptFileOoxml.C + "grouping",
                        new XAttribute("val", "clustered")),
                    new XElement(PptFileOoxml.C + "varyColors",
                        new XAttribute("val", "0")),
                    series,
                    new XElement(PptFileOoxml.C + "gapWidth",
                        new XAttribute("val", "150")),
                    new XElement(PptFileOoxml.C + "axId",
                        new XAttribute("val", categoryAxisId)),
                    new XElement(PptFileOoxml.C + "axId",
                        new XAttribute("val", valueAxisId))),
                "line" => new XElement(PptFileOoxml.C + "lineChart",
                    new XElement(PptFileOoxml.C + "grouping",
                        new XAttribute("val", "standard")),
                    new XElement(PptFileOoxml.C + "varyColors",
                        new XAttribute("val", "0")),
                    series,
                    new XElement(PptFileOoxml.C + "marker",
                        new XAttribute("val", "1")),
                    new XElement(PptFileOoxml.C + "axId",
                        new XAttribute("val", categoryAxisId)),
                    new XElement(PptFileOoxml.C + "axId",
                        new XAttribute("val", valueAxisId))),
                "area" => new XElement(PptFileOoxml.C + "areaChart",
                    new XElement(PptFileOoxml.C + "grouping",
                        new XAttribute("val", "standard")),
                    new XElement(PptFileOoxml.C + "varyColors",
                        new XAttribute("val", "0")),
                    series,
                    new XElement(PptFileOoxml.C + "axId",
                        new XAttribute("val", categoryAxisId)),
                    new XElement(PptFileOoxml.C + "axId",
                        new XAttribute("val", valueAxisId))),
                "radar" => new XElement(PptFileOoxml.C + "radarChart",
                    new XElement(PptFileOoxml.C + "radarStyle",
                        new XAttribute("val", "marker")),
                    new XElement(PptFileOoxml.C + "varyColors",
                        new XAttribute("val", "0")),
                    series,
                    new XElement(PptFileOoxml.C + "axId",
                        new XAttribute("val", categoryAxisId)),
                    new XElement(PptFileOoxml.C + "axId",
                        new XAttribute("val", valueAxisId))),
                _ => new XElement(PptFileOoxml.C + "barChart",
                    new XElement(PptFileOoxml.C + "barDir",
                        new XAttribute("val", "col")),
                    new XElement(PptFileOoxml.C + "grouping",
                        new XAttribute("val", "clustered")),
                    new XElement(PptFileOoxml.C + "varyColors",
                        new XAttribute("val", "0")),
                    series,
                    new XElement(PptFileOoxml.C + "gapWidth",
                        new XAttribute("val", "150")),
                    new XElement(PptFileOoxml.C + "axId",
                        new XAttribute("val", categoryAxisId)),
                    new XElement(PptFileOoxml.C + "axId",
                        new XAttribute("val", valueAxisId))),
            };
            var horizontalBar = type == "bar";
            plot = new XElement(PptFileOoxml.C + "plotArea",
                new XElement(PptFileOoxml.C + "layout"),
                chartElement,
                BuildCategoryAxis(categoryAxisId, valueAxisId,
                    horizontalBar ? "l" : "b"),
                BuildValueAxis(valueAxisId, categoryAxisId,
                    horizontalBar ? "b" : "l"));
        }

        var chartElementRoot = new XElement(PptFileOoxml.C + "chart");
        if (!string.IsNullOrWhiteSpace(chart.Title))
            chartElementRoot.Add(BuildChartTitle(chart.Title!));
        chartElementRoot.Add(new XElement(PptFileOoxml.C + "autoTitleDeleted",
            new XAttribute("val", string.IsNullOrWhiteSpace(chart.Title) ? "1" : "0")));
        chartElementRoot.Add(plot);
        if (chart.HasLegend)
        {
            chartElementRoot.Add(new XElement(PptFileOoxml.C + "legend",
                new XElement(PptFileOoxml.C + "legendPos",
                    new XAttribute("val", "b")),
                new XElement(PptFileOoxml.C + "overlay",
                    new XAttribute("val", "0"))));
        }
        chartElementRoot.Add(new XElement(PptFileOoxml.C + "plotVisOnly",
            new XAttribute("val", "1")));
        chartElementRoot.Add(new XElement(PptFileOoxml.C + "dispBlanksAs",
            new XAttribute("val", "gap")));

        var chartSpace = new XElement(PptFileOoxml.C + "chartSpace",
            new XAttribute(XNamespace.Xmlns + "c", PptFileOoxml.C),
            new XAttribute(XNamespace.Xmlns + "a", PptFileOoxml.A),
            new XAttribute(XNamespace.Xmlns + "r", PptFileOoxml.R),
            chartElementRoot);
        if (!string.IsNullOrEmpty(externalDataRelationshipId))
        {
            chartSpace.Add(new XElement(PptFileOoxml.C + "externalData",
                new XAttribute(PptFileOoxml.R + "id", externalDataRelationshipId!),
                new XElement(PptFileOoxml.C + "autoUpdate",
                    new XAttribute("val", "0"))));
        }
        return chartSpace;
    }

    private static XElement BuildChartTitle(string title)
    {
        var paragraph = new XElement(PptFileOoxml.A + "p",
            new XElement(PptFileOoxml.A + "pPr",
                new XElement(PptFileOoxml.A + "defRPr",
                    new XAttribute("sz", "1200"))));
        AppendTextContent(paragraph, title,
            new XElement(PptFileOoxml.A + "rPr",
                new XAttribute("lang", "zh-CN"),
                new XAttribute("sz", "1200")));
        return new XElement(PptFileOoxml.C + "title",
            new XElement(PptFileOoxml.C + "tx",
                new XElement(PptFileOoxml.C + "rich",
                    new XElement(PptFileOoxml.A + "bodyPr"),
                    new XElement(PptFileOoxml.A + "lstStyle"),
                    paragraph)),
            new XElement(PptFileOoxml.C + "overlay",
                new XAttribute("val", "0")));
    }

    private static XElement BuildSeries(PptChartSeries series,
        List<string> categories, int index, bool markers)
    {
        var valueColumn = ColumnName(index + 1);
        var categoryPoints = categories.Select((category, pointIndex) =>
            new XElement(PptFileOoxml.C + "pt",
                new XAttribute("idx", pointIndex),
                new XElement(PptFileOoxml.C + "v", category))).ToList();
        var values = series.Values.Select((value, pointIndex) =>
            new XElement(PptFileOoxml.C + "pt",
                new XAttribute("idx", pointIndex),
                new XElement(PptFileOoxml.C + "v",
                    value.ToString(CultureInfo.InvariantCulture)))).ToList();
        var element = new XElement(PptFileOoxml.C + "ser",
            new XElement(PptFileOoxml.C + "idx", new XAttribute("val", index)),
            new XElement(PptFileOoxml.C + "order", new XAttribute("val", index)),
            new XElement(PptFileOoxml.C + "tx",
                new XElement(PptFileOoxml.C + "strRef",
                    new XElement(PptFileOoxml.C + "f",
                        $"Sheet1!${valueColumn}$1"),
                    new XElement(PptFileOoxml.C + "strCache",
                        new XElement(PptFileOoxml.C + "ptCount",
                            new XAttribute("val", "1")),
                        new XElement(PptFileOoxml.C + "pt",
                            new XAttribute("idx", "0"),
                            new XElement(PptFileOoxml.C + "v", series.Name))))));
        if (markers)
        {
            element.Add(new XElement(PptFileOoxml.C + "marker",
                new XElement(PptFileOoxml.C + "symbol",
                    new XAttribute("val", "circle")),
                new XElement(PptFileOoxml.C + "size",
                    new XAttribute("val", "5"))));
        }
        element.Add(new XElement(PptFileOoxml.C + "cat",
            new XElement(PptFileOoxml.C + "strRef",
                new XElement(PptFileOoxml.C + "f",
                    $"Sheet1!$A$2:$A${categories.Count + 1}"),
                new XElement(PptFileOoxml.C + "strCache",
                    new XElement(PptFileOoxml.C + "ptCount",
                        new XAttribute("val", categoryPoints.Count)),
                    categoryPoints))));
        if (values.Count > 0)
        {
            element.Add(new XElement(PptFileOoxml.C + "val",
                new XElement(PptFileOoxml.C + "numRef",
                    new XElement(PptFileOoxml.C + "f",
                        $"Sheet1!${valueColumn}$2:${valueColumn}${values.Count + 1}"),
                    new XElement(PptFileOoxml.C + "numCache",
                        new XElement(PptFileOoxml.C + "formatCode", "General"),
                        new XElement(PptFileOoxml.C + "ptCount",
                            new XAttribute("val", values.Count)),
                        values))));
        }
        else
        {
            // 空 series 没有可引用的工作表区间，保留空 numLit 缓存。
            element.Add(new XElement(PptFileOoxml.C + "val",
                new XElement(PptFileOoxml.C + "numLit",
                    new XElement(PptFileOoxml.C + "formatCode", "General"),
                    new XElement(PptFileOoxml.C + "ptCount",
                        new XAttribute("val", "0")))));
        }
        return element;
    }

    /// <summary>按 0 基列号生成 Excel 列名（0=A, 26=AA）。</summary>
    private static string ColumnName(int index)
    {
        var name = "";
        index++;
        while (index > 0)
        {
            index--;
            name = (char)('A' + index % 26) + name;
            index /= 26;
        }
        return name;
    }

    /// <summary>生成图表内嵌工作簿：第 1 行 series 名，A 列 categories，其余列为数值。</summary>
    public static byte[] BuildChartWorkbook(PptChartDef chart)
    {
        using var stream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(
                   stream, SpreadsheetDocumentType.Workbook, autoSave: true))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new XSpreadsheet.Workbook();
            var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
            sharedStringPart.SharedStringTable = new XSpreadsheet.SharedStringTable();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

            var sharedStrings = new List<string>();
            var sharedStringIndex = new Dictionary<string, int>(StringComparer.Ordinal);
            var sharedStringCount = 0;
            int SharedString(string value)
            {
                sharedStringCount++;
                if (sharedStringIndex.TryGetValue(value, out var existing))
                    return existing;
                var index = sharedStrings.Count;
                sharedStrings.Add(value);
                sharedStringIndex[value] = index;
                return index;
            }

            var rows = new List<XSpreadsheet.Row>();
            var header = new XSpreadsheet.Row { RowIndex = 1 };
            for (var seriesIndex = 0; seriesIndex < chart.Series.Count; seriesIndex++)
                header.Append(SharedStringCell(
                    $"{ColumnName(seriesIndex + 1)}1",
                    SharedString(chart.Series[seriesIndex].Name)));
            rows.Add(header);

            // 行数取 categories 与最长 series 的较大值，避免 values 长于 categories 时
            // chart 的 numRef 公式指向工作簿中不存在的单元格。
            var dataRowCount = Math.Max(chart.Categories.Count,
                chart.Series.Count == 0 ? 0 : chart.Series.Max(series => series.Values.Count));
            for (var rowIndex = 0; rowIndex < dataRowCount; rowIndex++)
            {
                var row = new XSpreadsheet.Row { RowIndex = (uint)(rowIndex + 2) };
                if (rowIndex < chart.Categories.Count)
                {
                    row.Append(SharedStringCell($"A{rowIndex + 2}",
                        SharedString(chart.Categories[rowIndex])));
                }
                for (var seriesIndex = 0; seriesIndex < chart.Series.Count; seriesIndex++)
                {
                    var series = chart.Series[seriesIndex];
                    if (rowIndex >= series.Values.Count)
                        continue;
                    row.Append(new XSpreadsheet.Cell
                    {
                        CellReference = $"{ColumnName(seriesIndex + 1)}{rowIndex + 2}",
                        CellValue = new XSpreadsheet.CellValue(
                            series.Values[rowIndex].ToString(CultureInfo.InvariantCulture)),
                    });
                }
                rows.Add(row);
            }

            foreach (var value in sharedStrings)
            {
                var item = new XSpreadsheet.SharedStringItem(
                    new XSpreadsheet.Text(value));
                item.Text!.Space = SpaceProcessingModeValues.Preserve;
                sharedStringPart.SharedStringTable.Append(item);
            }
            sharedStringPart.SharedStringTable.Count = (uint)sharedStringCount;
            sharedStringPart.SharedStringTable.UniqueCount = (uint)sharedStrings.Count;

            var sheetData = new XSpreadsheet.SheetData();
            foreach (var row in rows)
                sheetData.Append(row);
            var lastColumn = ColumnName(Math.Max(0, chart.Series.Count));
            worksheetPart.Worksheet = new XSpreadsheet.Worksheet(
                new XSpreadsheet.SheetDimension
                {
                    Reference = $"A1:{lastColumn}{dataRowCount + 1}",
                },
                sheetData);
            workbookPart.Workbook.Append(new XSpreadsheet.Sheets(
                new XSpreadsheet.Sheet
                {
                    Name = "Sheet1",
                    SheetId = 1U,
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                }));
            sharedStringPart.SharedStringTable.Save();
            worksheetPart.Worksheet.Save();
            workbookPart.Workbook.Save();
            document.Save();
        }
        return stream.ToArray();
    }

    private static XSpreadsheet.Cell SharedStringCell(string reference, int index) =>
        new()
        {
            CellReference = reference,
            DataType = XSpreadsheet.CellValues.SharedString,
            CellValue = new XSpreadsheet.CellValue(
                index.ToString(CultureInfo.InvariantCulture)),
        };

    private static XElement BuildCategoryAxis(string axisId, string crossAxisId,
        string position) =>
        new(PptFileOoxml.C + "catAx",
            new XElement(PptFileOoxml.C + "axId", new XAttribute("val", axisId)),
            new XElement(PptFileOoxml.C + "scaling",
                new XElement(PptFileOoxml.C + "orientation",
                    new XAttribute("val", "minMax"))),
            new XElement(PptFileOoxml.C + "delete", new XAttribute("val", "0")),
            new XElement(PptFileOoxml.C + "axPos", new XAttribute("val", position)),
            new XElement(PptFileOoxml.C + "numFmt",
                new XAttribute("formatCode", "General"),
                new XAttribute("sourceLinked", "1")),
            new XElement(PptFileOoxml.C + "majorTickMark",
                new XAttribute("val", "out")),
            new XElement(PptFileOoxml.C + "minorTickMark",
                new XAttribute("val", "none")),
            new XElement(PptFileOoxml.C + "tickLblPos",
                new XAttribute("val", "nextTo")),
            new XElement(PptFileOoxml.C + "crossAx",
                new XAttribute("val", crossAxisId)),
            new XElement(PptFileOoxml.C + "crosses",
                new XAttribute("val", "autoZero")),
            new XElement(PptFileOoxml.C + "auto", new XAttribute("val", "1")),
            new XElement(PptFileOoxml.C + "lblAlgn", new XAttribute("val", "ctr")),
            new XElement(PptFileOoxml.C + "lblOffset", new XAttribute("val", "100")),
            new XElement(PptFileOoxml.C + "noMultiLvlLbl",
                new XAttribute("val", "0")));

    private static XElement BuildValueAxis(string axisId, string crossAxisId,
        string position) =>
        new(PptFileOoxml.C + "valAx",
            new XElement(PptFileOoxml.C + "axId", new XAttribute("val", axisId)),
            new XElement(PptFileOoxml.C + "scaling",
                new XElement(PptFileOoxml.C + "orientation",
                    new XAttribute("val", "minMax"))),
            new XElement(PptFileOoxml.C + "delete", new XAttribute("val", "0")),
            new XElement(PptFileOoxml.C + "axPos", new XAttribute("val", position)),
            new XElement(PptFileOoxml.C + "majorGridlines"),
            new XElement(PptFileOoxml.C + "numFmt",
                new XAttribute("formatCode", "General"),
                new XAttribute("sourceLinked", "1")),
            new XElement(PptFileOoxml.C + "majorTickMark",
                new XAttribute("val", "out")),
            new XElement(PptFileOoxml.C + "minorTickMark",
                new XAttribute("val", "none")),
            new XElement(PptFileOoxml.C + "tickLblPos",
                new XAttribute("val", "nextTo")),
            new XElement(PptFileOoxml.C + "crossAx",
                new XAttribute("val", crossAxisId)),
            new XElement(PptFileOoxml.C + "crosses",
                new XAttribute("val", "autoZero")),
            new XElement(PptFileOoxml.C + "crossBetween",
                new XAttribute("val", "between")));

    // ── 图片/超链接/数学工具 ────────────────────────────────────

    private static readonly HttpClient HttpClient = new()
    {
        Timeout = TimeSpan.FromSeconds(30),
    };

    public const int MaxDownloadedImageBytes = 20 * 1024 * 1024;

    public static byte[]? DownloadImage(string url, out string extension)
    {
        extension = ".png";
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri)
            || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
            return null;
        try
        {
            using var response = HttpClient.GetAsync(uri).GetAwaiter().GetResult();
            if (!response.IsSuccessStatusCode)
                return null;
            if (response.Content.Headers.ContentLength is > MaxDownloadedImageBytes)
                return null;
            using var stream = response.Content.ReadAsStreamAsync().GetAwaiter().GetResult();
            using var buffer = new MemoryStream();
            var chunk = new byte[81920];
            int read;
            while ((read = stream.Read(chunk, 0, chunk.Length)) > 0)
            {
                if (buffer.Length + read > MaxDownloadedImageBytes)
                    return null;
                buffer.Write(chunk, 0, read);
            }
            var bytes = buffer.ToArray();
            if (bytes.Length == 0)
                return null;
            var contentType = response.Content.Headers.ContentType?.MediaType ?? "";
            extension = contentType.ToLowerInvariant() switch
            {
                "image/jpeg" => ".jpg",
                "image/gif" => ".gif",
                "image/bmp" => ".bmp",
                "image/tiff" => ".tiff",
                "image/x-emf" => ".emf",
                "image/x-wmf" => ".wmf",
                _ => Path.GetExtension(uri.AbsolutePath).ToLowerInvariant() is { Length: > 0 } ext
                    ? ext
                    : ".png",
            };
            return bytes;
        }
        catch (Exception)
        {
            return null;
        }
    }

    public static bool LooksLikeSvg(byte[] bytes)
    {
        var length = Math.Min(bytes.Length, 1024);
        var head = System.Text.Encoding.UTF8.GetString(bytes, 0, length)
            .TrimStart('\uFEFF', ' ', '\r', '\n', '\t')
            .ToLowerInvariant();
        return head.StartsWith("<svg", StringComparison.Ordinal)
               || (head.StartsWith("<?xml", StringComparison.Ordinal)
                   && head.Contains("<svg", StringComparison.Ordinal));
    }

    /// <summary>按文件魔数判定图片格式（不信任扩展名）。</summary>
    public static bool TryDetectImageType(byte[] bytes, out PartTypeInfo type)
    {
        type = ImagePartType.Png;
        if (bytes.Length >= 8 && bytes[0] == 0x89 && bytes[1] == 0x50
            && bytes[2] == 0x4E && bytes[3] == 0x47 && bytes[4] == 0x0D
            && bytes[5] == 0x0A && bytes[6] == 0x1A && bytes[7] == 0x0A)
        {
            type = ImagePartType.Png;
            return true;
        }
        if (bytes.Length >= 3 && bytes[0] == 0xFF && bytes[1] == 0xD8
            && bytes[2] == 0xFF)
        {
            type = ImagePartType.Jpeg;
            return true;
        }
        if (bytes.Length >= 6 && bytes[0] == (byte)'G' && bytes[1] == (byte)'I'
            && bytes[2] == (byte)'F' && bytes[3] == (byte)'8'
            && (bytes[4] == (byte)'7' || bytes[4] == (byte)'9')
            && bytes[5] == (byte)'a')
        {
            type = ImagePartType.Gif;
            return true;
        }
        if (bytes.Length >= 2 && bytes[0] == (byte)'B' && bytes[1] == (byte)'M')
        {
            type = ImagePartType.Bmp;
            return true;
        }
        if (bytes.Length >= 4
            && ((bytes[0] == 0x49 && bytes[1] == 0x49 && bytes[2] == 0x2A
                 && bytes[3] == 0x00)
                || (bytes[0] == 0x4D && bytes[1] == 0x4D && bytes[2] == 0x00
                    && bytes[3] == 0x2A)))
        {
            type = ImagePartType.Tiff;
            return true;
        }
        if (bytes.Length >= 44 && bytes[0] == 0x01 && bytes[1] == 0x00
            && bytes[2] == 0x00 && bytes[3] == 0x00
            && bytes[40] == 0x20 && bytes[41] == 0x45
            && bytes[42] == 0x4D && bytes[43] == 0x46)
        {
            type = ImagePartType.Emf;
            return true;
        }
        if (bytes.Length >= 4 && bytes[0] == 0xD7 && bytes[1] == 0xCD
            && bytes[2] == 0xC6 && bytes[3] == 0x9A)
        {
            type = ImagePartType.Wmf;
            return true;
        }
        if (bytes.Length >= 6 && bytes[0] == 0x01 && bytes[1] == 0x00
            && bytes[2] == 0x09 && bytes[3] == 0x00)
        {
            type = ImagePartType.Wmf;
            return true;
        }
        return false;
    }

    public static bool TryImageSize(byte[] bytes, out int width, out int height)
    {
        width = 0;
        height = 0;
        if (bytes.Length >= 24 && bytes[0] == 0x89 && bytes[1] == 0x50
            && bytes[2] == 0x4E && bytes[3] == 0x47)
        {
            width = (bytes[16] << 24) | (bytes[17] << 16) | (bytes[18] << 8) | bytes[19];
            height = (bytes[20] << 24) | (bytes[21] << 16) | (bytes[22] << 8) | bytes[23];
            return width > 0 && height > 0;
        }
        if (bytes.Length >= 10 && bytes[0] == 'G' && bytes[1] == 'I' && bytes[2] == 'F')
        {
            width = bytes[6] | (bytes[7] << 8);
            height = bytes[8] | (bytes[9] << 8);
            return width > 0 && height > 0;
        }
        if (bytes.Length >= 26 && bytes[0] == 'B' && bytes[1] == 'M')
        {
            width = BitConverter.ToInt32(bytes, 18);
            height = Math.Abs(BitConverter.ToInt32(bytes, 22));
            return width > 0 && height > 0;
        }
        if (bytes.Length >= 4 && bytes[0] == 0xFF && bytes[1] == 0xD8)
        {
            var offset = 2;
            while (offset + 9 < bytes.Length)
            {
                if (bytes[offset] != 0xFF)
                {
                    offset++;
                    continue;
                }
                var marker = bytes[offset + 1];
                if (marker is 0xD8 or 0xD9 or 0x01 or >= 0xD0 and <= 0xD7)
                {
                    offset += 2;
                    continue;
                }
                var length = (bytes[offset + 2] << 8) | bytes[offset + 3];
                if (length < 2)
                    return false;
                var isStartOfFrame = marker is >= 0xC0 and <= 0xC3
                    or >= 0xC5 and <= 0xC7 or >= 0xC9 and <= 0xCB
                    or >= 0xCD and <= 0xCF;
                if (isStartOfFrame)
                {
                    height = (bytes[offset + 5] << 8) | bytes[offset + 6];
                    width = (bytes[offset + 7] << 8) | bytes[offset + 8];
                    return width > 0 && height > 0;
                }
                offset += 2 + length;
            }
        }
        return false;
    }

    public static string AddHyperlinkRelationship(SlidePart slidePart, string url)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri))
            throw new NativeUsageException($"超链接地址不是合法绝对 URL: {url}");
        return slidePart.AddHyperlinkRelationship(uri, true).Id;
    }

    public static long Percent(double fraction) =>
        Math.Clamp((long)Math.Round(fraction * 100000, MidpointRounding.AwayFromZero),
            0, 100000);

    public static uint MaxShapeId(XDocument document)
    {
        uint maximum = 1;
        foreach (var element in document.Descendants()
                     .Where(element => element.Name.LocalName == "cNvPr"))
        {
            if (uint.TryParse((string?)element.Attribute("id"), out var id) && id > maximum)
                maximum = id;
        }
        return maximum;
    }

    public static string HexMixWhite(string hex, double weight)
    {
        var (r, g, b) = ParseHex(hex);
        var w = Math.Clamp(weight, 0, 1);
        return $"{RoundChannel(r + (255 - r) * w):X2}{RoundChannel(g + (255 - g) * w):X2}"
               + $"{RoundChannel(b + (255 - b) * w):X2}";
    }

    public static string HexMixBlack(string hex, double weight)
    {
        var (r, g, b) = ParseHex(hex);
        var w = Math.Clamp(weight, 0, 1);
        return $"{RoundChannel(r * (1 - w)):X2}{RoundChannel(g * (1 - w)):X2}"
               + $"{RoundChannel(b * (1 - w)):X2}";
    }

    private static (int R, int G, int B) ParseHex(string hex)
    {
        var text = (hex ?? "").Trim().TrimStart('#');
        if (text.Length != 6)
            return (0, 0, 0);
        return (Convert.ToInt32(text[..2], 16),
            Convert.ToInt32(text.Substring(2, 2), 16),
            Convert.ToInt32(text.Substring(4, 2), 16));
    }

    private static int RoundChannel(double value) =>
        (int)Math.Round(Math.Clamp(value, 0, 255), MidpointRounding.AwayFromZero);

    public static XElement AddShadow(XElement shapeProperties, double blurPoints,
        double distancePoints, int opacityPercent)
    {
        shapeProperties.Element(PptFileOoxml.A + "effectLst")?.Remove();
        shapeProperties.Add(new XElement(PptFileOoxml.A + "effectLst",
            new XElement(PptFileOoxml.A + "outerShdw",
                new XAttribute("blurRad", PptFileOoxml.Points(blurPoints)),
                new XAttribute("dist", PptFileOoxml.Points(distancePoints)),
                new XAttribute("dir", "5400000"),
                new XAttribute("rotWithShape", "0"),
                new XElement(PptFileOoxml.A + "srgbClr",
                    new XAttribute("val", "000000"),
                    new XElement(PptFileOoxml.A + "alpha",
                        new XAttribute("val", opacityPercent * 1000))))));
        return shapeProperties;
    }

    private static readonly Regex ChineseRuns = new(
        @"([\u4e00-\u9fff\u3400-\u4dbf\u3000-\u303f\uff00-\uffef]+)",
        RegexOptions.Compiled);

    public static IEnumerable<(string Text, bool Chinese)> SplitMixedFont(string text)
    {
        foreach (var segment in ChineseRuns.Split(text))
        {
            if (segment.Length == 0)
                continue;
            yield return (segment, ChineseRuns.IsMatch(segment)
                && ChineseRuns.Match(segment).Value == segment);
        }
    }

    public static XElement BuildMixedFontParagraph(string text, string fontCn,
        string fontEn, int fontSize, bool bold, string color)
    {
        var paragraph = new XElement(PptFileOoxml.A + "p",
            new XElement(PptFileOoxml.A + "pPr",
                new XAttribute("algn", "ctr")));
        var lines = (text ?? "").Split('\n');
        for (var lineIndex = 0; lineIndex < lines.Length; lineIndex++)
        {
            if (lineIndex > 0)
                paragraph.Add(new XElement(PptFileOoxml.A + "br"));
            foreach (var (segment, chinese) in SplitMixedFont(lines[lineIndex]))
            {
                var font = chinese ? fontCn : fontEn;
                var properties = new XElement(PptFileOoxml.A + "rPr",
                    new XAttribute("sz", fontSize * 100),
                    new XAttribute("lang", chinese ? "zh-CN" : "en-US"));
                if (bold)
                    properties.SetAttributeValue("b", "1");
                properties.Add(PptFileOoxml.SolidFill(color));
                properties.Add(new XElement(PptFileOoxml.A + "latin",
                    new XAttribute("typeface", font)));
                properties.Add(new XElement(PptFileOoxml.A + "ea",
                    new XAttribute("typeface", font)));
                paragraph.Add(new XElement(PptFileOoxml.A + "r",
                    properties, TextElement(segment)));
            }
        }
        return paragraph;
    }
}
