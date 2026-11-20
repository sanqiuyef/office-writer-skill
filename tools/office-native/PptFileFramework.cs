using System.Globalization;
using System.Text.Json;
using System.Xml.Linq;
using XPresentation = DocumentFormat.OpenXml.Presentation;

namespace OfficeNative;

/// <summary>框架图生成：阶段标签 + 节点 + 箭头（与 ppt_writer.create_framework_diagram 对齐）。</summary>
internal static class PptFileFramework
{
    public static Dictionary<string, object?> CreateDiagram(JsonElement p)
    {
        var outputPath = NativeCommon.RequireString(p, "output_path");
        PptFileCommands.EnsurePptExtension(outputPath, "输出文件必须使用 .pptx 扩展名");
        var diagram = PptModels.ParseFrameworkDiagram(p);
        var resolved = PptFileCommands.ResolveOutput(outputPath, p);
        var scheme = PptModels.ColorSchemes.GetValueOrDefault(diagram.ColorScheme)
                     ?? PptModels.ColorSchemes["professional-blue"];
        var arrowColor = string.IsNullOrWhiteSpace(diagram.ArrowColor)
            ? "94A3B8" : diagram.ArrowColor;
        var work = PptFileOoxml.WorkPath(resolved);
        try
        {
            PptFileOoxml.CreateScratch(work, 13.333, 7.5);
            uint slideCount;
            using (var document = PptFileOoxml.Open(work, true))
            {
                var presentationPart = PptFileOoxml.MainPart(document);
                var writer = PptFileContent.AddSlide(presentationPart, 6,
                    scheme.BackgroundColor);
                slideCount = 1;
                BuildSlide(writer, diagram, scheme, arrowColor);
                writer.Save();
                document.Save();
            }

            var backupPath = PptFileOoxml.ReplaceWithBackup(work, resolved);
            var empty = diagram.Stages.Count == 0;
            var fileSize = empty ? 0 : new FileInfo(resolved).Length;
            var message = empty
                ? "空框架图已创建"
                : $"框架图生成成功: {diagram.Stages.Count} 个阶段, "
                  + $"{diagram.NodeCount} 个节点";
            return new Dictionary<string, object?>
            {
                ["success"] = true,
                ["file_path"] = Path.GetFullPath(resolved),
                ["file_size"] = fileSize,
                ["slides_count"] = (int)slideCount,
                ["message"] = message,
                ["backup_path"] = backupPath,
            };
        }
        finally
        {
            PptFileOoxml.EnsureCleanup(work);
        }
    }

    private static void BuildSlide(PptSlideWriter writer, PptFrameworkDiagram diagram,
        PptColorScheme scheme, string arrowColor)
    {
        var accentBarY = 0.62;
        var titleYOffset = 0.2;
        if (!string.IsNullOrEmpty(diagram.Title))
        {
            writer.Append(TextBox(writer.TakeId(), "Title", 0.8, 0.15, 11.5, 0.45,
                wrap: true,
                [
                    PptFileContent.BuildMixedFontParagraph(diagram.Title,
                        diagram.FontCn, diagram.FontEn, diagram.TitleFontSize,
                        bold: true, scheme.TitleColor),
                ]));
            titleYOffset = 0.7;
        }
        if (!string.IsNullOrEmpty(diagram.Subtitle))
        {
            writer.Append(TextBox(writer.TakeId(), "Subtitle", 0.8, 0.5, 11.5, 0.3,
                wrap: true,
                [
                    SimpleParagraph(diagram.Subtitle!, 12, scheme.BodyColor,
                        scheme.BodyFont, bold: false, italic: false),
                ]));
            titleYOffset = 0.95;
            accentBarY = 0.9;
        }

        if (!string.IsNullOrEmpty(diagram.Title))
        {
            const double barWidth = 11.5 * 0.6;
            var bar = AutoShape(writer.TakeId(), "Accent Bar", "rect",
                0.8 + (11.5 - barWidth) / 2, accentBarY, barWidth, 0.04,
                fill: scheme.AccentColor, noFill: false,
                lineColor: null, lineWidth: null, noLine: true,
                adjustment: null, textBody: null);
            writer.Append(bar);
        }

        var stageCount = diagram.Stages.Count;
        if (stageCount == 0)
            return;

        const double contentLeft = 1.15;
        const double contentRight = 13.333 - 0.3;
        const double contentWidth = contentRight - contentLeft;
        var usableHeight = 7.5 - titleYOffset - 0.3;
        var stageUnitHeight = usableHeight / stageCount - diagram.StageGap;

        var positions = new Dictionary<string, (double X, double Y, double Width,
            double Height)>(StringComparer.Ordinal);

        for (var stageIndex = 0; stageIndex < diagram.Stages.Count; stageIndex++)
        {
            var stage = diagram.Stages[stageIndex];
            var stageY = titleYOffset
                         + stageIndex * (stageUnitHeight + diagram.StageGap);
            var stageColor = string.IsNullOrWhiteSpace(stage.Color)
                ? scheme.AccentColor : stage.Color;

            var background = AutoShape(writer.TakeId(), "Stage Background", "roundRect",
                contentLeft + 0.55, stageY, contentWidth - 0.55, stageUnitHeight,
                fill: PptFileContent.HexMixWhite(stageColor, 0.96), noFill: false,
                lineColor: PptFileContent.HexMixWhite(stageColor, 0.88),
                lineWidth: 0.75, noLine: false, adjustment: null, textBody: null);
            PptFileContent.AddShadow(
                background.Element(PptFileOoxml.P + "spPr")!, 10, 2, 25);
            writer.Append(background);

            const double pillWidth = 1.25;
            const double pillHeight = 0.38;
            var pillBody = TextBody(wrap: false, leftInset: 0.02, rightInset: 0.02,
                autoFit: false,
                paragraphs:
                [
                    PptFileContent.BuildMixedFontParagraph(stage.Label,
                        diagram.FontCn, diagram.FontEn, 12, bold: true, "FFFFFF"),
                ]);
            var pill = AutoShape(writer.TakeId(), "Stage Label", "roundRect",
                0.25, stageY + stageUnitHeight / 2 - pillHeight / 2,
                pillWidth, pillHeight,
                fill: stageColor, noFill: false,
                lineColor: PptFileContent.HexMixBlack(stageColor, 0.2),
                lineWidth: 0.75, noLine: false, adjustment: 0.5, textBody: pillBody);
            PptFileContent.AddShadow(
                pill.Element(PptFileOoxml.P + "spPr")!, 8, 3, 50);
            writer.Append(pill);

            var nodes = stage.Nodes;
            if (nodes.Count == 0)
                continue;

            var nodeWidth = diagram.NodeWidth;
            var nodeHeight = diagram.NodeHeight;
            var totalNodesWidth = nodes.Count * nodeWidth + (nodes.Count - 1) * 0.3;
            var startX = contentLeft + 0.6
                         + Math.Max(0, (contentWidth - 0.6 - totalNodesWidth) / 2);
            var nodeCenterY = stageY + stageUnitHeight / 2 - nodeHeight / 2;

            for (var nodeIndex = 0; nodeIndex < nodes.Count; nodeIndex++)
            {
                var node = nodes[nodeIndex];
                var nodeX = startX + nodeIndex * (nodeWidth + 0.3);
                var nodeColor = string.IsNullOrWhiteSpace(node.FillColor)
                    ? stageColor : node.FillColor;
                var preset = NodePreset(node.Style);

                var paragraphs = new List<XElement>();
                var lines = node.Text.Replace("\\n", "\n").Split('\n');
                for (var lineIndex = 0; lineIndex < lines.Length; lineIndex++)
                {
                    var line = lines[lineIndex];
                    if (string.IsNullOrWhiteSpace(line) && lineIndex != 0)
                        continue;
                    paragraphs.Add(PptFileContent.BuildMixedFontParagraph(
                        line, diagram.FontCn, diagram.FontEn,
                        node.FontSize > 0 ? node.FontSize : 14,
                        node.Bold, "FFFFFF"));
                }
                if (paragraphs.Count == 0)
                    paragraphs.Add(SimpleParagraph("", node.FontSize, "FFFFFF",
                        diagram.FontEn, false, false));

                var nodeBody = TextBody(wrap: true, leftInset: 0.04, rightInset: 0.04,
                    topInset: 0.02, bottomInset: 0.02, autoFit: true,
                    paragraphs: paragraphs);
                var shape = AutoShape(writer.TakeId(), $"Node {node.Id}", preset,
                    nodeX, nodeCenterY, nodeWidth, nodeHeight,
                    fill: nodeColor, noFill: false,
                    lineColor: PptFileContent.HexMixBlack(nodeColor, 0.25),
                    lineWidth: 1.0, noLine: false, adjustment: null,
                    textBody: nodeBody);
                PptFileContent.AddShadow(
                    shape.Element(PptFileOoxml.P + "spPr")!, 8, 3, 50);
                writer.Append(shape);

                if (nodes.Count >= 2 && node.Style != "terminator")
                {
                    const double badgeDiameter = 0.24;
                    var badgeBody = TextBody(wrap: false, leftInset: 0, rightInset: 0,
                        topInset: 0, bottomInset: 0, autoFit: false,
                        paragraphs: [
                            new XElement(PptFileOoxml.A + "p",
                                new XElement(PptFileOoxml.A + "pPr",
                                    new XAttribute("algn", "ctr")),
                                new XElement(PptFileOoxml.A + "r",
                                    new XElement(PptFileOoxml.A + "rPr",
                                        new XAttribute("sz", "900"),
                                        new XAttribute("b", "1"),
                                        PptFileOoxml.SolidFill(nodeColor),
                                        new XElement(PptFileOoxml.A + "latin",
                                            new XAttribute("typeface", "Consolas"))),
                                    PptFileContent.TextElement((nodeIndex + 1)
                                        .ToString(CultureInfo.InvariantCulture)))),
                        ]);
                    var badge = AutoShape(writer.TakeId(), "Node Badge", "ellipse",
                        nodeX - 0.08, nodeCenterY - 0.08,
                        badgeDiameter, badgeDiameter,
                        fill: "FFFFFF", noFill: false, lineColor: nodeColor,
                        lineWidth: 1.0, noLine: false, adjustment: null,
                        textBody: badgeBody);
                    PptFileContent.AddShadow(
                        badge.Element(PptFileOoxml.P + "spPr")!, 4, 1.5, 35);
                    writer.Append(badge);
                }

                positions[node.Id] = (nodeX, nodeCenterY, nodeWidth, nodeHeight);

                if (nodeIndex > 0)
                {
                    var previousId = nodes[nodeIndex - 1].Id;
                    if (positions.TryGetValue(previousId, out var previous))
                    {
                        var connector = new PptSlideConnector
                        {
                            X1 = previous.X + previous.Width,
                            Y1 = previous.Y + previous.Height / 2,
                            X2 = nodeX,
                            Y2 = nodeCenterY + nodeHeight / 2,
                            Color = arrowColor,
                            Width = 1.5,
                            Arrowhead = true,
                            ArrowheadStyle = "triangle",
                        };
                        writer.Append(PptFileContent.BuildConnectorElement(
                            writer.TakeId(), connector, "straight",
                            tailArrow: true, headArrow: false));
                    }
                }
            }
        }

        foreach (var edge in diagram.Edges)
        {
            if (!positions.TryGetValue(edge.FromId, out var from)
                || !positions.TryGetValue(edge.ToId, out var to))
                continue;
            var fromX = from.X + from.Width / 2;
            var fromY = from.Y + from.Height;
            var toX = to.X + to.Width / 2;
            var toY = to.Y;
            var style = edge.EdgeStyle is "elbow" or "curve" ? edge.EdgeStyle : "straight";
            var connector = new PptSlideConnector
            {
                X1 = fromX,
                Y1 = fromY,
                X2 = toX,
                Y2 = toY,
                Color = arrowColor,
                Width = 1.5,
                Arrowhead = true,
                ArrowheadStyle = "triangle",
            };
            writer.Append(PptFileContent.BuildConnectorElement(
                writer.TakeId(), connector, style,
                tailArrow: true, headArrow: false));

            if (!string.IsNullOrEmpty(edge.Label))
            {
                var labelX = (fromX + toX) / 2 - 0.5;
                var labelY = (fromY + toY) / 2 - 0.15;
                writer.Append(TextBox(writer.TakeId(), "Edge Label",
                    labelX, labelY, 1.0, 0.3, wrap: true,
                    [
                        SimpleParagraph(edge.Label!, 9, arrowColor, null,
                            bold: false, italic: true),
                    ]));
            }
        }
    }

    private static string NodePreset(string style) => style switch
    {
        "rounded_rect" => "roundRect",
        "ellipse" => "ellipse",
        "diamond" => "diamond",
        "process" => "flowChartProcess",
        "decision" => "flowChartDecision",
        "data" => "flowChartInputOutput",
        "document" => "flowChartDocument",
        "terminator" => "flowChartTerminator",
        "predefined" => "flowChartPredefinedProcess",
        "cylinder" => "can",
        "parallelogram" => "parallelogram",
        "chevron" => "chevron",
        "right_arrow" => "rightArrow",
        "left_arrow" => "leftArrow",
        "callout" => "wedgeRoundRectCallout",
        "cloud" => "cloud",
        _ => "rect",
    };

    private static XElement TextBox(uint id, string name, double x, double y,
        double width, double height, bool wrap, IEnumerable<XElement> paragraphs)
    {
        return new XElement(PptFileOoxml.P + "sp",
            new XElement(PptFileOoxml.P + "nvSpPr",
                new XElement(PptFileOoxml.P + "cNvPr",
                    new XAttribute("id", id), new XAttribute("name", $"{name} {id}")),
                new XElement(PptFileOoxml.P + "cNvSpPr",
                    new XAttribute("txBox", "1")),
                new XElement(PptFileOoxml.P + "nvPr")),
            new XElement(PptFileOoxml.P + "spPr",
                Transform(x, y, width, height),
                new XElement(PptFileOoxml.A + "prstGeom",
                    new XAttribute("prst", "rect"),
                    new XElement(PptFileOoxml.A + "avLst")),
                new XElement(PptFileOoxml.A + "noFill")),
            new XElement(PptFileOoxml.P + "txBody",
                new XElement(PptFileOoxml.A + "bodyPr",
                    new XAttribute("wrap", wrap ? "square" : "none")),
                new XElement(PptFileOoxml.A + "lstStyle"),
                paragraphs));
    }

    private static XElement AutoShape(uint id, string name, string preset,
        double x, double y, double width, double height,
        string? fill, bool noFill, string? lineColor, double? lineWidth,
        bool noLine, double? adjustment, XElement? textBody)
    {
        var geometry = new XElement(PptFileOoxml.A + "prstGeom",
            new XAttribute("prst", preset),
            adjustment is null
                ? new XElement(PptFileOoxml.A + "avLst")
                : new XElement(PptFileOoxml.A + "avLst",
                    new XElement(PptFileOoxml.A + "gd",
                        new XAttribute("name", "adj"),
                        new XAttribute("fmla",
                            $"val {(int)Math.Round(adjustment.Value * 100000)}"))));
        var shapeProperties = new XElement(PptFileOoxml.P + "spPr",
            Transform(x, y, width, height),
            geometry);
        if (noFill)
            shapeProperties.Add(new XElement(PptFileOoxml.A + "noFill"));
        else if (!string.IsNullOrEmpty(fill))
            shapeProperties.Add(PptFileOoxml.SolidFill(fill!));

        if (noLine)
        {
            shapeProperties.Add(new XElement(PptFileOoxml.A + "ln",
                new XElement(PptFileOoxml.A + "noFill")));
        }
        else if (!string.IsNullOrEmpty(lineColor) || lineWidth is not null)
        {
            var line = new XElement(PptFileOoxml.A + "ln");
            if (lineWidth is not null)
                line.SetAttributeValue("w", PptFileOoxml.Points(lineWidth.Value));
            if (!string.IsNullOrEmpty(lineColor))
                line.Add(PptFileOoxml.SolidFill(lineColor!));
            shapeProperties.Add(line);
        }

        var shape = new XElement(PptFileOoxml.P + "sp",
            new XElement(PptFileOoxml.P + "nvSpPr",
                new XElement(PptFileOoxml.P + "cNvPr",
                    new XAttribute("id", id), new XAttribute("name", $"{name} {id}")),
                new XElement(PptFileOoxml.P + "cNvSpPr",
                    new XElement(PptFileOoxml.A + "spLocks",
                        new XAttribute("noGrp", "1"))),
                new XElement(PptFileOoxml.P + "nvPr")),
            shapeProperties);
        if (textBody is not null)
        {
            shape.Add(textBody);
        }
        else
        {
            shape.Add(new XElement(PptFileOoxml.P + "txBody",
                new XElement(PptFileOoxml.A + "bodyPr"),
                new XElement(PptFileOoxml.A + "lstStyle"),
                new XElement(PptFileOoxml.A + "p")));
        }
        return shape;
    }

    private static XElement TextBody(bool wrap, double? leftInset, double? rightInset,
        double? topInset = null, double? bottomInset = null, bool autoFit = false,
        IEnumerable<XElement>? paragraphs = null)
    {
        var body = new XElement(PptFileOoxml.A + "bodyPr",
            new XAttribute("wrap", wrap ? "square" : "none"));
        if (leftInset is not null)
            body.SetAttributeValue("lIns", PptFileOoxml.Inches(leftInset.Value));
        if (rightInset is not null)
            body.SetAttributeValue("rIns", PptFileOoxml.Inches(rightInset.Value));
        if (topInset is not null)
            body.SetAttributeValue("tIns", PptFileOoxml.Inches(topInset.Value));
        if (bottomInset is not null)
            body.SetAttributeValue("bIns", PptFileOoxml.Inches(bottomInset.Value));
        if (autoFit)
            body.Add(new XElement(PptFileOoxml.A + "normAutofit"));
        return new XElement(PptFileOoxml.P + "txBody",
            body,
            new XElement(PptFileOoxml.A + "lstStyle"),
            paragraphs);
    }

    private static XElement Transform(double x, double y, double width, double height) =>
        new(PptFileOoxml.A + "xfrm",
            new XElement(PptFileOoxml.A + "off",
                new XAttribute("x", PptFileOoxml.Inches(x)),
                new XAttribute("y", PptFileOoxml.Inches(y))),
            new XElement(PptFileOoxml.A + "ext",
                new XAttribute("cx", PptFileOoxml.Inches(width)),
                new XAttribute("cy", PptFileOoxml.Inches(height))));

    private static XElement SimpleParagraph(string text, int fontSize, string color,
        string? font, bool bold, bool italic)
    {
        var properties = new XElement(PptFileOoxml.A + "rPr",
            new XAttribute("sz", fontSize * 100));
        if (bold)
            properties.SetAttributeValue("b", "1");
        if (italic)
            properties.SetAttributeValue("i", "1");
        properties.Add(PptFileOoxml.SolidFill(color));
        if (!string.IsNullOrEmpty(font))
        {
            properties.Add(new XElement(PptFileOoxml.A + "latin",
                new XAttribute("typeface", font)));
            properties.Add(new XElement(PptFileOoxml.A + "ea",
                new XAttribute("typeface", font)));
        }
        return new XElement(PptFileOoxml.A + "p",
            new XElement(PptFileOoxml.A + "pPr", new XAttribute("algn", "ctr")),
            new XElement(PptFileOoxml.A + "r",
                properties, PptFileContent.TextElement(text)));
    }
}
