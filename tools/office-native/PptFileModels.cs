using System.Globalization;
using System.Text.Json;

namespace OfficeNative;

// ── 与 office_engine/models.py 对齐的 PowerPoint 文件级模型 ──────────────

internal sealed class PptTextStyle
{
    public string? FontName { get; init; }
    public int? FontSize { get; init; }
    public bool? Bold { get; init; }
    public bool? Italic { get; init; }
    public string? Color { get; init; }
    public string? Alignment { get; init; }
}

internal sealed class PptSlideTable
{
    public List<string> Headers { get; init; } = [];
    public List<List<string>> Rows { get; init; } = [];
    public double? X { get; init; }
    public double? Y { get; init; }
    public double? Width { get; init; }
    public double? Height { get; init; }
    public List<double>? ColumnWidths { get; init; }
    public bool HeaderBold { get; init; } = true;
}

internal sealed class PptSlideImage
{
    public string? Path { get; init; }
    public string? Url { get; init; }
    public double? X { get; init; }
    public double? Y { get; init; }
    public double? Width { get; init; }
    public double? Height { get; init; }
    public double? CropLeft { get; init; }
    public double? CropRight { get; init; }
    public double? CropTop { get; init; }
    public double? CropBottom { get; init; }
    public string? HyperlinkUrl { get; init; }
}

internal sealed class PptSlideShape
{
    public string ShapeType { get; init; } = "rect";
    public string? Text { get; init; }
    public string? HyperlinkUrl { get; init; }
    public double? X { get; init; }
    public double? Y { get; init; }
    public double Width { get; init; }
    public double Height { get; init; }
    public string? BgColor { get; init; }
    public string? FontColor { get; init; }
    public int? FontSize { get; init; }
    public bool? Bold { get; init; }
    public bool NoFill { get; init; }
    public string? LineColor { get; init; }
    public double? LineWidth { get; init; }
    public bool LineDash { get; init; }
    public bool VerticalText { get; init; }
}

internal sealed class PptSlideConnector
{
    public double X1 { get; init; }
    public double Y1 { get; init; }
    public double X2 { get; init; }
    public double Y2 { get; init; }
    public string? Color { get; init; }
    public double Width { get; init; } = 1.5;
    public bool Dash { get; init; }
    public bool Arrowhead { get; init; } = true;
    public string ArrowheadStyle { get; init; } = "triangle";
    public bool BeginArrowhead { get; init; }
}

internal sealed class PptChartSeries
{
    public string Name { get; init; } = "";
    public List<double> Values { get; init; } = [];
}

internal sealed class PptChartDef
{
    public string ChartType { get; init; } = "column";
    public List<string> Categories { get; init; } = [];
    public List<PptChartSeries> Series { get; init; } = [];
    public bool HasLegend { get; init; } = true;
    public string? Title { get; init; }
    public double? X { get; init; }
    public double? Y { get; init; }
    public double? Width { get; init; }
    public double? Height { get; init; }
}

internal sealed class PptSlideItem
{
    public string Type { get; init; } = "paragraph";
    public string? Text { get; init; }
    public int Level { get; init; }
    public PptTextStyle? Style { get; init; }
    public PptSlideTable? Table { get; init; }
    public PptSlideImage? Image { get; init; }
    public PptSlideShape? Shape { get; init; }
    public PptSlideConnector? Connector { get; init; }
    public List<string>? Items { get; init; }
    public PptChartDef? Chart { get; init; }
}

internal sealed class PptSlideDef
{
    public string Layout { get; init; } = "content";
    public string? Title { get; init; }
    public string? Subtitle { get; init; }
    public List<PptSlideItem>? Body { get; init; }
    public List<PptSlideItem>? Left { get; init; }
    public List<PptSlideItem>? Right { get; init; }
    public string? Notes { get; init; }
    public string? BgColor { get; init; }
    public double? SlideWidth { get; init; }
    public double? SlideHeight { get; init; }
    public string? FooterText { get; init; }
    public bool ShowSlideNumber { get; init; }
}

/// <summary>幻灯片节（仅创建时写入 p14:sectionLst）。</summary>
internal sealed class PptSectionDef
{
    public string Name { get; init; } = "";
    public int SlideStart { get; init; }
}

/// <summary>放映设置（仅创建时写入 ppt/presProps.xml 的 p:showPr）。</summary>
internal sealed class PptShowSettingsDef
{
    public string ShowType { get; init; } = "presenter";
    public bool Loop { get; init; }
    public bool ShowAnimation { get; init; } = true;
    public bool ShowNarration { get; init; } = true;
    public bool UseTimings { get; init; } = true;
    public int? RangeStart { get; init; }
    public int? RangeEnd { get; init; }
}

internal sealed class PptFrameworkNode
{
    public string Id { get; init; } = "";
    public string Text { get; init; } = "";
    public string Style { get; init; } = "rounded_rect";
    public string? FillColor { get; init; }
    public int FontSize { get; init; } = 12;
    public bool Bold { get; init; } = true;
}

internal sealed class PptFrameworkEdge
{
    public string FromId { get; init; } = "";
    public string ToId { get; init; } = "";
    public string? Label { get; init; }
    public string EdgeStyle { get; init; } = "straight";
}

internal sealed class PptFrameworkStage
{
    public string Label { get; init; } = "";
    public string? Color { get; init; }
    public List<PptFrameworkNode> Nodes { get; init; } = [];
}

internal sealed class PptFrameworkDiagram
{
    public string Title { get; init; } = "";
    public string? Subtitle { get; init; }
    public List<PptFrameworkStage> Stages { get; init; } = [];
    public List<PptFrameworkEdge> Edges { get; init; } = [];
    public string ColorScheme { get; init; } = "professional-blue";
    public double NodeWidth { get; init; } = 1.4;
    public double NodeHeight { get; init; } = 0.55;
    public double StageGap { get; init; } = 0.15;
    public string? ArrowColor { get; init; }
    public string FontCn { get; init; } = "宋体";
    public string FontEn { get; init; } = "Times New Roman";
    public int FontSize { get; init; } = 16;
    public int TitleFontSize { get; init; } = 24;

    public int NodeCount => Stages.Sum(stage => stage.Nodes.Count);
}

internal sealed class PptHyperlinkDef
{
    public int Slide { get; init; }
    public string Url { get; init; } = "";
    public int? ShapeIndex { get; init; }
    public string? ShapeText { get; init; }
    public string? TextMatch { get; init; }
    public bool MatchCase { get; init; }
}

internal sealed class PptColorScheme
{
    public required string Name { get; init; }
    public required string DisplayName { get; init; }
    public string? Description { get; init; }
    public required string TitleColor { get; init; }
    public required string BodyColor { get; init; }
    public required string AccentColor { get; init; }
    public required string BackgroundColor { get; init; }
    public required string ShapeFillColor { get; init; }
    public required string TableHeaderColor { get; init; }
    public string TitleFont { get; init; } = "微软雅黑";
    public string BodyFont { get; init; } = "微软雅黑";
    public int TitleSize { get; init; } = 32;
    public int BodySize { get; init; } = 18;
}

internal static class PptModels
{
    public static readonly Dictionary<string, int> LayoutMap = new(StringComparer.Ordinal)
    {
        ["title"] = 0,
        ["text"] = 1,
        ["content"] = 1,
        ["section"] = 2,
        ["two_content"] = 3,
        ["comparison"] = 4,
        ["title_only"] = 5,
        ["blank"] = 6,
        ["content_caption"] = 7,
        ["picture_caption"] = 8,
    };

    public static readonly string[] LayoutNames =
    [
        "Title Slide", "Title and Content", "Section Header", "Two Content",
        "Comparison", "Title Only", "Blank", "Content with Caption",
        "Picture with Caption",
    ];

    // 12 套预设配色（与 ppt_writer._PRESET_COLOR_SCHEMES 完全一致）
    public static readonly Dictionary<string, PptColorScheme> ColorSchemes = BuildSchemes();

    public static int GetLayoutIndex(string? layout) =>
        layout is not null && LayoutMap.TryGetValue(layout, out var index) ? index : 1;

    public static Dictionary<string, object?> SchemePreview(PptColorScheme scheme) => new()
    {
        ["title_color"] = scheme.TitleColor,
        ["body_color"] = scheme.BodyColor,
        ["accent_color"] = scheme.AccentColor,
        ["background_color"] = scheme.BackgroundColor,
    };

    public static List<Dictionary<string, object?>> ListSchemes() =>
        ColorSchemes.Values.Select(scheme => new Dictionary<string, object?>
        {
            ["name"] = scheme.Name,
            ["display_name"] = scheme.DisplayName,
            ["description"] = scheme.Description,
            ["preview"] = SchemePreview(scheme),
        }).ToList();

    // ── JSON 解析 ────────────────────────────────────────────────

    public static List<PptSlideDef> ParseSlides(JsonElement p, string name = "slides")
    {
        if (p.ValueKind != JsonValueKind.Object || !p.TryGetProperty(name, out var value)
            || value.ValueKind != JsonValueKind.Array)
            throw new NativeUsageException($"缺少参数 {name}（JSON 数组）");
        var slides = new List<PptSlideDef>();
        foreach (var item in value.EnumerateArray())
            slides.Add(ParseSlide(item));
        return slides;
    }

    public static List<PptSlideDef> ParseOptionalSlides(JsonElement p, string name)
    {
        if (p.ValueKind != JsonValueKind.Object || !p.TryGetProperty(name, out var value)
            || value.ValueKind != JsonValueKind.Array)
            return [];
        return value.EnumerateArray().Select(ParseSlide).ToList();
    }

    private static PptSlideDef ParseSlide(JsonElement element)
    {
        if (element.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException("slides 的每一项必须是 JSON 对象");
        var layout = Str(element, "layout") ?? "content";
        if (!LayoutMap.ContainsKey(layout))
            throw new NativeUsageException(
                $"未知 layout: {layout}（支持 {string.Join("/", LayoutMap.Keys)}）");
        return new PptSlideDef
        {
            Layout = layout,
            Title = Str(element, "title"),
            Subtitle = Str(element, "subtitle"),
            Body = ParseItems(element, "body"),
            Left = ParseItems(element, "left"),
            Right = ParseItems(element, "right"),
            Notes = Str(element, "notes"),
            BgColor = ColorStr(element, "bg_color"),
            SlideWidth = Dbl(element, "slide_width"),
            SlideHeight = Dbl(element, "slide_height"),
            FooterText = Str(element, "footer_text"),
            ShowSlideNumber = Bool(element, "show_slide_number") ?? false,
        };
    }

    /// <summary>读取顶层 sections（可缺省）；slide_start 必须是 1 基正整数。</summary>
    public static List<PptSectionDef> ParseSections(JsonElement p)
    {
        if (p.ValueKind != JsonValueKind.Object || !p.TryGetProperty("sections", out var value)
            || value.ValueKind == JsonValueKind.Null)
            return [];
        if (value.ValueKind != JsonValueKind.Array)
            throw new NativeUsageException("sections 必须是 JSON 数组");
        var sections = new List<PptSectionDef>();
        var index = 0;
        foreach (var element in value.EnumerateArray())
        {
            index++;
            if (element.ValueKind != JsonValueKind.Object)
                throw new NativeUsageException("sections 的每一项必须是 JSON 对象");
            var name = Str(element, "name");
            if (string.IsNullOrWhiteSpace(name))
                throw new NativeUsageException($"sections[{index - 1}] 缺少 name");
            var slideStart = Int(element, "slide_start");
            if (slideStart is null)
                throw new NativeUsageException($"sections[{index - 1}] 缺少 slide_start");
            if (slideStart < 1)
                throw new NativeUsageException(
                    $"sections[{index - 1}].slide_start 必须 >= 1（1 起），收到 {slideStart}");
            sections.Add(new PptSectionDef
            {
                Name = name!.Trim(),
                SlideStart = slideStart!.Value,
            });
        }
        return sections;
    }

    /// <summary>读取顶层 show_settings（可缺省）；未知 show_type 直接报错。</summary>
    public static PptShowSettingsDef? ParseShowSettings(JsonElement p)
    {
        if (p.ValueKind != JsonValueKind.Object
            || !p.TryGetProperty("show_settings", out var value)
            || value.ValueKind == JsonValueKind.Null)
            return null;
        if (value.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException("show_settings 必须是 JSON 对象");
        var showType = Str(value, "show_type") ?? "presenter";
        if (showType is not ("presenter" or "browse" or "kiosk"))
            throw new NativeUsageException(
                $"未知 show_settings.show_type: {showType}（支持 presenter/browse/kiosk）");
        var rangeStart = Int(value, "range_start");
        var rangeEnd = Int(value, "range_end");
        if (rangeStart is < 1)
            throw new NativeUsageException("show_settings.range_start 必须 >= 1（1 起）");
        if (rangeEnd is < 1)
            throw new NativeUsageException("show_settings.range_end 必须 >= 1（1 起）");
        return new PptShowSettingsDef
        {
            ShowType = showType,
            Loop = Bool(value, "loop") ?? false,
            ShowAnimation = Bool(value, "show_animation") ?? true,
            ShowNarration = Bool(value, "show_narration") ?? true,
            UseTimings = Bool(value, "use_timings") ?? true,
            RangeStart = rangeStart,
            RangeEnd = rangeEnd,
        };
    }

    private static List<PptSlideItem>? ParseItems(JsonElement parent, string name)
    {
        if (!parent.TryGetProperty(name, out var value) || value.ValueKind != JsonValueKind.Array)
            return null;
        return value.EnumerateArray().Select(ParseItem).ToList();
    }

    private static PptSlideItem ParseItem(JsonElement element)
    {
        if (element.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException("SlideItem 必须是 JSON 对象");
        var type = Str(element, "type") ?? "paragraph";
        if (type is not ("paragraph" or "table" or "image" or "shape" or "connector"
            or "bullet_list" or "numbered_list" or "chart"))
            throw new NativeUsageException(
                $"未知 SlideItem.type: {type}（支持 paragraph/table/image/shape/connector/"
                + "bullet_list/numbered_list/chart）");
        return new PptSlideItem
        {
            Type = type,
            Text = Str(element, "text"),
            Level = Math.Clamp(Int(element, "level") ?? 0, 0, 8),
            Style = ParseStyle(Obj(element, "style")),
            Table = ParseTable(Obj(element, "table")),
            Image = ParseImage(Obj(element, "image")),
            Shape = ParseShape(Obj(element, "shape")),
            Connector = ParseConnector(Obj(element, "connector")),
            Items = StringList(element, "items"),
            Chart = ParseChart(Obj(element, "chart")),
        };
    }

    private static PptTextStyle? ParseStyle(JsonElement? element)
    {
        if (element is null)
            return null;
        var alignment = Str(element.Value, "alignment");
        if (alignment is not null && alignment is not ("left" or "center" or "right"))
            throw new NativeUsageException(
                $"TextStyle.alignment 必须是 left/center/right，收到 {alignment}");
        return new PptTextStyle
        {
            FontName = Str(element.Value, "font_name"),
            FontSize = Int(element.Value, "font_size"),
            Bold = Bool(element.Value, "bold"),
            Italic = Bool(element.Value, "italic"),
            Color = ColorStr(element.Value, "color"),
            Alignment = alignment,
        };
    }

    private static PptSlideTable? ParseTable(JsonElement? element)
    {
        if (element is null)
            return null;
        var rows = new List<List<string>>();
        if (element.Value.TryGetProperty("rows", out var rowsElement)
            && rowsElement.ValueKind == JsonValueKind.Array)
        {
            foreach (var row in rowsElement.EnumerateArray())
                rows.Add(row.ValueKind == JsonValueKind.Array
                    ? row.EnumerateArray().Select(JsonToString).ToList()
                    : []);
        }
        return new PptSlideTable
        {
            Headers = StringList(element.Value, "headers"),
            Rows = rows,
            X = Dbl(element.Value, "x"),
            Y = Dbl(element.Value, "y"),
            Width = Dbl(element.Value, "width"),
            Height = Dbl(element.Value, "height"),
            ColumnWidths = DoubleList(element.Value, "column_widths"),
            HeaderBold = Bool(element.Value, "header_bold") ?? true,
        };
    }

    private static PptSlideImage? ParseImage(JsonElement? element)
    {
        if (element is null)
            return null;
        var path = Str(element.Value, "path");
        var url = Str(element.Value, "url");
        if (string.IsNullOrWhiteSpace(path) && string.IsNullOrWhiteSpace(url))
            throw new NativeUsageException("SlideImageDef 必须提供 path 或 url 之一");
        if (!string.IsNullOrWhiteSpace(path) && !string.IsNullOrWhiteSpace(url))
            throw new NativeUsageException("SlideImageDef 不能同时提供 path 和 url，请二选一");
        return new PptSlideImage
        {
            Path = path,
            Url = url,
            X = Dbl(element.Value, "x"),
            Y = Dbl(element.Value, "y"),
            Width = Dbl(element.Value, "width"),
            Height = Dbl(element.Value, "height"),
            CropLeft = Dbl(element.Value, "crop_left"),
            CropRight = Dbl(element.Value, "crop_right"),
            CropTop = Dbl(element.Value, "crop_top"),
            CropBottom = Dbl(element.Value, "crop_bottom"),
            HyperlinkUrl = Str(element.Value, "hyperlink_url"),
        };
    }

    private static PptSlideShape? ParseShape(JsonElement? element)
    {
        if (element is null)
            return null;
        if (Dbl(element.Value, "width") is null || Dbl(element.Value, "height") is null)
            throw new NativeUsageException("SlideShapeDef 必须提供 width 和 height（英寸）");
        var shapeType = Str(element.Value, "shape_type") ?? "rect";
        if (shapeType is not ("rect" or "rounded_rect" or "oval" or "arrow"
            or "cylinder" or "diamond"))
            throw new NativeUsageException(
                $"未知 shape_type: {shapeType}（支持 rect/rounded_rect/oval/arrow/"
                + "cylinder/diamond）");
        return new PptSlideShape
        {
            ShapeType = shapeType,
            Text = Str(element.Value, "text"),
            HyperlinkUrl = Str(element.Value, "hyperlink_url"),
            X = Dbl(element.Value, "x"),
            Y = Dbl(element.Value, "y"),
            Width = Dbl(element.Value, "width")!.Value,
            Height = Dbl(element.Value, "height")!.Value,
            BgColor = ColorStr(element.Value, "bg_color"),
            FontColor = ColorStr(element.Value, "font_color"),
            FontSize = Int(element.Value, "font_size"),
            Bold = Bool(element.Value, "bold"),
            NoFill = Bool(element.Value, "no_fill") ?? false,
            LineColor = ColorStr(element.Value, "line_color"),
            LineWidth = Dbl(element.Value, "line_width"),
            LineDash = Bool(element.Value, "line_dash") ?? false,
            VerticalText = Bool(element.Value, "vertical_text") ?? false,
        };
    }

    private static PptSlideConnector? ParseConnector(JsonElement? element)
    {
        if (element is null)
            return null;
        if (Dbl(element.Value, "x1") is null || Dbl(element.Value, "y1") is null
            || Dbl(element.Value, "x2") is null || Dbl(element.Value, "y2") is null)
            throw new NativeUsageException("SlideConnectorDef 必须提供 x1/y1/x2/y2（英寸）");
        var arrowheadStyle = Str(element.Value, "arrowhead_style") ?? "triangle";
        if (arrowheadStyle is not ("triangle" or "open" or "stealth"))
            throw new NativeUsageException(
                $"未知 arrowhead_style: {arrowheadStyle}（支持 triangle/open/stealth）");
        return new PptSlideConnector
        {
            X1 = Dbl(element.Value, "x1")!.Value,
            Y1 = Dbl(element.Value, "y1")!.Value,
            X2 = Dbl(element.Value, "x2")!.Value,
            Y2 = Dbl(element.Value, "y2")!.Value,
            Color = ColorStr(element.Value, "color"),
            Width = Dbl(element.Value, "width") ?? 1.5,
            Dash = Bool(element.Value, "dash") ?? false,
            Arrowhead = Bool(element.Value, "arrowhead") ?? true,
            ArrowheadStyle = arrowheadStyle,
            BeginArrowhead = Bool(element.Value, "begin_arrowhead") ?? false,
        };
    }

    private static PptChartDef? ParseChart(JsonElement? element)
    {
        if (element is null)
            return null;
        var categories = new List<string>();
        if (element.Value.TryGetProperty("categories", out var cats)
            && cats.ValueKind == JsonValueKind.Array)
            categories.AddRange(cats.EnumerateArray().Select(JsonToString));
        var series = new List<PptChartSeries>();
        if (element.Value.TryGetProperty("series", out var seriesElement)
            && seriesElement.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in seriesElement.EnumerateArray())
            {
                var values = new List<double>();
                if (item.TryGetProperty("values", out var vals)
                    && vals.ValueKind == JsonValueKind.Array)
                {
                    foreach (var value in vals.EnumerateArray())
                    {
                        if (value.ValueKind == JsonValueKind.Number && value.TryGetDouble(out var number))
                            values.Add(number);
                        else if (value.ValueKind == JsonValueKind.String
                                 && double.TryParse(value.GetString(), NumberStyles.Float,
                                     CultureInfo.InvariantCulture, out var parsed))
                            values.Add(parsed);
                        else
                            values.Add(0);
                    }
                }
                series.Add(new PptChartSeries
                {
                    Name = Str(item, "name") ?? "",
                    Values = values,
                });
            }
        }
        if (categories.Count == 0 || series.Count == 0)
            throw new NativeUsageException("ChartDef 必须提供 categories 与 series");
        var chartType = Str(element.Value, "chart_type") ?? "column";
        if (chartType is not ("column" or "bar" or "line" or "pie" or "doughnut"
            or "area" or "radar"))
            throw new NativeUsageException(
                $"未知 chart_type: {chartType}（支持 column/bar/line/pie/doughnut/area/radar）");
        return new PptChartDef
        {
            ChartType = chartType,
            Categories = categories,
            Series = series,
            HasLegend = Bool(element.Value, "has_legend") ?? true,
            Title = Str(element.Value, "title"),
            X = Dbl(element.Value, "x"),
            Y = Dbl(element.Value, "y"),
            Width = Dbl(element.Value, "width"),
            Height = Dbl(element.Value, "height"),
        };
    }

    public static PptFrameworkDiagram ParseFrameworkDiagram(JsonElement p)
    {
        if (p.ValueKind != JsonValueKind.Object || !p.TryGetProperty("diagram", out var element)
            || element.ValueKind != JsonValueKind.Object)
            throw new NativeUsageException("缺少参数 diagram（JSON 对象）");
        if (!element.TryGetProperty("stages", out var stagesElement)
            || stagesElement.ValueKind != JsonValueKind.Array)
            throw new NativeUsageException("缺少参数 diagram.stages（JSON 数组）");
        var stages = new List<PptFrameworkStage>();
        {
            foreach (var stage in stagesElement.EnumerateArray())
            {
                var nodes = new List<PptFrameworkNode>();
                if (stage.TryGetProperty("nodes", out var nodesElement)
                    && nodesElement.ValueKind == JsonValueKind.Array)
                {
                    foreach (var node in nodesElement.EnumerateArray())
                    {
                        var id = Str(node, "id");
                        if (string.IsNullOrWhiteSpace(id))
                            throw new NativeUsageException("FrameworkNode 必须提供 id");
                        var text = Str(node, "text");
                        if (text is null)
                            throw new NativeUsageException($"FrameworkNode {id} 必须提供 text");
                        var style = Str(node, "style") ?? "rounded_rect";
                        if (style is not ("rect" or "rounded_rect" or "ellipse" or "diamond"
                            or "process" or "decision" or "data" or "document" or "terminator"
                            or "predefined" or "cylinder" or "parallelogram" or "chevron"
                            or "right_arrow" or "left_arrow" or "callout" or "cloud"))
                            throw new NativeUsageException(
                                $"未知 FrameworkNode.style: {style}");
                        nodes.Add(new PptFrameworkNode
                        {
                            Id = id,
                            Text = text,
                            Style = style,
                            FillColor = ColorStr(node, "fill_color"),
                            FontSize = Int(node, "font_size") ?? 12,
                            Bold = Bool(node, "bold") ?? true,
                        });
                    }
                }
                stages.Add(new PptFrameworkStage
                {
                    Label = Str(stage, "label") ?? "",
                    Color = ColorStr(stage, "color"),
                    Nodes = nodes,
                });
            }
        }
        var edges = new List<PptFrameworkEdge>();
        if (element.TryGetProperty("edges", out var edgesElement)
            && edgesElement.ValueKind == JsonValueKind.Array)
        {
            foreach (var edge in edgesElement.EnumerateArray())
            {
                var from = Str(edge, "from_id") ?? "";
                var to = Str(edge, "to_id") ?? "";
                if (from.Length == 0 || to.Length == 0)
                    throw new NativeUsageException("FrameworkEdge 必须提供 from_id 与 to_id");
                var edgeStyle = Str(edge, "edge_style") ?? "straight";
                if (edgeStyle is not ("straight" or "elbow" or "curve"))
                    throw new NativeUsageException(
                        $"未知 FrameworkEdge.edge_style: {edgeStyle}"
                        + "（支持 straight/elbow/curve）");
                edges.Add(new PptFrameworkEdge
                {
                    FromId = from,
                    ToId = to,
                    Label = Str(edge, "label"),
                    EdgeStyle = edgeStyle,
                });
            }
        }
        return new PptFrameworkDiagram
        {
            Title = Str(element, "title") ?? "",
            Subtitle = Str(element, "subtitle"),
            Stages = stages,
            Edges = edges,
            ColorScheme = Str(element, "color_scheme") ?? "professional-blue",
            NodeWidth = Dbl(element, "node_width") ?? 1.4,
            NodeHeight = Dbl(element, "node_height") ?? 0.55,
            StageGap = Dbl(element, "stage_gap") ?? 0.15,
            ArrowColor = ColorStr(element, "arrow_color"),
            FontCn = Str(element, "font_cn") ?? "宋体",
            FontEn = Str(element, "font_en") ?? "Times New Roman",
            FontSize = Int(element, "font_size") ?? 16,
            TitleFontSize = Int(element, "title_font_size") ?? 24,
        };
    }

    public static List<PptHyperlinkDef> ParseHyperlinks(JsonElement p)
    {
        if (p.ValueKind != JsonValueKind.Object || !p.TryGetProperty("links", out var links)
            || links.ValueKind != JsonValueKind.Array)
            throw new NativeUsageException("缺少参数 links（JSON 数组）");
        var result = new List<PptHyperlinkDef>();
        foreach (var link in links.EnumerateArray())
        {
            var slide = Int(link, "slide") ?? 0;
            if (slide < 1)
                throw new NativeUsageException("PptHyperlinkDef.slide 必须 >= 1");
            var url = Str(link, "url");
            if (string.IsNullOrWhiteSpace(url))
                throw new NativeUsageException("PptHyperlinkDef 必须提供 url");
            var shapeIndex = Int(link, "shape_index");
            var shapeText = Str(link, "shape_text");
            if (shapeIndex is < 1)
                throw new NativeUsageException("PptHyperlinkDef.shape_index 必须 >= 1");
            if (shapeIndex is null && string.IsNullOrWhiteSpace(shapeText))
                throw new NativeUsageException(
                    "PptHyperlinkDef 必须提供 shape_index 或 shape_text");
            result.Add(new PptHyperlinkDef
            {
                Slide = slide,
                Url = url,
                ShapeIndex = shapeIndex,
                ShapeText = shapeText,
                TextMatch = Str(link, "text_match"),
                MatchCase = Bool(link, "match_case") ?? false,
            });
        }
        return result;
    }

    // ── JSON 基础读取 ────────────────────────────────────────────

    public static JsonElement? Obj(JsonElement parent, string name) =>
        parent.ValueKind == JsonValueKind.Object && parent.TryGetProperty(name, out var value)
            && value.ValueKind == JsonValueKind.Object
            ? value
            : null;

    public static string? Str(JsonElement parent, string name)
    {
        if (parent.ValueKind != JsonValueKind.Object || !parent.TryGetProperty(name, out var value))
            return null;
        return value.ValueKind switch
        {
            JsonValueKind.String => value.GetString(),
            JsonValueKind.Null or JsonValueKind.Undefined => null,
            _ => value.ToString(),
        };
    }

    public static double? Dbl(JsonElement parent, string name)
    {
        if (parent.ValueKind != JsonValueKind.Object || !parent.TryGetProperty(name, out var value))
            return null;
        if (value.ValueKind == JsonValueKind.Number && value.TryGetDouble(out var number))
            return number;
        if (value.ValueKind == JsonValueKind.String
            && double.TryParse(value.GetString(), NumberStyles.Float,
                CultureInfo.InvariantCulture, out var parsed))
            return parsed;
        return null;
    }

    public static int? Int(JsonElement parent, string name)
    {
        var value = Dbl(parent, name);
        if (value is null)
            return null;
        return (int)Math.Round(value.Value, MidpointRounding.AwayFromZero);
    }

    public static bool? Bool(JsonElement parent, string name)
    {
        if (parent.ValueKind != JsonValueKind.Object || !parent.TryGetProperty(name, out var value))
            return null;
        if (value.ValueKind is JsonValueKind.True or JsonValueKind.False)
            return value.GetBoolean();
        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var number))
            return number switch { 1 => true, 0 => false, _ => null };
        if (value.ValueKind != JsonValueKind.String)
            return null;
        return value.GetString()?.Trim().ToLowerInvariant() switch
        {
            "true" or "1" or "yes" or "on" => true,
            "false" or "0" or "no" or "off" or "" => false,
            _ => null,
        };
    }

    public static bool OptBool(JsonElement parent, string name, bool fallback) =>
        Bool(parent, name) ?? fallback;

    public static string? ColorStr(JsonElement parent, string name)
    {
        var value = Str(parent, name);
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }

    public static List<string> StringList(JsonElement parent, string name)
    {
        if (parent.ValueKind != JsonValueKind.Object || !parent.TryGetProperty(name, out var value)
            || value.ValueKind != JsonValueKind.Array)
            return [];
        return value.EnumerateArray().Select(JsonToString).ToList();
    }

    public static List<double>? DoubleList(JsonElement parent, string name)
    {
        if (parent.ValueKind != JsonValueKind.Object || !parent.TryGetProperty(name, out var value)
            || value.ValueKind != JsonValueKind.Array)
            return null;
        var result = new List<double>();
        foreach (var item in value.EnumerateArray())
        {
            if (item.ValueKind == JsonValueKind.Number && item.TryGetDouble(out var number))
                result.Add(number);
            else if (item.ValueKind == JsonValueKind.String
                     && double.TryParse(item.GetString(), NumberStyles.Float,
                         CultureInfo.InvariantCulture, out var parsed))
                result.Add(parsed);
        }
        return result;
    }

    public static string JsonToString(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.Null or JsonValueKind.Undefined => "",
        JsonValueKind.String => value.GetString() ?? "",
        JsonValueKind.True => "True",
        JsonValueKind.False => "False",
        _ => value.GetRawText(),
    };

    private static Dictionary<string, PptColorScheme> BuildSchemes()
    {
        PptColorScheme Make(string name, string display, string description,
            string title, string body, string accent, string background,
            string shapeFill, string header, int titleSize = 32, int bodySize = 18,
            string titleFont = "微软雅黑", string bodyFont = "微软雅黑") => new()
        {
            Name = name,
            DisplayName = display,
            Description = description,
            TitleColor = title,
            BodyColor = body,
            AccentColor = accent,
            BackgroundColor = background,
            ShapeFillColor = shapeFill,
            TableHeaderColor = header,
            TitleFont = titleFont,
            BodyFont = bodyFont,
            TitleSize = titleSize,
            BodySize = bodySize,
        };

        var schemes = new Dictionary<string, PptColorScheme>(StringComparer.Ordinal)
        {
            ["professional-blue"] = Make("professional-blue", "商务蓝",
                "专业商务风格，深蓝主色配合浅灰背景", "000000", "333333", "2563EB",
                "F8FAFC", "DBEAFE", "1A3C6D"),
            ["modern-dark"] = Make("modern-dark", "暗夜模式",
                "现代深色主题，白字深灰底，适合大屏演示", "FFFFFF", "D1D5DB", "60A5FA",
                "1F2937", "374151", "111827"),
            ["warm-earth"] = Make("warm-earth", "暖色大地",
                "温暖自然风格，橙色与棕色搭配", "7C2D12", "44403C", "D97706",
                "FFF7ED", "FED7AA", "7C2D12"),
            ["elegant-minimal"] = Make("elegant-minimal", "极简优雅",
                "黑白灰极简风格，紫色点缀，适合学术/设计", "18181B", "52525B", "8B5CF6",
                "FFFFFF", "F4F4F5", "18181B", 28, 16),
            ["vibrant"] = Make("vibrant", "活力缤纷",
                "高对比度配色，玫红+青绿，适合创意/营销", "BE185D", "27272A", "14B8A6",
                "FFF1F2", "FCE7F3", "BE185D", 34, 20),
            ["corporate-green"] = Make("corporate-green", "企业绿",
                "稳重大气，深绿主色配合浅绿背景", "064E3B", "374151", "10B981",
                "ECFDF5", "D1FAE5", "064E3B"),
            ["tech-purple"] = Make("tech-purple", "科技紫",
                "科技感紫色渐变，适合互联网/AI/创新主题", "4C1D95", "374151", "7C3AED",
                "F5F3FF", "EDE9FE", "4C1D95"),
            ["classic-red"] = Make("classic-red", "经典中国红",
                "庄重典雅的中国红，适合党政/党建/传统节日主题", "991B1B", "44403C", "DC2626",
                "FEF2F2", "FEE2E2", "991B1B", 34, 20),
            ["fresh-nature"] = Make("fresh-nature", "清新自然",
                "清爽绿色系，适合环保/健康/农业/教育主题", "166534", "374151", "22C55E",
                "F0FDF4", "DCFCE7", "166534", 30, 18),
            ["ocean-blue"] = Make("ocean-blue", "海洋蓝",
                "清爽海洋蓝，适合医疗/教育/水利/科技主题", "000000", "334155", "0EA5E9",
                "F0F9FF", "E0F2FE", "0C4A6E"),
            ["sunset-orange"] = Make("sunset-orange", "日落暖橙",
                "温暖活力的橙色系，适合餐饮/旅游/创意/生活主题", "9A3412", "44403C", "F97316",
                "FFF7ED", "FFEDD5", "9A3412", 32, 20),
            ["monochrome-gray"] = Make("monochrome-gray", "高级灰",
                "克制高级灰，适合设计/建筑/金融/高端品牌主题", "1E293B", "475569", "64748B",
                "F8FAFC", "F1F5F9", "1E293B", 28, 16),
        };
        return schemes;
    }
}
