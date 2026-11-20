using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeNative;

/// <summary>Excel 图表：直接生成 DrawingML Chart XML（ChartPart）并写入工作表锚点。</summary>
internal static class ExcelFileCharts
{
    private static readonly XNamespace C = ExcelFileDrawing.C;
    private static readonly XNamespace A = ExcelFileDrawing.A;
    private static readonly XNamespace Xdr = ExcelFileDrawing.Xdr;
    private static readonly XNamespace ROwn = ExcelFileDrawing.ROwn;

    private const int CategoryAxisId = 48650112;
    private const int ValueAxisId = 48672768;

    public static List<string> AddChartToSheet(
        WorksheetPart worksheetPart,
        ExcelChartDef chart,
        string defaultSheet,
        IReadOnlySet<string> sheetNames,
        bool insertWorksheetDrawingElement)
    {
        var warnings = new List<string>();
        var chartXml = BuildChartSpace(chart, defaultSheet, sheetNames, warnings);

        var drawingsPart = worksheetPart.DrawingsPart
                           ?? worksheetPart.AddNewPart<DrawingsPart>();
        var drawingDocument = ExcelFileDrawing.LoadDrawing(drawingsPart);
        var shapeId = ExcelFileDrawing.NextShapeId(drawingDocument);

        var chartPart = drawingsPart.AddNewPart<ChartPart>();
        using (var buffer = new MemoryStream())
        {
            chartXml.Save(buffer, SaveOptions.DisableFormatting);
            buffer.Position = 0;
            chartPart.FeedData(buffer);
        }
        var relId = drawingsPart.GetIdOfPart(chartPart);

        if (!ExcelA1.TryParseCell(chart.Anchor, out var anchorCol, out var anchorRow))
            throw new NativeUsageException($"无效的图表锚点单元格: {chart.Anchor}");
        var widthEmu = ExcelFileDrawing.CmToEmu(chart.Width, "chart.width");
        var heightEmu = ExcelFileDrawing.CmToEmu(chart.Height, "chart.height");
        var (toCol, toRow) = ComputeEndAnchor(
            anchorCol - 1, anchorRow - 1, chart.Width, chart.Height);

        var anchorElement = new XElement(Xdr + "twoCellAnchor",
            new XAttribute("editAs", "oneCell"),
            ExcelFileDrawing.Marker(Xdr + "from", anchorCol - 1, anchorRow - 1),
            ExcelFileDrawing.Marker(Xdr + "to", toCol, toRow),
            new XElement(Xdr + "graphicFrame",
                new XAttribute("macro", ""),
                new XElement(Xdr + "nvGraphicFramePr",
                    new XElement(Xdr + "cNvPr",
                        new XAttribute("id", shapeId.ToString(CultureInfo.InvariantCulture)),
                        new XAttribute("name", $"Chart {shapeId}")),
                    new XElement(Xdr + "cNvGraphicFramePr")),
                new XElement(Xdr + "xfrm",
                    new XElement(A + "off", new XAttribute("x", "0"), new XAttribute("y", "0")),
                    new XElement(A + "ext",
                        new XAttribute("cx", widthEmu.ToString(CultureInfo.InvariantCulture)),
                        new XAttribute("cy", heightEmu.ToString(CultureInfo.InvariantCulture)))),
                new XElement(A + "graphic",
                    new XElement(A + "graphicData",
                        new XAttribute("uri", "http://schemas.openxmlformats.org/drawingml/2006/chart"),
                        new XElement(C + "chart",
                            new XAttribute(XNamespace.Xmlns + "c", C.NamespaceName),
                            new XAttribute(ROwn + "id", relId))))),
            ExcelFileDrawing.ClientData());

        ExcelFileDrawing.AddDrawingAnchor(drawingDocument, anchorElement);
        ExcelFileDrawing.SaveDrawing(drawingsPart, drawingDocument);

        if (insertWorksheetDrawingElement && worksheetPart.Worksheet is not null)
            ExcelFileDrawing.EnsureDrawingElement(
                worksheetPart.Worksheet, worksheetPart.GetIdOfPart(drawingsPart));

        return warnings;
    }

    private static XDocument BuildChartSpace(
        ExcelChartDef def,
        string defaultSheet,
        IReadOnlySet<string> sheetNames,
        List<string> warnings)
    {
        var chartType = def.ChartType;
        var plot = new XElement(C + "plotArea", new XElement(C + "layout"));
        XElement chartElement;
        var axes = new List<XElement>();

        if (chartType == "scatter")
        {
            var series = BuildScatterSeries(def, defaultSheet, sheetNames);
            chartElement = new XElement(C + "scatterChart",
                Val("scatterStyle", "lineMarker"),
                Val("varyColors", 0),
                series,
                Val("axId", CategoryAxisId),
                Val("axId", ValueAxisId));
            axes.Add(ValueAxis(CategoryAxisId, ValueAxisId, "b",
                def.XAxisTitle is null ? null : def.XAxisTitle));
            axes.Add(ValueAxis(ValueAxisId, CategoryAxisId, "l",
                def.YAxisTitle is null ? null : def.YAxisTitle));
        }
        else
        {
            var data = ResolveRange(def.DataRange, defaultSheet, sheetNames, "图表数据范围");
            var categories = def.CategoriesRange is null
                ? (ExcelRange?)null
                : ResolveRange(def.CategoriesRange, data.Sheet!, sheetNames, "图表分类范围");
            var series = BuildSeries(def, data, categories);

            switch (chartType)
            {
                case "column":
                case "bar":
                    var horizontal = chartType == "bar";
                    chartElement = new XElement(C + "barChart",
                        Val("barDir", horizontal ? "bar" : "col"),
                        Val("grouping", "clustered"),
                        Val("varyColors", 0),
                        series,
                        Val("gapWidth", 150),
                        Val("axId", CategoryAxisId),
                        Val("axId", ValueAxisId));
                    axes.Add(CategoryAxis(CategoryAxisId, ValueAxisId,
                        horizontal ? "l" : "b", def.XAxisTitle));
                    axes.Add(ValueAxis(ValueAxisId, CategoryAxisId,
                        horizontal ? "b" : "l", def.YAxisTitle));
                    break;
                case "line":
                    chartElement = new XElement(C + "lineChart",
                        Val("grouping", "standard"),
                        Val("varyColors", 0),
                        series,
                        Val("axId", CategoryAxisId),
                        Val("axId", ValueAxisId));
                    axes.Add(CategoryAxis(CategoryAxisId, ValueAxisId, "b", def.XAxisTitle));
                    axes.Add(ValueAxis(ValueAxisId, CategoryAxisId, "l", def.YAxisTitle));
                    break;
                case "area":
                    chartElement = new XElement(C + "areaChart",
                        Val("grouping", "standard"),
                        Val("varyColors", 0),
                        series,
                        Val("axId", CategoryAxisId),
                        Val("axId", ValueAxisId));
                    axes.Add(CategoryAxis(CategoryAxisId, ValueAxisId, "b", def.XAxisTitle));
                    axes.Add(ValueAxis(ValueAxisId, CategoryAxisId, "l", def.YAxisTitle));
                    break;
                case "pie":
                    chartElement = new XElement(C + "pieChart",
                        Val("varyColors", 1),
                        series);
                    AddAxisTitleWarnings(chartType, def, warnings);
                    break;
                case "doughnut":
                    chartElement = new XElement(C + "doughnutChart",
                        Val("varyColors", 1),
                        series,
                        Val("holeSize", 50));
                    AddAxisTitleWarnings(chartType, def, warnings);
                    break;
                case "radar":
                    chartElement = new XElement(C + "radarChart",
                        Val("radarStyle", "marker"),
                        Val("varyColors", 0),
                        series,
                        Val("axId", CategoryAxisId),
                        Val("axId", ValueAxisId));
                    axes.Add(CategoryAxis(CategoryAxisId, ValueAxisId, "b", def.XAxisTitle));
                    axes.Add(ValueAxis(ValueAxisId, CategoryAxisId, "l", def.YAxisTitle));
                    break;
                default:
                    throw new NativeUsageException($"不支持的图表类型: {chartType}");
            }
        }

        plot.Add(chartElement);
        foreach (var axis in axes)
            plot.Add(axis);

        var chart = new XElement(C + "chart");
        if (!string.IsNullOrWhiteSpace(def.Title))
            chart.Add(RichTitle(def.Title!));
        else
            chart.Add(Val("autoTitleDeleted", 1));
        chart.Add(plot);
        if (def.Legend != "none")
        {
            var position = def.Legend switch
            {
                "right" => "r",
                "bottom" => "b",
                "top" => "t",
                "left" => "l",
                _ => "r",
            };
            chart.Add(new XElement(C + "legend",
                Val("legendPos", position),
                Val("overlay", 0)));
        }
        chart.Add(Val("plotVisOnly", 1));
        chart.Add(Val("dispBlanksAs", "gap"));

        var chartSpace = new XElement(C + "chartSpace",
            new XAttribute(XNamespace.Xmlns + "c", C.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "a", A.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "r", ROwn.NamespaceName),
            Val("date1904", 0),
            Val("roundedCorners", 0));
        if (def.Style is not null)
            chartSpace.Add(Val("style", def.Style.Value));
        chartSpace.Add(chart);
        return new XDocument(chartSpace);
    }

    private static void AddAxisTitleWarnings(
        string chartType, ExcelChartDef def, List<string> warnings)
    {
        if (!string.IsNullOrWhiteSpace(def.XAxisTitle))
            warnings.Add($"{chartType} 图表没有 x 轴，已忽略 x_axis_title");
        if (!string.IsNullOrWhiteSpace(def.YAxisTitle))
            warnings.Add($"{chartType} 图表没有 y 轴，已忽略 y_axis_title");
    }

    private static List<XElement> BuildSeries(
        ExcelChartDef def, ExcelRange data, ExcelRange? categories)
    {
        var startRow = data.MinRow + (def.TitlesFromData ? 1 : 0);
        if (startRow > data.MaxRow)
            throw new NativeRuntimeException($"图表数据区域没有数据行: {def.DataRange}");

        var series = new List<XElement>();
        var index = 0;
        for (var col = data.MinCol; col <= data.MaxCol; col++)
        {
            var valueFormula = ExcelA1.AbsoluteRangeFormula(
                data.Sheet!, col, startRow, col, data.MaxRow);
            var titleFormula = def.TitlesFromData
                ? ExcelA1.AbsoluteCellFormula(data.Sheet!, col, data.MinRow)
                : null;
            var categoryFormula = categories is null
                ? null
                : ExcelA1.AbsoluteRangeFormula(
                    categories.Value.Sheet!, categories.Value.MinCol, categories.Value.MinRow,
                    categories.Value.MaxCol, categories.Value.MaxRow);
            series.Add(BasicSeries(index, titleFormula, categoryFormula, valueFormula));
            index++;
        }
        return series;
    }

    private static List<XElement> BuildScatterSeries(
        ExcelChartDef def, string defaultSheet, IReadOnlySet<string> sheetNames)
    {
        var data = ResolveRange(def.DataRange, defaultSheet, sheetNames, "图表数据范围");
        var series = new List<XElement>();
        if (def.CategoriesRange is not null)
        {
            var xRange = ResolveRange(
                def.CategoriesRange, data.Sheet!, sheetNames, "图表分类范围");
            var startRow = data.MinRow + (def.TitlesFromData ? 1 : 0);
            if (startRow > data.MaxRow)
                throw new NativeRuntimeException($"图表数据区域没有数据行: {def.DataRange}");
            var xFormula = ExcelA1.AbsoluteRangeFormula(
                xRange.Sheet!, xRange.MinCol, xRange.MinRow, xRange.MaxCol, xRange.MaxRow);
            var index = 0;
            for (var col = data.MinCol; col <= data.MaxCol; col++)
            {
                var yFormula = ExcelA1.AbsoluteRangeFormula(
                    data.Sheet!, col, startRow, col, data.MaxRow);
                var titleFormula = def.TitlesFromData
                    ? ExcelA1.AbsoluteCellFormula(data.Sheet!, col, data.MinRow)
                    : null;
                series.Add(ScatterSeries(index, titleFormula, xFormula, yFormula));
                index++;
            }
            return series;
        }

        if (data.MaxCol - data.MinCol != 1)
            throw new NativeRuntimeException(
                "scatter 的 data_range 需要恰好两列（X,Y），或显式提供 categories_range");
        var row = data.MinRow + (def.TitlesFromData ? 1 : 0);
        if (row > data.MaxRow)
            throw new NativeRuntimeException($"图表数据区域没有数据行: {def.DataRange}");
        var x = ExcelA1.AbsoluteRangeFormula(
            data.Sheet!, data.MinCol, row, data.MinCol, data.MaxRow);
        var y = ExcelA1.AbsoluteRangeFormula(
            data.Sheet!, data.MaxCol, row, data.MaxCol, data.MaxRow);
        var title = def.TitlesFromData
            ? ExcelA1.AbsoluteCellFormula(data.Sheet!, data.MaxCol, data.MinRow)
            : null;
        series.Add(ScatterSeries(0, title, x, y));
        return series;
    }

    private static XElement BasicSeries(
        int index, string? titleFormula, string? categoryFormula, string valueFormula)
    {
        var elements = new List<XElement>
        {
            Val("idx", index),
            Val("order", index),
        };
        if (titleFormula is not null)
            elements.Add(new XElement(C + "tx", StringReference(titleFormula)));
        if (categoryFormula is not null)
            elements.Add(new XElement(C + "cat", StringReference(categoryFormula)));
        elements.Add(new XElement(C + "val", NumberReference(valueFormula)));
        return new XElement(C + "ser", elements);
    }

    private static XElement ScatterSeries(
        int index, string? titleFormula, string xFormula, string yFormula)
    {
        var elements = new List<XElement>
        {
            Val("idx", index),
            Val("order", index),
        };
        if (titleFormula is not null)
            elements.Add(new XElement(C + "tx", StringReference(titleFormula)));
        elements.Add(new XElement(C + "xVal", NumberReference(xFormula)));
        elements.Add(new XElement(C + "yVal", NumberReference(yFormula)));
        return new XElement(C + "ser", elements);
    }

    private static XElement StringReference(string formula) => new(C + "strRef",
        new XElement(C + "f", formula));

    private static XElement NumberReference(string formula) => new(C + "numRef",
        new XElement(C + "f", formula));

    private static XElement CategoryAxis(
        int axisId, int crossAxisId, string position, string? title)
    {
        var axis = new XElement(C + "catAx",
            Val("axId", axisId),
            new XElement(C + "scaling", Val("orientation", "minMax")),
            Val("delete", 0),
            Val("axPos", position));
        if (!string.IsNullOrWhiteSpace(title))
            axis.Add(RichTitle(title!));
        axis.Add(NumberFormat());
        axis.Add(Val("majorTickMark", "out"));
        axis.Add(Val("minorTickMark", "none"));
        axis.Add(Val("tickLblPos", "nextTo"));
        axis.Add(Val("crossAx", crossAxisId));
        axis.Add(Val("crosses", "autoZero"));
        axis.Add(Val("auto", 1));
        axis.Add(Val("lblAlgn", "ctr"));
        axis.Add(Val("lblOffset", 100));
        axis.Add(Val("noMultiLvlLbl", 0));
        return axis;
    }

    private static XElement ValueAxis(
        int axisId, int crossAxisId, string position, string? title)
    {
        var axis = new XElement(C + "valAx",
            Val("axId", axisId),
            new XElement(C + "scaling", Val("orientation", "minMax")),
            Val("delete", 0),
            Val("axPos", position));
        if (!string.IsNullOrWhiteSpace(title))
            axis.Add(RichTitle(title!));
        axis.Add(NumberFormat());
        axis.Add(Val("majorTickMark", "out"));
        axis.Add(Val("minorTickMark", "none"));
        axis.Add(Val("tickLblPos", "nextTo"));
        axis.Add(Val("crossAx", crossAxisId));
        axis.Add(Val("crosses", "autoZero"));
        axis.Add(Val("crossBetween", "between"));
        return axis;
    }

    private static XElement NumberFormat() => new(C + "numFmt",
        new XAttribute("formatCode", "General"),
        new XAttribute("sourceLinked", "1"));

    private static XElement RichTitle(string text) => new(C + "title",
        new XElement(C + "tx",
            new XElement(C + "rich",
                new XElement(A + "bodyPr"),
                new XElement(A + "lstStyle"),
                new XElement(A + "p",
                    new XElement(A + "pPr", new XElement(A + "defRPr")),
                    new XElement(A + "r",
                        new XElement(A + "rPr",
                            new XAttribute("lang", "zh-CN"),
                            new XAttribute("dirty", "0")),
                        new XElement(A + "t",
                            new XAttribute(XNamespace.Xml + "space", "preserve"),
                            text)),
                    new XElement(A + "endParaRPr",
                        new XAttribute("lang", "zh-CN"),
                        new XAttribute("dirty", "0"))))),
        Val("overlay", 0));

    private static XElement Val(string name, object value) => new(C + name,
        new XAttribute("val", Convert.ToString(value, CultureInfo.InvariantCulture)
                            ?? string.Empty));

    private static ExcelRange ResolveRange(
        string text, string defaultSheet, IReadOnlySet<string> sheetNames, string label)
    {
        if (!ExcelA1.TryParseRange(text, out var range))
            throw new NativeUsageException(
                $"{label} 必须是含具体行列的 A1 范围（如 A1:C10），不支持整列/整行或空值: {text}");
        var sheet = range.Sheet ?? defaultSheet;
        if (!sheetNames.Contains(sheet))
            throw new NativeRuntimeException($"图表引用的工作表不存在: {sheet}");
        return range with { Sheet = sheet };
    }

    private static (int Col, int Row) ComputeEndAnchor(
        int startCol, int startRow, double widthCm, double heightCm)
    {
        const double pxPerColumn = 64;
        const double pxPerRow = 20;
        const int maxCol = 16383;   // 0-based XFD
        const int maxRow = 1048575; // 0-based 1048576
        var widthPx = widthCm / 2.54 * 96;
        var heightPx = heightCm / 2.54 * 96;
        var columns = double.IsFinite(widthPx) && widthPx > 0
            ? (int)Math.Ceiling(widthPx / pxPerColumn)
            : 1;
        var rows = double.IsFinite(heightPx) && heightPx > 0
            ? (int)Math.Ceiling(heightPx / pxPerRow)
            : 1;
        // 裁剪到网格剩余空间，保证 to 锚点不超过 XFD1048576。
        columns = Math.Clamp(columns, 1, Math.Max(1, maxCol - startCol));
        rows = Math.Clamp(rows, 1, Math.Max(1, maxRow - startRow));
        return (Math.Min(startCol + columns, maxCol), Math.Min(startRow + rows, maxRow));
    }
}
