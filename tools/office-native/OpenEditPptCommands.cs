using System.Text.Json;

namespace OfficeNative;

/// <summary>
/// 活动演示文稿 COM 编辑：
/// edit-open-powerpoint 追加幻灯片；ppt-add-flowchart-slide 追加流程图幻灯片。
/// 对应 Python office_com.try_edit_open_powerpoint / try_add_flowchart_slide。
/// </summary>
internal static class OpenEditPptCommands
{
    // PpSlideLayout（1 基），与 office_com.PPT_COM_LAYOUT_MAP 一致。
    private static readonly Dictionary<string, int> LayoutMap = new(StringComparer.Ordinal)
    {
        ["title"] = 1,
        ["text"] = 2,
        ["content"] = 2,
        ["section"] = 3,
        ["two_content"] = 4,
        ["comparison"] = 5,
        ["title_only"] = 6,
        ["blank"] = 12,
        ["content_caption"] = 7,
        ["picture_caption"] = 8,
    };

    // MsoAutoShapeType（与 office_com._PPT_SHAPE_TYPE_MAP 对齐）。
    private static readonly Dictionary<string, int> ShapeTypeMap = new(StringComparer.Ordinal)
    {
        ["rect"] = 1,
        ["rounded_rect"] = 5,
        ["oval"] = 9,
        ["ellipse"] = 9,
        ["diamond"] = 4,
        ["arrow"] = 33,
        ["right_arrow"] = 33,
        ["left_arrow"] = 34,
        ["chevron"] = 52,
        ["parallelogram"] = 2,
        ["process"] = 61,
        ["decision"] = 63,
        ["data"] = 62,
        ["document"] = 60,
        ["terminator"] = 78,
        ["predefined"] = 77,
        ["cylinder"] = 98,
        ["callout"] = 105,
        ["cloud"] = 179,
    };

    // MsoArrowheadStyle：triangle=2, open=3, stealth=4。
    private static readonly Dictionary<string, int> ArrowheadMap = new(StringComparer.Ordinal)
    {
        ["triangle"] = 2,
        ["open"] = 3,
        ["stealth"] = 4,
    };

    // MsoConnectorType：straight=1, elbow=2, curve=3。
    private static readonly Dictionary<string, int> ConnectorTypeMap = new(StringComparer.Ordinal)
    {
        ["straight"] = 1,
        ["elbow"] = 2,
        ["curve"] = 3,
    };

    private static readonly HashSet<string> SupportedItemTypes = new(StringComparer.Ordinal)
    {
        "paragraph", "table", "image", "shape", "connector",
    };

    // ── edit-open-powerpoint ────────────────────────────────────

    public static Dictionary<string, object?> RunEditOpenPowerpoint(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        var slides = PptModels.ParseSlides(p, "append_slides");
        NativeCommon.EnsureFileLight(filePath);
        return TryEditOpenPowerpoint(Path.GetFullPath(filePath), slides)
               ?? throw new NativeRuntimeException(
                   $"目标文件未在活动 PowerPoint/WPS 实例中打开: {filePath}；请先打开文件");
    }

    /// <summary>
    /// 若目标文件已在 PowerPoint/WPS 中打开，则通过 COM 追加幻灯片；
    /// 未找到活动文档时返回 null（供文件级 edit_powerpoint 回退）。
    /// </summary>
    internal static Dictionary<string, object?>? TryEditOpenPowerpoint(
        string filePath, List<PptSlideDef> slides)
    {
        var app = OpenEditCommon.FindPowerPoint(filePath, out var presentation, out var progId);
        if (app is null || presentation is null)
        {
            NativeCommon.ReleaseCom(presentation);
            NativeCommon.ReleaseCom(app);
            return null;
        }
        var previousAlerts = GetDisplayAlerts(app);
        SetDisplayAlerts(app, 1); // ppAlertsNone
        try
        {
            OpenEditCommon.EnsureWritablePresentation(presentation);
            ValidateSlides(slides);
            var initialSlideCount = (int)presentation.Slides.Count;
            try
            {
                foreach (var slideDef in slides)
                {
                    var layout = LayoutMap.GetValueOrDefault(slideDef.Layout, 2);
                    var position = (int)presentation.Slides.Count + 1;
                    dynamic slide = presentation.Slides.Add(position, layout);
                    try
                    {
                        FillSlide(slide, slideDef);
                    }
                    finally
                    {
                        NativeCommon.ReleaseCom(slide);
                    }
                }
                // Save 失败也必须回滚本次新增的幻灯片。
                OpenEditCommon.SavePresentation(presentation);
            }
            catch (Exception error)
            {
                try
                {
                    RollbackSlides(presentation, initialSlideCount);
                }
                catch (Exception rollbackError)
                {
                    throw new NativeRuntimeException(
                        $"活动 PowerPoint 演示文稿编辑失败: {error.Message}；"
                        + $"回滚失败，请勿保存该演示文稿: {rollbackError.Message}");
                }
                throw new NativeRuntimeException(
                    $"活动 PowerPoint 演示文稿编辑失败，已回滚本次新增幻灯片: {error.Message}");
            }

            return OpenEditCommon.EditResult(filePath, slides.Count,
                $"已通过 {progId} 修改并保存活动 PowerPoint 演示文稿");
        }
        finally
        {
            RestoreDisplayAlerts(app, previousAlerts);
            NativeCommon.ReleaseCom(presentation);
            NativeCommon.ReleaseCom(app);
        }
    }

    private static int? GetDisplayAlerts(dynamic app)
    {
        try { return (int)app.DisplayAlerts; }
        catch { return null; }
    }

    private static void SetDisplayAlerts(dynamic app, int value)
    {
        try { app.DisplayAlerts = value; }
        catch { /* WPS/旧版本可能不支持 */ }
    }

    private static void RestoreDisplayAlerts(dynamic app, int? previous)
    {
        if (previous is int value)
            SetDisplayAlerts(app, value);
    }

    private static void ValidateSlides(List<PptSlideDef> slides)
    {
        for (var index = 0; index < slides.Count; index++)
        {
            var slideNumber = index + 1;
            foreach (var items in new[] { slides[index].Body, slides[index].Left, slides[index].Right })
            {
                if (items is null)
                    continue;
                foreach (var item in items)
                {
                    if (!SupportedItemTypes.Contains(item.Type))
                        throw new NativeRuntimeException(
                            $"幻灯片 #{slideNumber} 的 '{item.Type}' 类型暂不支持活动 PPT 编辑");
                    if (HasComUnsupportedFields(item))
                        throw new NativeRuntimeException(
                            $"幻灯片 #{slideNumber} 的 '{item.Type}' 使用了活动 PPT 暂不支持的"
                            + "几何/裁剪/超链接字段；请关闭文件后使用文件级 edit_powerpoint");
                }
            }
        }
    }

    private static bool HasComUnsupportedFields(PptSlideItem item)
    {
        if (item.Type == "table" && item.Table is { } table)
            return table.X is not null || table.Y is not null
                || table.Width is not null || table.Height is not null
                || table.ColumnWidths is not null;
        if (item.Type == "image" && item.Image is { } image)
            return image.CropLeft is not null || image.CropRight is not null
                || image.CropTop is not null || image.CropBottom is not null
                || image.HyperlinkUrl is not null;
        if (item.Type == "shape" && item.Shape is { } shape)
            return shape.HyperlinkUrl is not null;
        return false;
    }

    private static void FillSlide(dynamic slide, PptSlideDef slideDef)
    {
        if (!string.IsNullOrWhiteSpace(slideDef.BgColor))
        {
            slide.FollowMasterBackground = false;
            slide.Background.Fill.ForeColor.RGB =
                OpenEditCommon.HexToBgr(slideDef.BgColor, "背景色");
        }

        if (!string.IsNullOrEmpty(slideDef.Title))
        {
            dynamic shapes = slide.Shapes;
            if ((int)shapes.HasTitle != 0)
            {
                dynamic title = shapes.Title;
                try { title.TextFrame.TextRange.Text = ComMultiline(slideDef.Title, "\v"); }
                finally { NativeCommon.ReleaseCom(title); }
            }
        }

        if (!string.IsNullOrEmpty(slideDef.Subtitle))
        {
            dynamic shapes = slide.Shapes;
            try
            {
                var count = (int)shapes.Count;
                for (var index = 1; index <= count; index++)
                {
                    dynamic shape = shapes.Item(index);
                    try
                    {
                        if ((int)shape.HasTextFrame != 0
                            && (int)shape.PlaceholderFormat.Type == 4) // ppPlaceholderSubtitle
                        {
                            shape.TextFrame.TextRange.Text =
                                ComMultiline(slideDef.Subtitle, "\v");
                            break;
                        }
                    }
                    catch { /* 非占位符形状没有 PlaceholderFormat */ }
                    finally { NativeCommon.ReleaseCom(shape); }
                }
            }
            finally { NativeCommon.ReleaseCom(shapes); }
        }

        if (slideDef.Body is { Count: > 0 })
            AddSlideContent(slide, slideDef.Body, 0.5, 1.5, 8.0, 4.5);
        if (slideDef.Left is { Count: > 0 })
            AddSlideContent(slide, slideDef.Left, 0.5, 1.5, 4.0, 4.5);
        if (slideDef.Right is { Count: > 0 })
            AddSlideContent(slide, slideDef.Right, 5.0, 1.5, 4.0, 4.5);

        if (!string.IsNullOrEmpty(slideDef.Notes))
            SetSlideNotes(slide, slideDef.Notes!);
    }

    private static void SetSlideNotes(dynamic slide, string notes)
    {
        dynamic shapes = slide.NotesPage.Shapes;
        try
        {
            var count = (int)shapes.Count;
            for (var index = 1; index <= count; index++)
            {
                dynamic shape = shapes.Item(index);
                try
                {
                    if ((int)shape.HasTextFrame != 0
                        && (int)shape.PlaceholderFormat.Type == 2) // ppPlaceholderBody
                    {
                        shape.TextFrame.TextRange.Text = ComMultiline(notes, "\r");
                        break;
                    }
                }
                catch { /* 非占位符形状 */ }
                finally { NativeCommon.ReleaseCom(shape); }
            }
        }
        finally { NativeCommon.ReleaseCom(shapes); }
    }

    private static void AddSlideContent(
        dynamic slide, List<PptSlideItem> items,
        double defaultLeft, double defaultTop, double defaultWidth, double defaultHeight)
    {
        foreach (var item in items)
        {
            switch (item.Type)
            {
                case "paragraph" when !string.IsNullOrEmpty(item.Text):
                    AddParagraph(slide, item, defaultLeft, defaultTop, defaultWidth, defaultHeight);
                    break;
                case "table" when item.Table is not null:
                    AddTable(slide, item.Table, defaultLeft, defaultTop, defaultWidth, defaultHeight);
                    break;
                case "image" when item.Image is not null:
                    AddImage(slide, item.Image, defaultLeft, defaultTop);
                    break;
                case "shape" when item.Shape is not null:
                    AddShape(slide, item.Shape, defaultLeft, defaultTop);
                    break;
                case "connector" when item.Connector is not null:
                    AddConnector(slide, item.Connector);
                    break;
            }
        }
    }

    private static void AddParagraph(
        dynamic slide, PptSlideItem item,
        double defaultLeft, double defaultTop, double defaultWidth, double defaultHeight)
    {
        dynamic shape = slide.Shapes.AddTextbox(1,
            Pt(defaultLeft), Pt(defaultTop), Pt(defaultWidth), Pt(defaultHeight));
        try
        {
            dynamic textRange = shape.TextFrame.TextRange;
            textRange.Text = ComMultiline(item.Text, "\v");
            ApplyTextStyle(textRange, item.Style);
            if (!string.IsNullOrEmpty(item.Style?.Alignment))
            {
                textRange.ParagraphFormat.Alignment = item.Style.Alignment switch
                {
                    "left" => 1,
                    "center" => 2,
                    "right" => 3,
                    _ => 1,
                };
            }
        }
        finally
        {
            NativeCommon.ReleaseCom(shape);
        }
    }

    private static void AddTable(
        dynamic slide, PptSlideTable table,
        double defaultLeft, double defaultTop, double defaultWidth, double defaultHeight)
    {
        var rowCount = table.Rows.Count + (table.Headers.Count > 0 ? 1 : 0);
        var columnCount = Math.Max(
            table.Headers.Count,
            table.Rows.Count == 0 ? 0 : table.Rows.Max(row => row.Count));
        if (rowCount == 0 || columnCount == 0)
            return;

        dynamic shape = slide.Shapes.AddTable(rowCount, columnCount,
            Pt(defaultLeft), Pt(defaultTop), Pt(defaultWidth), Pt(defaultHeight));
        try
        {
            dynamic comTable = shape.Table;
            try
            {
                if (table.Headers.Count > 0)
                {
                    for (var columnIndex = 1; columnIndex <= table.Headers.Count; columnIndex++)
                    {
                        dynamic cell = comTable.Cell(1, columnIndex);
                        try
                        {
                            cell.Shape.TextFrame.TextRange.Text = ComMultiline(
                                table.Headers[columnIndex - 1], "\v");
                        }
                        finally { NativeCommon.ReleaseCom(cell); }
                    }
                }
                var startDataRow = table.Headers.Count > 0 ? 2 : 1;
                for (var rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
                {
                    var row = table.Rows[rowIndex];
                    for (var columnIndex = 1;
                         columnIndex <= row.Count && columnIndex <= columnCount;
                         columnIndex++)
                    {
                        dynamic cell = comTable.Cell(startDataRow + rowIndex, columnIndex);
                        try
                        {
                            cell.Shape.TextFrame.TextRange.Text = ComMultiline(
                                row[columnIndex - 1], "\v");
                        }
                        finally { NativeCommon.ReleaseCom(cell); }
                    }
                }
            }
            finally
            {
                NativeCommon.ReleaseCom(comTable);
            }
        }
        finally
        {
            NativeCommon.ReleaseCom(shape);
        }
    }

    private static void AddImage(
        dynamic slide, PptSlideImage image, double defaultLeft, double defaultTop)
    {
        byte[]? bytes = null;
        string? path = image.Path;
        string? downloaded = null;
        try
        {
            if (string.IsNullOrWhiteSpace(path) && !string.IsNullOrWhiteSpace(image.Url))
            {
                bytes = PptFileContent.DownloadImage(image.Url!, out var extension);
                if (bytes is null)
                    throw new NativeRuntimeException($"图片下载失败: {image.Url}");
                downloaded = Path.Combine(Path.GetTempPath(),
                    "office-native-image-" + Guid.NewGuid().ToString("N") + extension);
                File.WriteAllBytes(downloaded, bytes);
                path = downloaded;
            }
            else if (!string.IsNullOrWhiteSpace(path))
            {
                if (!File.Exists(path))
                    throw new NativeRuntimeException($"图片文件不存在: {path}");
                bytes = File.ReadAllBytes(path);
            }
            if (bytes is null || bytes.Length == 0)
                throw new NativeRuntimeException("图片内容为空，无法插入");
            if (!PptFileContent.TryDetectImageType(bytes, out _))
            {
                if (PptFileContent.LooksLikeSvg(bytes))
                    throw new NativeRuntimeException(
                        "SVG 图片不受支持，请转换为 PNG/JPEG 后重试");
                throw new NativeRuntimeException(
                    "无法识别的图片格式（仅支持 PNG/JPEG/GIF/BMP/TIFF/EMF/WMF）");
            }
            dynamic shape = slide.Shapes.AddPicture(
                path, false, true,
                Pt(image.X ?? defaultLeft), Pt(image.Y ?? defaultTop),
                Pt(image.Width ?? 4.0), Pt(image.Height ?? 3.0));
            NativeCommon.ReleaseCom(shape);
        }
        finally
        {
            if (downloaded is not null)
            {
                try { File.Delete(downloaded); }
                catch { /* 临时图片删除失败不影响插入结果 */ }
            }
        }
    }

    private static void AddShape(
        dynamic slide, PptSlideShape shapeDef, double defaultLeft, double defaultTop)
    {
        var shapeType = ShapeTypeMap.GetValueOrDefault(shapeDef.ShapeType, 1);
        dynamic shape = slide.Shapes.AddShape(shapeType,
            Pt(shapeDef.X ?? defaultLeft), Pt(shapeDef.Y ?? defaultTop),
            Pt(shapeDef.Width), Pt(shapeDef.Height));
        try
        {
            if (shapeDef.NoFill)
                shape.Fill.Visible = 0; // msoFalse
            else if (!string.IsNullOrWhiteSpace(shapeDef.BgColor))
                shape.Fill.ForeColor.RGB =
                    OpenEditCommon.HexToBgr(shapeDef.BgColor, "形状填充色");

            if (!string.IsNullOrWhiteSpace(shapeDef.LineColor))
                shape.Line.ForeColor.RGB =
                    OpenEditCommon.HexToBgr(shapeDef.LineColor, "形状边框色");
            if (shapeDef.LineWidth is not null)
                shape.Line.Weight = shapeDef.LineWidth.Value;
            if (shapeDef.LineDash)
                shape.Line.DashStyle = 2; // msoLineDash

            if (!string.IsNullOrEmpty(shapeDef.Text))
            {
                dynamic textRange = shape.TextFrame.TextRange;
                textRange.Text = ComMultiline(shapeDef.Text, "\v");
                textRange.ParagraphFormat.Alignment = 2; // ppAlignCenter
                if (shapeDef.VerticalText)
                {
                    try { shape.TextFrame.Orientation = 5; } // msoTextOrientationVerticalFarEast
                    catch { /* 部分 WPS 版本不支持竖排 */ }
                }
                if (!string.IsNullOrWhiteSpace(shapeDef.FontColor))
                    textRange.Font.Color.RGB =
                        OpenEditCommon.HexToBgr(shapeDef.FontColor, "文字颜色");
                if (shapeDef.FontSize is not null)
                    textRange.Font.Size = shapeDef.FontSize.Value;
                if (shapeDef.Bold == true)
                    textRange.Font.Bold = true;
            }
        }
        finally
        {
            NativeCommon.ReleaseCom(shape);
        }
    }

    private static void AddConnector(dynamic slide, PptSlideConnector connectorDef)
    {
        dynamic connector = slide.Shapes.AddConnector(1,
            Pt(connectorDef.X1), Pt(connectorDef.Y1),
            Pt(connectorDef.X2), Pt(connectorDef.Y2));
        try
        {
            if (!string.IsNullOrWhiteSpace(connectorDef.Color))
                connector.Line.ForeColor.RGB =
                    OpenEditCommon.HexToBgr(connectorDef.Color, "连接线颜色");
            connector.Line.Weight = connectorDef.Width;
            if (connectorDef.Dash)
                connector.Line.DashStyle = 2; // msoLineDash
            if (connectorDef.Arrowhead)
                connector.Line.EndArrowheadStyle =
                    ArrowheadMap.GetValueOrDefault(connectorDef.ArrowheadStyle, 2);
            if (connectorDef.BeginArrowhead)
                connector.Line.BeginArrowheadStyle = 2;
        }
        finally
        {
            NativeCommon.ReleaseCom(connector);
        }
    }

    private static void ApplyTextStyle(dynamic textRange, PptTextStyle? style)
    {
        if (style is null)
            return;
        dynamic font = textRange.Font;
        if (!string.IsNullOrEmpty(style.FontName))
            font.Name = style.FontName;
        if (style.FontSize is not null)
            font.Size = (float)style.FontSize.Value;
        if (style.Bold is not null)
            font.Bold = style.Bold.Value;
        if (style.Italic is not null)
            font.Italic = style.Italic.Value;
        if (!string.IsNullOrWhiteSpace(style.Color))
            font.Color.RGB = OpenEditCommon.HexToBgr(style.Color, "文字颜色");
    }

    // ── ppt-add-flowchart-slide ─────────────────────────────────

    public static Dictionary<string, object?> RunAddFlowchartSlide(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        var diagram = PptModels.ParseFrameworkDiagram(p);
        NativeCommon.EnsureFileLight(filePath);

        var app = OpenEditCommon.FindPowerPoint(filePath, out var presentation, out var progId)
            ?? throw new NativeRuntimeException(
                $"目标文件未在活动 PowerPoint/WPS 实例中打开: {filePath}；请先打开文件");
        if (presentation is null)
        {
            NativeCommon.ReleaseCom(app);
            throw new NativeRuntimeException(
                $"目标文件未在活动 PowerPoint/WPS 实例中打开: {filePath}");
        }
        var previousAlerts = GetDisplayAlerts(app);
        SetDisplayAlerts(app, 1); // ppAlertsNone
        try
        {
            OpenEditCommon.EnsureWritablePresentation(presentation);
            var initialSlideCount = (int)presentation.Slides.Count;
            try
            {
                dynamic slide = presentation.Slides.Add(initialSlideCount + 1, 12); // ppLayoutBlank
                try
                {
                    RenderFlowchart(slide, presentation, diagram);
                }
                finally
                {
                    NativeCommon.ReleaseCom(slide);
                }
                // Save 失败也必须回滚本次新增的幻灯片。
                OpenEditCommon.SavePresentation(presentation);
            }
            catch (Exception error)
            {
                try
                {
                    RollbackSlides(presentation, initialSlideCount);
                }
                catch (Exception rollbackError)
                {
                    throw new NativeRuntimeException(
                        $"流程图生成失败: {error.Message}；"
                        + $"回滚失败，请勿保存该演示文稿: {rollbackError.Message}");
                }
                throw new NativeRuntimeException(
                    $"流程图生成失败，已回滚新增幻灯片: {error.Message}");
            }

            var nodeCount = diagram.NodeCount;
            return OpenEditCommon.EditResult(filePath, nodeCount,
                $"已通过 {progId} 追加流程图幻灯片"
                + $"（{diagram.Stages.Count} 个阶段, {nodeCount} 个节点）");
        }
        finally
        {
            RestoreDisplayAlerts(app, previousAlerts);
            NativeCommon.ReleaseCom(presentation);
            NativeCommon.ReleaseCom(app);
        }
    }

    /// <summary>与 ppt_writer.create_framework_diagram 相同的坐标与配色绘制流程图。</summary>
    private static void RenderFlowchart(
        dynamic slide, dynamic presentation, PptFrameworkDiagram diagram)
    {
        var scheme = PptModels.ColorSchemes.GetValueOrDefault(diagram.ColorScheme)
                     ?? PptModels.ColorSchemes["professional-blue"];
        var arrowColor = string.IsNullOrWhiteSpace(diagram.ArrowColor)
            ? "94A3B8"
            : diagram.ArrowColor;
        var bodyFont = diagram.FontCn;

        var slideWidthIn = (float)presentation.PageSetup.SlideWidth / 72.0;
        var slideHeightIn = (float)presentation.PageSetup.SlideHeight / 72.0;

        slide.FollowMasterBackground = false;
        slide.Background.Fill.ForeColor.RGB =
            OpenEditCommon.HexToBgr(scheme.BackgroundColor, "背景色");

        var titleYOffset = 0.2;
        var accentBarY = 0.62;
        if (!string.IsNullOrEmpty(diagram.Title))
        {
            dynamic titleBox = AddTextbox(slide, 0.8, 0.15, 11.5, 0.45);
            try
            {
                dynamic textRange = titleBox.TextFrame.TextRange;
                textRange.Text = diagram.Title;
                textRange.ParagraphFormat.Alignment = 2;
                textRange.Font.Size = diagram.TitleFontSize;
                textRange.Font.Bold = true;
                textRange.Font.Name = bodyFont;
                textRange.Font.Color.RGB =
                    OpenEditCommon.HexToBgr(scheme.TitleColor, "标题颜色");
            }
            finally
            {
                NativeCommon.ReleaseCom(titleBox);
            }
            titleYOffset = 0.7;
        }

        if (!string.IsNullOrEmpty(diagram.Subtitle))
        {
            dynamic subtitleBox = AddTextbox(slide, 0.8, 0.5, 11.5, 0.3);
            try
            {
                dynamic textRange = subtitleBox.TextFrame.TextRange;
                textRange.Text = diagram.Subtitle;
                textRange.ParagraphFormat.Alignment = 2;
                textRange.Font.Size = 12;
                textRange.Font.Color.RGB =
                    OpenEditCommon.HexToBgr(scheme.BodyColor, "副标题颜色");
            }
            finally
            {
                NativeCommon.ReleaseCom(subtitleBox);
            }
            titleYOffset = 0.95;
            accentBarY = 0.9;
        }

        if (!string.IsNullOrEmpty(diagram.Title))
        {
            const double barWidth = 11.5 * 0.6;
            dynamic bar = AddShape(slide, 1,
                0.8 + (11.5 - barWidth) / 2, accentBarY, barWidth, 0.04);
            try
            {
                bar.Fill.ForeColor.RGB = OpenEditCommon.HexToBgr(
                    scheme.AccentColor, "强调色");
                bar.Line.Visible = 0;
            }
            finally
            {
                NativeCommon.ReleaseCom(bar);
            }
        }

        var stageCount = diagram.Stages.Count;
        if (stageCount == 0)
            return;

        const double contentLeft = 1.15;
        var contentRight = slideWidthIn - 0.3;
        var contentWidth = contentRight - contentLeft;
        var usableHeight = slideHeightIn - titleYOffset - 0.3;
        var stageUnitHeight = usableHeight / stageCount - diagram.StageGap;

        var positions = new Dictionary<string, (double X, double Y, double W, double H)>(
            StringComparer.Ordinal);

        for (var stageIndex = 0; stageIndex < diagram.Stages.Count; stageIndex++)
        {
            var stage = diagram.Stages[stageIndex];
            var stageY = titleYOffset + stageIndex * (stageUnitHeight + diagram.StageGap);
            var rawStageColor = string.IsNullOrWhiteSpace(stage.Color)
                ? scheme.AccentColor
                : stage.Color!;
            var stageColor = NormaliseHex(rawStageColor, "阶段颜色");

            // 阶段背景容器（极淡同色系底 + 同色系细边 + 阴影）
            dynamic background = AddShape(slide, 5,
                contentLeft + 0.55, stageY, contentWidth - 0.55, stageUnitHeight);
            try
            {
                background.Fill.ForeColor.RGB = OpenEditCommon.HexToBgr(
                    PptFileContent.HexMixWhite(stageColor, 0.96), "阶段背景色");
                background.Line.ForeColor.RGB = OpenEditCommon.HexToBgr(
                    PptFileContent.HexMixWhite(stageColor, 0.88), "阶段边框色");
                background.Line.Weight = 0.75;
                AddShadow(background, blurPt: 10, offsetPt: 2, transparency: 0.75);
            }
            finally
            {
                NativeCommon.ReleaseCom(background);
            }

            // 阶段标签（实心主色胶囊 + 白字）
            const double pillWidth = 1.25;
            const double pillHeight = 0.38;
            dynamic pill = AddShape(slide, 5,
                0.25, stageY + stageUnitHeight / 2 - pillHeight / 2,
                pillWidth, pillHeight);
            try
            {
                pill.Fill.ForeColor.RGB = OpenEditCommon.HexToBgr(stageColor, "阶段颜色");
                pill.Line.ForeColor.RGB = OpenEditCommon.HexToBgr(
                    PptFileContent.HexMixBlack(stageColor, 0.2), "阶段边框色");
                pill.Line.Weight = 0.75;
                AddShadow(pill);
                SetShapeText(pill, stage.Label, bodyFont, 12, true, "FFFFFF");
            }
            finally
            {
                NativeCommon.ReleaseCom(pill);
            }

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
                var nodeColor = NormaliseHex(
                    string.IsNullOrWhiteSpace(node.FillColor) ? stageColor : node.FillColor!,
                    "节点颜色");
                var shapeType = ShapeTypeMap.GetValueOrDefault(node.Style, 5);

                dynamic shape = AddShape(slide, shapeType,
                    nodeX, nodeCenterY, nodeWidth, nodeHeight);
                try
                {
                    shape.Fill.ForeColor.RGB =
                        OpenEditCommon.HexToBgr(nodeColor, "节点颜色");
                    shape.Line.ForeColor.RGB = OpenEditCommon.HexToBgr(
                        PptFileContent.HexMixBlack(nodeColor, 0.25), "节点边框色");
                    shape.Line.Weight = 1.0;
                    AddShadow(shape);
                    var nodeFontSize = node.FontSize == 0 ? 14 : node.FontSize;
                    SetShapeText(shape, node.Text, bodyFont, nodeFontSize,
                        node.Bold, "FFFFFF");
                }
                finally
                {
                    NativeCommon.ReleaseCom(shape);
                }

                // 步骤编号徽章（阶段内节点 ≥2 且非 terminator）
                if (nodes.Count >= 2 && node.Style != "terminator")
                {
                    const double badgeDiameter = 0.24;
                    dynamic badge = AddShape(slide, 9,
                        nodeX - 0.08, nodeCenterY - 0.08,
                        badgeDiameter, badgeDiameter);
                    try
                    {
                        badge.Fill.ForeColor.RGB =
                            OpenEditCommon.HexToBgr("FFFFFF", "徽章底色");
                        badge.Line.ForeColor.RGB =
                            OpenEditCommon.HexToBgr(nodeColor, "徽章边框色");
                        badge.Line.Weight = 1.0;
                        try
                        {
                            dynamic textRange = badge.TextFrame.TextRange;
                            textRange.Text = (nodeIndex + 1).ToString();
                            textRange.ParagraphFormat.Alignment = 2;
                            textRange.Font.Size = 9;
                            textRange.Font.Bold = true;
                            textRange.Font.Name = "Consolas";
                            textRange.Font.Color.RGB =
                                OpenEditCommon.HexToBgr(nodeColor, "徽章文字色");
                        }
                        catch { /* WPS 部分版本不支持徽章文字 */ }
                    }
                    finally
                    {
                        NativeCommon.ReleaseCom(badge);
                    }
                }

                positions[node.Id] = (nodeX, nodeCenterY, nodeWidth, nodeHeight);

                // 同阶段相邻节点自动箭头（中性灰弱化）
                if (nodeIndex > 0)
                {
                    var previousId = nodes[nodeIndex - 1].Id;
                    if (positions.TryGetValue(previousId, out var previous))
                    {
                        AddConnectorLine(slide, 1,
                            previous.X + previous.W,
                            previous.Y + previous.H / 2,
                            nodeX,
                            nodeCenterY + nodeHeight / 2,
                            arrowColor);
                    }
                }
            }
        }

        // ── 显式 edges（跨阶段连接） ───────────────────────────
        foreach (var edge in diagram.Edges)
        {
            if (!positions.TryGetValue(edge.FromId, out var from)
                || !positions.TryGetValue(edge.ToId, out var to))
                continue;
            var fromX = from.X + from.W / 2;
            var fromY = from.Y + from.H;
            var toX = to.X + to.W / 2;
            var toY = to.Y;
            var connectorType = ConnectorTypeMap.GetValueOrDefault(edge.EdgeStyle, 1);
            AddConnectorLine(slide, connectorType,
                fromX, fromY, toX, toY, arrowColor);

            if (!string.IsNullOrEmpty(edge.Label))
            {
                dynamic labelBox = AddTextbox(slide,
                    (fromX + toX) / 2 - 0.5,
                    (fromY + toY) / 2 - 0.25,
                    1.2, 0.35);
                try
                {
                    dynamic textRange = labelBox.TextFrame.TextRange;
                    textRange.Text = edge.Label;
                    textRange.ParagraphFormat.Alignment = 2;
                    textRange.Font.Size = 11;
                    textRange.Font.Italic = true;
                    textRange.Font.Color.RGB =
                        OpenEditCommon.HexToBgr(arrowColor, "箭头标签颜色");
                }
                finally
                {
                    NativeCommon.ReleaseCom(labelBox);
                }
            }
        }
    }

    private static dynamic AddTextbox(
        dynamic slide, double leftIn, double topIn, double widthIn, double heightIn) =>
        slide.Shapes.AddTextbox(1, Pt(leftIn), Pt(topIn), Pt(widthIn), Pt(heightIn));

    private static dynamic AddShape(
        dynamic slide, int shapeType,
        double leftIn, double topIn, double widthIn, double heightIn) =>
        slide.Shapes.AddShape(shapeType, Pt(leftIn), Pt(topIn), Pt(widthIn), Pt(heightIn));

    private static void AddConnectorLine(
        dynamic slide, int connectorType,
        double x1, double y1, double x2, double y2,
        string colorHex, double widthPt = 1.5, bool arrowhead = true)
    {
        dynamic connector = slide.Shapes.AddConnector(connectorType,
            Pt(x1), Pt(y1), Pt(x2), Pt(y2));
        try
        {
            connector.Line.ForeColor.RGB =
                OpenEditCommon.HexToBgr(colorHex, "箭头颜色");
            connector.Line.Weight = widthPt;
            if (arrowhead)
                connector.Line.EndArrowheadStyle = 2; // msoArrowheadTriangle
        }
        finally
        {
            NativeCommon.ReleaseCom(connector);
        }
    }

    private static void SetShapeText(
        dynamic shape, string text, string fontName, int fontSize, bool bold, string colorHex)
    {
        dynamic textRange = shape.TextFrame.TextRange;
        textRange.Text = text.Replace("\\n", "\n").Replace("\n", "\v");
        textRange.ParagraphFormat.Alignment = 2; // ppAlignCenter
        textRange.Font.Name = fontName;
        textRange.Font.Size = fontSize;
        textRange.Font.Bold = bold;
        textRange.Font.Color.RGB = OpenEditCommon.HexToBgr(colorHex, "形状文字色");
    }

    private static void AddShadow(
        dynamic shape, double blurPt = 8.0, double offsetPt = 3.0, double transparency = 0.5)
    {
        try
        {
            dynamic shadow = shape.Shadow;
            shadow.Visible = -1; // msoTrue
            shadow.Blur = blurPt;
            shadow.OffsetX = offsetPt;
            shadow.OffsetY = offsetPt;
            shadow.Transparency = transparency; // 0=不透明, 1=全透明
        }
        catch { /* WPS 兼容层可能不支持部分属性 */ }
    }

    private static string NormaliseHex(string value, string label)
    {
        var text = (value ?? "").Trim().TrimStart('#');
        if (text.Length == 8)
            text = text[2..];
        if (text.Length != 6
            || !uint.TryParse(text, System.Globalization.NumberStyles.HexNumber,
                System.Globalization.CultureInfo.InvariantCulture, out _))
            throw new NativeRuntimeException($"无效的{label}: '{value}'");
        return text.ToUpperInvariant();
    }

    private static string ComMultiline(string? text, string separator) =>
        (text ?? "").Replace("\\n", "\n").Replace("\n", separator);

    private static int Pt(double inches) => (int)(inches * 72);

    private static void RollbackSlides(dynamic presentation, int initialSlideCount)
    {
        while ((int)presentation.Slides.Count > initialSlideCount)
        {
            var count = (int)presentation.Slides.Count;
            dynamic slide = presentation.Slides.Item(count);
            try { slide.Delete(); }
            finally { NativeCommon.ReleaseCom(slide); }
        }
    }
}
