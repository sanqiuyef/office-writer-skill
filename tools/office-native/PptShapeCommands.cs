using System.Text.Json;

namespace OfficeNative;

/// <summary>
/// PPT 形状绘制：add/list。
/// add 优先走已打开演示的活动文档 COM（无需关闭文件），未打开时用 Open XML 文件级写入并备份。
/// </summary>
internal static class PptShapeCommands
{
    public static Dictionary<string, object?> RunShapes(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        // 活动文档已被 Office 锁定，不能先做解压校验；文件级分支会在 COM 未命中后完整校验。
        NativeCommon.EnsureFileLight(filePath, ".pptx", ".potx");
        var slideNumber = RequireSlideNumber(p);
        var action = NativeCommon.OptString(p, "action") ?? "add";
        return action switch
        {
            "add" => RunAdd(filePath, slideNumber, p),
            "list" => RunList(filePath, slideNumber),
            _ => throw new NativeUsageException("action 必须是 add/list"),
        };
    }

    private static int RequireSlideNumber(JsonElement p)
    {
        if (!p.TryGetProperty("slide", out var element)
            || element.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return 1;
        if (element.ValueKind != JsonValueKind.Number || !element.TryGetInt32(out var value))
            throw new NativeUsageException("slide 必须是整数（1 起）");
        if (value < 1)
            throw new NativeUsageException("slide 必须是 1 起的幻灯片序号");
        return value;
    }

    // ── add ─────────────────────────────────────────────────────

    private static Dictionary<string, object?> RunAdd(
        string filePath, int slideNumber, JsonElement p)
    {
        var shapeDef = PptModels.ParseShapeArgument(p);
        var fullPath = Path.GetFullPath(filePath);
        if (OpenEditPptCommands.TryAddShapeToOpenPresentation(
                fullPath, slideNumber, shapeDef) is { } comResult)
        {
            return comResult;
        }
        return AddToFile(fullPath, slideNumber, shapeDef);
    }

    private static Dictionary<string, object?> AddToFile(
        string fullPath, int slideNumber, PptSlideShape shapeDef)
    {
        NativeCommon.EnsureFile(fullPath, ".pptx", ".potx");
        var work = PptFileOoxml.WorkPath(fullPath);
        try
        {
            File.Copy(fullPath, work, overwrite: true);
            int shapeIndex;
            string shapeName;
            using (var document = PptFileOoxml.Open(work, true))
            {
                var slideParts = PptFileOoxml.SlideParts(PptFileOoxml.MainPart(document));
                if (slideNumber > slideParts.Count)
                    throw new NativeUsageException(
                        $"slide={slideNumber} 超出范围（共 {slideParts.Count} 张）");
                var slidePart = slideParts[slideNumber - 1];
                var slideDocument = PptFileOoxml.SlideDocument(slidePart);
                var writer = new PptSlideWriter(slidePart, slideDocument);
                writer.AddShape(shapeDef);
                writer.Save();
                document.Save();

                var shapes = PptFileOoxml.ShapeElements(slideDocument).ToList();
                var added = shapes[^1];
                shapeIndex = shapes.Count;
                shapeName = PptFileOoxml.ShapeName(added);
            }
            var backupPath = PptFileOoxml.ReplaceWithBackup(work, fullPath);
            return new Dictionary<string, object?>
            {
                ["success"] = true,
                ["action"] = "add",
                ["mode"] = "file",
                ["file_path"] = fullPath,
                ["slide"] = slideNumber,
                ["shape_index"] = shapeIndex,
                ["shape_name"] = shapeName,
                ["backup_path"] = backupPath,
                ["message"] = $"已向第 {slideNumber} 张幻灯片添加形状（文件级写入）",
            };
        }
        finally
        {
            PptFileOoxml.EnsureCleanup(work);
        }
    }

    // ── list ────────────────────────────────────────────────────

    private static Dictionary<string, object?> RunList(string filePath, int slideNumber)
    {
        var fullPath = Path.GetFullPath(filePath);
        if (TryListOpenPresentation(fullPath, slideNumber) is { } openResult)
            return openResult;

        NativeCommon.EnsureFile(fullPath, ".pptx", ".potx");
        using var document = PptFileOoxml.Open(fullPath, false);
        var slideParts = PptFileOoxml.SlideParts(PptFileOoxml.MainPart(document));
        if (slideNumber > slideParts.Count)
            throw new NativeUsageException(
                $"slide={slideNumber} 超出范围（共 {slideParts.Count} 张）");
        var slideDocument = PptFileOoxml.SlideDocument(slideParts[slideNumber - 1]);
        var context = PptFileOoxml.ContextFor(slideParts[slideNumber - 1]);
        var shapes = PptFileOoxml.ShapeElements(slideDocument).ToList();
        var items = new List<Dictionary<string, object?>>();
        for (var index = 0; index < shapes.Count; index++)
        {
            var shape = shapes[index];
            var box = PptFileOoxml.GetShapeBox(shape, context);
            items.Add(new Dictionary<string, object?>
            {
                ["index"] = index + 1,
                ["shape_id"] = PptFileOoxml.ShapeId(shape),
                ["name"] = PptFileOoxml.ShapeName(shape),
                ["type"] = shape.Name.LocalName,
                ["preset"] = shape.Element(PptFileOoxml.P + "spPr")
                    ?.Element(PptFileOoxml.A + "prstGeom")?.Attribute("prst")?.Value,
                ["x"] = box is null ? null : Round(PptFileOoxml.ToInches(box.Value.X)),
                ["y"] = box is null ? null : Round(PptFileOoxml.ToInches(box.Value.Y)),
                ["width"] = box is null ? null : Round(PptFileOoxml.ToInches(box.Value.Width)),
                ["height"] = box is null ? null : Round(PptFileOoxml.ToInches(box.Value.Height)),
                ["text"] = PptFileOoxml.TxBodyText(PptFileOoxml.TextBodyOf(shape)),
            });
        }
        return ListResult(fullPath, slideNumber, "file", items);
    }

    private static Dictionary<string, object?>? TryListOpenPresentation(
        string filePath, int slideNumber)
    {
        var app = OpenEditCommon.FindPowerPoint(filePath, out var presentation, out var progId);
        if (app is null || presentation is null)
        {
            NativeCommon.ReleaseCom(presentation);
            NativeCommon.ReleaseCom(app);
            return null;
        }
        dynamic? slide = null;
        dynamic? shapes = null;
        try
        {
            var slideCount = (int)presentation.Slides.Count;
            if (slideNumber > slideCount)
                throw new NativeUsageException(
                    $"slide={slideNumber} 超出范围（共 {slideCount} 张）");
            slide = presentation.Slides.Item(slideNumber);
            shapes = slide.Shapes;
            var items = new List<Dictionary<string, object?>>();
            var count = (int)shapes.Count;
            for (var index = 1; index <= count; index++)
            {
                dynamic shape = shapes.Item(index);
                try
                {
                    var hasText = (int)shape.HasTextFrame != 0
                                  && (int)shape.TextFrame.HasText != 0;
                    items.Add(new Dictionary<string, object?>
                    {
                        ["index"] = index,
                        ["shape_id"] = (int)shape.Id,
                        ["name"] = (string)shape.Name,
                        ["type"] = (int)shape.Type,
                        ["x"] = Round((double)shape.Left / 72.0),
                        ["y"] = Round((double)shape.Top / 72.0),
                        ["width"] = Round((double)shape.Width / 72.0),
                        ["height"] = Round((double)shape.Height / 72.0),
                        ["text"] = hasText
                            ? ((string)shape.TextFrame.TextRange.Text).Replace("\r", "\n")
                            : "",
                    });
                }
                finally
                {
                    NativeCommon.ReleaseCom(shape);
                }
            }
            return ListResult(filePath, slideNumber, "open_document", items, progId);
        }
        finally
        {
            NativeCommon.ReleaseCom(shapes);
            NativeCommon.ReleaseCom(slide);
            NativeCommon.ReleaseCom(presentation);
            NativeCommon.ReleaseCom(app);
        }
    }

    private static Dictionary<string, object?> ListResult(
        string filePath, int slideNumber, string mode,
        List<Dictionary<string, object?>> items, string? engine = null)
    {
        var data = new Dictionary<string, object?>
        {
            ["success"] = true,
            ["action"] = "list",
            ["mode"] = mode,
            ["file_path"] = filePath,
            ["slide"] = slideNumber,
            ["count"] = items.Count,
            ["shapes"] = items,
        };
        if (engine is not null)
            data["engine"] = engine;
        return data;
    }

    private static double Round(double value) => Math.Round(value, 4);
}
