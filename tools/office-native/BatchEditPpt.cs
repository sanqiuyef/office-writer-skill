using System.Globalization;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeNative;

/// <summary>PowerPoint 批量编辑：/slide[N]/shape[N]、幻灯片属性、备注与背景。</summary>
internal static class BatchEditPpt
{
    private static readonly System.Text.RegularExpressions.Regex SlidePath =
        new(@"^/slide\[(\d+)\]$");
    private static readonly System.Text.RegularExpressions.Regex ShapePath =
        new(@"^/slide\[(\d+)\]/shape\[(\d+)\]$");

    private static readonly HashSet<string> ShapeProps = new(StringComparer.Ordinal)
    {
        "text", "x", "y", "width", "height", "fill", "font_color",
        "font_size", "bold", "italic",
    };

    private static readonly HashSet<string> SlideProps = new(StringComparer.Ordinal)
    {
        "title", "notes", "background",
    };

    private sealed class PptTarget
    {
        public string Kind { get; init; } = "";
        public int SlideIndex { get; init; }
        public int ShapeIndex { get; init; }
        public SlidePart? SlidePart { get; init; }
        public XDocument? Document { get; init; }
        public XElement? Shape { get; init; }
    }

    public static List<string> Apply(
        string workPath, IReadOnlyList<OfficeOperation> operations)
    {
        using var document = PresentationDocument.Open(workPath, true);
        var presentationPart = document.PresentationPart
                               ?? throw new NativeRuntimeException(
                                   "演示文稿没有 PresentationPart");
        var messages = new List<string>(operations.Count);
        foreach (var operation in operations)
        {
            try
            {
                messages.Add(ApplyOperation(presentationPart, operation));
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
        PresentationOf(presentationPart).Save();
        return messages;
    }

    private static P.Presentation PresentationOf(PresentationPart part) =>
        part.Presentation
        ?? throw new NativeRuntimeException("演示文稿缺少 Presentation 根节点");

    private static string ApplyOperation(
        PresentationPart presentationPart, OfficeOperation operation)
    {
        return operation.Op switch
        {
            "set" => Set(presentationPart, operation),
            "add" => Add(presentationPart, operation),
            "remove" => Remove(presentationPart, operation),
            "move" => Move(presentationPart, operation),
            "replace" => Replace(presentationPart, operation),
            _ => throw new BatchEditException(
                "unsupported_operation", $"不支持的操作: {operation.Op}",
                operation.Path),
        };
    }

    // ── 路径解析 ─────────────────────────────────────────────────

    private static PptTarget ResolveTarget(
        PresentationPart presentationPart, string path)
    {
        if (path == "/")
            return new PptTarget { Kind = "presentation" };

        var shapeMatch = ShapePath.Match(path);
        if (shapeMatch.Success)
        {
            var slideIndex = ParseIndex(shapeMatch.Groups[1].Value, path);
            var (slidePart, slideCount) = ResolveSlide(presentationPart,
                slideIndex, path);
            var document = PptFileOoxml.SlideDocument(slidePart);
            var shapes = PptFileOoxml.ShapeElements(document).ToList();
            var shapeIndex = ParseIndex(shapeMatch.Groups[2].Value, path);
            if (shapeIndex < 1 || shapeIndex > shapes.Count)
                throw new BatchEditException(
                    "not_found", $"路径索引超出范围: {path}", path,
                    suggestion: $"有效索引为 1..{shapes.Count}");
            return new PptTarget
            {
                Kind = "shape",
                SlideIndex = slideIndex,
                ShapeIndex = shapeIndex,
                SlidePart = slidePart,
                Document = document,
                Shape = shapes[shapeIndex - 1],
            };
        }

        var slideMatch = SlidePath.Match(path);
        if (slideMatch.Success)
        {
            var slideIndex = ParseIndex(slideMatch.Groups[1].Value, path);
            var (slidePart, _) = ResolveSlide(presentationPart, slideIndex, path);
            return new PptTarget
            {
                Kind = "slide",
                SlideIndex = slideIndex,
                SlidePart = slidePart,
            };
        }

        throw new BatchEditException(
            "invalid_path", $"不支持的 PPT 路径: {path}", path,
            validValues: ["/", "/slide[N]", "/slide[N]/shape[N]"]);
    }

    private static int ParseIndex(string digits, string path) =>
        int.TryParse(digits, NumberStyles.Integer, CultureInfo.InvariantCulture,
            out var index) && index >= 1
            ? index
            : throw new BatchEditException(
                "not_found", $"路径索引超出范围: {path}", path);

    private static (SlidePart Part, int Count) ResolveSlide(
        PresentationPart presentationPart, int index, string path)
    {
        var slides = PptFileOoxml.SlideParts(presentationPart);
        if (index < 1 || index > slides.Count)
            throw new BatchEditException(
                "not_found", $"路径索引超出范围: {path}", path,
                suggestion: $"有效索引为 1..{slides.Count}");
        return (slides[index - 1], slides.Count);
    }

    // ── set ──────────────────────────────────────────────────────

    private static string Set(
        PresentationPart presentationPart, OfficeOperation operation)
    {
        var target = ResolveTarget(presentationPart, operation.Path);
        switch (target.Kind)
        {
            case "shape":
                SetShape(target, operation);
                break;
            case "slide":
                SetSlide(target, operation);
                break;
            default:
                throw new BatchEditException(
                    "unsupported_operation", "set 不支持整个演示文稿",
                    operation.Path);
        }
        return $"已更新 {operation.Path}";
    }

    private static void SetShape(PptTarget target, OfficeOperation operation)
    {
        BatchEditJson.EnsureAllowed(operation, ShapeProps, "PPT 形状");
        var shape = target.Shape
                    ?? throw new NativeRuntimeException("形状不存在");
        var textBody = PptFileOoxml.TextBodyOf(shape);

        if (BatchEditJson.TryGet(operation, "text", out var text))
        {
            if (textBody is null)
                throw new BatchEditException(
                    "unsupported_operation", "目标形状不包含文本框",
                    operation.Path);
            SetTextBody(textBody, BatchEditJson.PythonString(text));
        }

        SetShapeGeometry(shape, operation);
        if (BatchEditJson.TryGet(operation, "fill", out var fill))
            SetShapeFill(shape, BatchEditJson.RequireHexColor(
                operation, "fill", fill), operation.Path);

        if (textBody is not null)
            SetRunFonts(textBody, operation);

        PptFileOoxml.SaveSlide(target.SlidePart!, target.Document!);
    }

    private static void SetSlide(PptTarget target, OfficeOperation operation)
    {
        BatchEditJson.EnsureAllowed(operation, SlideProps, "PPT 幻灯片");
        var document = PptFileOoxml.SlideDocument(target.SlidePart!);

        if (BatchEditJson.TryGet(operation, "title", out var title))
        {
            var titleShape = FindTitlePlaceholder(document);
            if (titleShape is null)
                throw new BatchEditException(
                    "not_found", "幻灯片没有标题占位符", operation.Path);
            var textBody = PptFileOoxml.TextBodyOf(titleShape);
            if (textBody is null)
                throw new BatchEditException(
                    "unsupported_operation", "标题占位符不包含文本框",
                    operation.Path);
            SetTextBody(textBody, BatchEditJson.PythonString(title));
        }

        if (BatchEditJson.TryGet(operation, "notes", out var notes))
            SetNotes(target.SlidePart!, BatchEditJson.PythonString(notes));

        if (BatchEditJson.TryGet(operation, "background", out var background))
            SetBackground(document, BatchEditJson.RequireHexColor(
                operation, "background", background));

        PptFileOoxml.SaveSlide(target.SlidePart!, document);
    }

    private static XElement? FindTitlePlaceholder(XDocument document) =>
        PptFileOoxml.ShapeElements(document).FirstOrDefault(shape =>
        {
            var info = PptFileOoxml.PlaceholderInfo(shape);
            return info?.Type is { } type
                   && (string.Equals(type, "title", StringComparison.OrdinalIgnoreCase)
                       || string.Equals(type, "ctrTitle",
                           StringComparison.OrdinalIgnoreCase));
        });

    // ── add ──────────────────────────────────────────────────────

    private static string Add(
        PresentationPart presentationPart, OfficeOperation operation)
    {
        var target = ResolveTarget(presentationPart, operation.Path);
        if (target.Kind == "presentation")
        {
            var elementType = BatchEditJson.RequireElementType(
                operation, "PPT", ["slide"]);
            if (elementType != "slide")
                throw new BatchEditException(
                    "invalid_value", "演示文稿根节点只能添加 slide",
                    operation.Path, validValues: ["slide"]);

            var layoutIndex = BatchEditJson.OptInt(operation, "layout_index", 6);
            var layoutCount = PptFileOoxml.LayoutParts(presentationPart).Count;
            if (layoutIndex < 0 || layoutIndex >= layoutCount)
                throw new BatchEditException(
                    "invalid_value", $"layout_index 超出范围: {layoutIndex}",
                    operation.Path);
            var slideWriter = PptFileContent.AddSlide(
                presentationPart, layoutIndex, null);
            if (BatchEditJson.TryGet(operation, "title", out var title))
            {
                var titleShape = slideWriter.FindPlaceholderByType("title", "ctrTitle");
                if (titleShape is not null)
                {
                    var textBody = PptFileOoxml.TextBodyOf(titleShape);
                    if (textBody is not null)
                    {
                        SetTextBody(textBody,
                            BatchEditJson.PythonString(title));
                        slideWriter.Save();
                    }
                }
            }
            return "已添加 slide";
        }

        if (target.Kind != "slide")
            throw new BatchEditException(
                "invalid_path", "PPT 元素应添加到幻灯片路径",
                operation.Path);

        var type = BatchEditJson.RequireElementType(
            operation, "PPT", ["textbox", "shape"]);
        var x = BatchEditJson.OptDouble(operation, "x", 1);
        var y = BatchEditJson.OptDouble(operation, "y", 1);
        var width = BatchEditJson.OptDouble(operation, "width", 4);
        var height = BatchEditJson.OptDouble(operation, "height", 1);
        var text = BatchEditJson.PythonStringProp(operation, "text");
        var document = PptFileOoxml.SlideDocument(target.SlidePart!);
        var writer = new PptSlideWriter(target.SlidePart!, document);

        switch (type)
        {
            case "textbox":
                writer.AddTextBox(x, y, width, height, BuildParagraphs(text));
                break;
            case "shape":
            {
                var shapeType = BatchEditJson.PythonStringProp(
                    operation, "shape_type", "rect");
                var preset = shapeType switch
                {
                    "rect" => "rect",
                    "rounded_rect" => "roundRect",
                    "oval" => "ellipse",
                    "arrow" => "rightArrow",
                    _ => throw new BatchEditException(
                        "invalid_value", $"无效的 shape_type: {shapeType}",
                        operation.Path,
                        validValues:
                        ["rect", "rounded_rect", "oval", "arrow"]),
                };
                var shape = BuildAutoShape(writer, preset, text,
                    x, y, width, height);
                writer.Append(shape);
                break;
            }
            default:
                throw new BatchEditException(
                    "invalid_value", $"不支持的 PPT 元素类型: {type}",
                    operation.Path, validValues: ["textbox", "shape"]);
        }
        writer.Save();
        return $"已添加 {type}";
    }

    private static XElement BuildAutoShape(
        PptSlideWriter writer, string preset, string text,
        double x, double y, double width, double height)
    {
        var id = writer.TakeId();
        return new XElement(PptFileOoxml.P + "sp",
            new XElement(PptFileOoxml.P + "nvSpPr",
                new XElement(PptFileOoxml.P + "cNvPr",
                    new XAttribute("id", id),
                    new XAttribute("name", $"Shape {id}")),
                new XElement(PptFileOoxml.P + "cNvSpPr"),
                new XElement(PptFileOoxml.P + "nvPr")),
            new XElement(PptFileOoxml.P + "spPr",
                new XElement(PptFileOoxml.A + "xfrm",
                    new XElement(PptFileOoxml.A + "off",
                        new XAttribute("x", PptFileOoxml.Inches(x)),
                        new XAttribute("y", PptFileOoxml.Inches(y))),
                    new XElement(PptFileOoxml.A + "ext",
                        new XAttribute("cx", PptFileOoxml.Inches(width)),
                        new XAttribute("cy", PptFileOoxml.Inches(height)))),
                new XElement(PptFileOoxml.A + "prstGeom",
                    new XAttribute("prst", preset),
                    new XElement(PptFileOoxml.A + "avLst"))),
            PptFileContent.ShapeStyle(),
            new XElement(PptFileOoxml.P + "txBody",
                new XElement(PptFileOoxml.A + "bodyPr",
                    new XAttribute("wrap", "square")),
                new XElement(PptFileOoxml.A + "lstStyle"),
                BuildParagraphs(text)));
    }

    // ── remove ───────────────────────────────────────────────────

    private static string Remove(
        PresentationPart presentationPart, OfficeOperation operation)
    {
        var target = ResolveTarget(presentationPart, operation.Path);
        switch (target.Kind)
        {
            case "shape":
                target.Shape!.Remove();
                PptFileOoxml.SaveSlide(target.SlidePart!, target.Document!);
                break;
            case "slide":
            {
                var slidePart = target.SlidePart!;
                var slideIdList = PresentationOf(presentationPart).SlideIdList;
                if (slideIdList is null)
                    throw new BatchEditException(
                        "not_found", "演示文稿没有幻灯片列表",
                        operation.Path);
                var relationshipId = presentationPart.GetIdOfPart(slidePart);
                var slideId = slideIdList.Elements<P.SlideId>()
                    .FirstOrDefault(item =>
                        item.RelationshipId?.Value == relationshipId);
                if (slideId is null)
                    throw new BatchEditException(
                        "not_found", "无法定位幻灯片记录", operation.Path);
                slideId.Remove();
                presentationPart.DeletePart(slidePart);
                break;
            }
            default:
                throw new BatchEditException(
                    "unsupported_operation", "不能删除整个演示文稿",
                    operation.Path);
        }
        return $"已删除 {operation.Path}";
    }

    // ── move ─────────────────────────────────────────────────────

    private static string Move(
        PresentationPart presentationPart, OfficeOperation operation)
    {
        if (string.IsNullOrEmpty(operation.TargetPath))
            throw new BatchEditException(
                "invalid_value", "move 操作需要 target_path", operation.Path);
        var source = ResolveTarget(presentationPart, operation.Path);
        var target = ResolveTarget(presentationPart, operation.TargetPath);
        if (source.Kind == "presentation"
            || source.Kind != target.Kind
            || (source.Kind == "shape"
                && source.SlideIndex == target.SlideIndex
                && source.ShapeIndex == target.ShapeIndex)
            || (source.Kind == "slide"
                && source.SlideIndex == target.SlideIndex))
            throw new BatchEditException(
                "invalid_path",
                "源路径与 target_path 必须是不同的同类元素",
                operation.TargetPath);

        if (source.Kind == "shape")
        {
            XElement moveTarget;
            if (source.SlideIndex == target.SlideIndex)
            {
                var shapes = PptFileOoxml.ShapeElements(source.Document!)
                    .ToList();
                moveTarget = shapes[target.ShapeIndex - 1];
            }
            else
            {
                moveTarget = target.Shape!;
                MigrateShapeRelationships(source.SlidePart!, target.SlidePart!,
                    source.Shape!, operation.Path);
            }
            source.Shape!.Remove();
            if (operation.Position == "before")
                moveTarget.AddBeforeSelf(source.Shape);
            else
                moveTarget.AddAfterSelf(source.Shape);
            PptFileOoxml.SaveSlide(source.SlidePart!, source.Document!);
            if (source.SlideIndex != target.SlideIndex)
                PptFileOoxml.SaveSlide(target.SlidePart!, target.Document!);
        }
        else if (source.Kind == "slide")
        {
            var slideIdList = PresentationOf(presentationPart).SlideIdList
                              ?? throw new BatchEditException(
                                  "not_found", "演示文稿没有幻灯片列表",
                                  operation.Path);
            var ids = slideIdList.Elements<P.SlideId>().ToList();
            var sourceRelationship = presentationPart.GetIdOfPart(source.SlidePart!);
            var targetRelationship = presentationPart.GetIdOfPart(target.SlidePart!);
            var sourceId = ids.FirstOrDefault(item =>
                item.RelationshipId?.Value == sourceRelationship)
                ?? throw new BatchEditException(
                    "not_found", "无法定位源幻灯片记录", operation.Path);
            var targetId = ids.FirstOrDefault(item =>
                item.RelationshipId?.Value == targetRelationship)
                ?? throw new BatchEditException(
                    "not_found", "无法定位目标幻灯片记录", operation.TargetPath);
            sourceId.Remove();
            if (operation.Position == "before")
                targetId.InsertBeforeSelf(sourceId);
            else
                targetId.InsertAfterSelf(sourceId);
            PresentationOf(presentationPart).Save();
        }
        else
        {
            throw new BatchEditException(
                "unsupported_operation", "只能移动幻灯片或形状",
                operation.Path);
        }
        return $"已移动 {operation.Path}";
    }

    /// <summary>
    /// 跨 slide 移动形状前，把形状内引用的关系迁移到目标 slide：
    /// 图片/图表/内嵌对象/图表数据等内部部件建立到同一部件的新关系，
    /// 超链接/外链按原 URI 重建，随后改写 XML 中的 rId。任何无法迁移的
    /// 关系都直接失败，避免留下 PowerPoint 需要修复的悬空 rId。
    /// </summary>
    private static void MigrateShapeRelationships(
        SlidePart sourcePart, SlidePart targetPart, XElement shape, string path)
    {
        var attributes = shape.DescendantsAndSelf()
            .Attributes()
            .Where(attribute => attribute.Name.Namespace == PptFileOoxml.R
                                && !string.IsNullOrEmpty(attribute.Value))
            .ToList();
        if (attributes.Count == 0)
            return;

        var mapping = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var attribute in attributes)
        {
            if (mapping.ContainsKey(attribute.Value))
                continue;
            mapping[attribute.Value] = MigrateRelationship(
                sourcePart, targetPart, attribute.Value, path);
        }
        foreach (var attribute in attributes)
            attribute.Value = mapping[attribute.Value];
    }

    private static string MigrateRelationship(
        SlidePart sourcePart, SlidePart targetPart, string relationshipId,
        string path)
    {
        var hyperlink = sourcePart.HyperlinkRelationships
            .FirstOrDefault(item => item.Id == relationshipId);
        if (hyperlink is not null)
            return EnsureHyperlinkRelationship(targetPart, hyperlink.Uri);

        var external = sourcePart.ExternalRelationships
            .FirstOrDefault(item => item.Id == relationshipId);
        if (external is not null)
            return EnsureExternalRelationship(targetPart, external);

        if (sourcePart.DataPartReferenceRelationships
            .Any(item => item.Id == relationshipId))
            throw new BatchEditException(
                "unsupported_operation",
                $"关系 {relationshipId} 指向音视频等数据部件，暂不支持跨 slide 迁移",
                path);

        if (!sourcePart.TryGetPartById(relationshipId, out var part))
            throw new BatchEditException(
                "unsupported_operation",
                $"形状引用的关系 {relationshipId} 在源幻灯片中不存在，无法跨 slide 迁移",
                path);

        if (part is ImagePart or ChartPart or SlidePart
            or EmbeddedPackagePart or DiagramDataPart
            or DiagramLayoutDefinitionPart or DiagramStylePart
            or DiagramColorsPart)
            return EnsurePartRelationship(targetPart, part);

        throw new BatchEditException(
            "unsupported_operation",
            $"形状引用的关系类型 {part.GetType().Name} 不支持跨 slide 迁移",
            path);
    }

    private static string EnsureHyperlinkRelationship(SlidePart targetPart, Uri uri)
    {
        var existing = targetPart.HyperlinkRelationships
            .FirstOrDefault(item => item.Uri == uri);
        return existing?.Id
               ?? targetPart.AddHyperlinkRelationship(uri, true).Id;
    }

    private static string EnsureExternalRelationship(
        SlidePart targetPart, ExternalRelationship relationship)
    {
        var existing = targetPart.ExternalRelationships.FirstOrDefault(item =>
            item.Uri == relationship.Uri
            && item.RelationshipType == relationship.RelationshipType);
        return existing?.Id
               ?? targetPart.AddExternalRelationship(
                   relationship.RelationshipType, relationship.Uri).Id;
    }

    private static string EnsurePartRelationship(
        SlidePart targetPart, OpenXmlPart part)
    {
        foreach (var pair in targetPart.Parts)
        {
            if (ReferenceEquals(pair.OpenXmlPart, part))
                return pair.RelationshipId;
        }
        return targetPart.CreateRelationshipToPart(part);
    }

    // ── replace ──────────────────────────────────────────────────

    private static string Replace(
        PresentationPart presentationPart, OfficeOperation operation)
    {
        var find = BatchEditJson.PythonStringProp(operation, "find");
        var replacement = BatchEditJson.PythonStringProp(operation, "replace");
        if (find.Length == 0)
            throw new BatchEditException(
                "invalid_value", "replace 操作的 find 不能为空", operation.Path);

        var target = ResolveTarget(presentationPart, operation.Path);
        var count = 0;
        if (target.Kind == "shape")
        {
            count = ReplaceInShapes([target.Shape!], find, replacement);
            PptFileOoxml.SaveSlide(target.SlidePart!, target.Document!);
        }
        else if (target.Kind == "slide")
        {
            var document = PptFileOoxml.SlideDocument(target.SlidePart!);
            count = ReplaceInShapes(
                PptFileOoxml.ShapeElements(document).ToList(), find, replacement);
            PptFileOoxml.SaveSlide(target.SlidePart!, document);
        }
        else
        {
            foreach (var slidePart in PptFileOoxml.SlideParts(presentationPart))
            {
                var document = PptFileOoxml.SlideDocument(slidePart);
                var changed = ReplaceInShapes(
                    PptFileOoxml.ShapeElements(document).ToList(),
                    find, replacement);
                if (changed > 0)
                {
                    count += changed;
                    PptFileOoxml.SaveSlide(slidePart, document);
                }
            }
        }
        return $"已替换 {count} 处文本";
    }

    private static int ReplaceInShapes(
        IEnumerable<XElement> shapes, string find, string replacement)
    {
        var count = 0;
        foreach (var shape in shapes)
        {
            var textBody = PptFileOoxml.TextBodyOf(shape);
            if (textBody is null)
                continue;
            var text = PptFileOoxml.TextFrameText(shape);
            var occurrences = CountOccurrences(text, find);
            if (occurrences == 0)
                continue;
            count += occurrences;
            SetTextBody(textBody,
                text.Replace(find, replacement, StringComparison.Ordinal));
        }
        return count;
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

    // ── 文本/备注/背景 ───────────────────────────────────────────

    private static void SetTextBody(XElement textBody, string text)
    {
        foreach (var paragraph in textBody
                     .Elements(PptFileOoxml.A + "p").ToList())
            paragraph.Remove();
        var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        if (normalized.Length == 0)
        {
            textBody.Add(new XElement(PptFileOoxml.A + "p"));
            return;
        }
        foreach (var line in normalized.Split('\n'))
            textBody.Add(BuildParagraph(line));
    }

    private static List<XElement> BuildParagraphs(string text)
    {
        var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        if (normalized.Length == 0)
            return [new XElement(PptFileOoxml.A + "p")];
        return normalized.Split('\n')
            .Select(BuildParagraph)
            .ToList();
    }

    private static XElement BuildParagraph(string line)
    {
        var paragraph = new XElement(PptFileOoxml.A + "p");
        if (line.Length == 0)
            return paragraph;
        var segments = line.Split('\v');
        for (var index = 0; index < segments.Length; index++)
        {
            if (index > 0)
                paragraph.Add(new XElement(PptFileOoxml.A + "br"));
            if (segments[index].Length > 0)
                paragraph.Add(new XElement(PptFileOoxml.A + "r",
                    new XElement(PptFileOoxml.A + "t",
                        new XAttribute(XNamespace.Xml + "space", "preserve"),
                        segments[index])));
        }
        return paragraph;
    }

    private static void SetShapeGeometry(
        XElement shape, OfficeOperation operation)
    {
        var hasX = BatchEditJson.TryGet(operation, "x", out var xValue);
        var hasY = BatchEditJson.TryGet(operation, "y", out var yValue);
        var hasWidth = BatchEditJson.TryGet(operation, "width", out var widthValue);
        var hasHeight = BatchEditJson.TryGet(operation, "height", out var heightValue);
        if (!hasX && !hasY && !hasWidth && !hasHeight)
            return;

        var transform = EnsureTransform(shape, operation.Path);
        if (hasX || hasY)
        {
            var offset = transform.Element(PptFileOoxml.A + "off");
            if (offset is null)
            {
                offset = new XElement(PptFileOoxml.A + "off",
                    new XAttribute("x", "0"), new XAttribute("y", "0"));
                transform.AddFirst(offset);
            }
            if (hasX)
                offset.SetAttributeValue("x", PptFileOoxml.Inches(
                    BatchEditJson.ToDoubleValue(xValue, "x", operation.Path)));
            if (hasY)
                offset.SetAttributeValue("y", PptFileOoxml.Inches(
                    BatchEditJson.ToDoubleValue(yValue, "y", operation.Path)));
        }
        if (hasWidth || hasHeight)
        {
            var extent = transform.Element(PptFileOoxml.A + "ext");
            if (extent is null)
            {
                extent = new XElement(PptFileOoxml.A + "ext",
                    new XAttribute("cx", "0"), new XAttribute("cy", "0"));
                var offset = transform.Element(PptFileOoxml.A + "off");
                if (offset is not null)
                    offset.AddAfterSelf(extent);
                else
                    transform.AddFirst(extent);
            }
            if (hasWidth)
                extent.SetAttributeValue("cx", PptFileOoxml.Inches(
                    BatchEditJson.ToDoubleValue(widthValue, "width",
                        operation.Path)));
            if (hasHeight)
                extent.SetAttributeValue("cy", PptFileOoxml.Inches(
                    BatchEditJson.ToDoubleValue(heightValue, "height",
                        operation.Path)));
        }
    }

    private static XElement EnsureTransform(XElement shape, string path)
    {
        if (shape.Name == PptFileOoxml.P + "graphicFrame")
        {
            var transform = shape.Element(PptFileOoxml.P + "xfrm");
            if (transform is null)
            {
                transform = new XElement(PptFileOoxml.P + "xfrm");
                shape.AddFirst(transform);
            }
            return transform;
        }
        if (shape.Name == PptFileOoxml.P + "grpSp")
        {
            var properties = shape.Element(PptFileOoxml.P + "grpSpPr");
            if (properties is null)
            {
                properties = new XElement(PptFileOoxml.P + "grpSpPr");
                shape.AddFirst(properties);
            }
            var transform = properties.Element(PptFileOoxml.A + "xfrm");
            if (transform is null)
            {
                transform = new XElement(PptFileOoxml.A + "xfrm");
                properties.AddFirst(transform);
            }
            return transform;
        }
        var shapeProperties = shape.Element(PptFileOoxml.P + "spPr");
        if (shapeProperties is null)
            throw new BatchEditException(
                "unsupported_operation", "目标形状不支持几何设置", path);
        var existing = shapeProperties.Element(PptFileOoxml.A + "xfrm");
        if (existing is null)
        {
            existing = new XElement(PptFileOoxml.A + "xfrm");
            shapeProperties.AddFirst(existing);
        }
        return existing;
    }

    private static void SetShapeFill(
        XElement shape, string hex, string path)
    {
        // 只有普通形状/连接线的 p:spPr 可以作为填充目标；图片的 p:blipFill
        // 是图片数据本身，不能被“填充”覆盖或删除。
        if (shape.Name != PptFileOoxml.P + "sp"
            && shape.Name != PptFileOoxml.P + "cxnSp")
            throw new BatchEditException(
                "unsupported_operation",
                $"fill 仅支持普通形状或连接线，目标元素为 {shape.Name.LocalName}",
                path);
        var properties = shape.Element(PptFileOoxml.P + "spPr");
        if (properties is null)
            throw new BatchEditException(
                "unsupported_operation", "目标形状不支持填充", path);
        RemoveFillElements(properties);
        InsertBefore(
            properties,
            PptFileOoxml.SolidFill(hex),
            ["ln", "effectLst", "effectDag", "scene3d", "sp3d", "extLst"]);
    }

    private static void SetRunFonts(
        XElement textBody, OfficeOperation operation)
    {
        var hasColor = BatchEditJson.TryGet(operation, "font_color",
            out var colorValue);
        var hasSize = BatchEditJson.TryGet(operation, "font_size",
            out var sizeValue);
        var hasBold = BatchEditJson.TryGet(operation, "bold", out var boldValue);
        var hasItalic = BatchEditJson.TryGet(operation, "italic",
            out var italicValue);
        if (!hasColor && !hasSize && !hasBold && !hasItalic)
            return;

        var hex = hasColor
            ? BatchEditJson.RequireHexColor(operation, "font_color", colorValue)
            : null;
        var size = hasSize
            ? (long)Math.Round(BatchEditJson.ToDoubleValue(
                sizeValue, "font_size", operation.Path) * 100)
            : (long?)null;
        var bold = hasBold ? BatchEditJson.Truthy(boldValue) : (bool?)null;
        var italic = hasItalic ? BatchEditJson.Truthy(italicValue) : (bool?)null;

        foreach (var paragraph in textBody
                     .Elements(PptFileOoxml.A + "p"))
        foreach (var run in paragraph
                     .Elements(PptFileOoxml.A + "r"))
        {
            var properties = run.Element(PptFileOoxml.A + "rPr");
            if (properties is null)
            {
                properties = new XElement(PptFileOoxml.A + "rPr");
                run.AddFirst(properties);
            }
            if (hex is not null)
            {
                RemoveFillElements(properties);
                InsertBefore(properties, PptFileOoxml.SolidFill(hex),
                [
                    "effectLst", "effectDag", "highlight", "uLnTx", "uLn",
                    "uFillTx", "uFill", "latin", "ea", "cs", "sym",
                    "hlinkClick", "hlinkMouseOver", "rtl", "extLst",
                ]);
            }
            if (size is { } fontSize)
                properties.SetAttributeValue("sz",
                    fontSize.ToString(CultureInfo.InvariantCulture));
            if (bold is { } boldValueValue)
                properties.SetAttributeValue("b", boldValueValue ? "1" : "0");
            if (italic is { } italicValueValue)
                properties.SetAttributeValue("i", italicValueValue ? "1" : "0");
        }
    }

    private static void RemoveFillElements(XElement parent)
    {
        var names = new HashSet<string>(StringComparer.Ordinal)
        {
            "noFill", "solidFill", "gradFill", "blipFill", "pattFill",
            "grpFill",
        };
        foreach (var child in parent.Elements()
                     .Where(child => names.Contains(child.Name.LocalName)).ToList())
            child.Remove();
    }

    private static void InsertBefore(
        XElement parent, XElement child, IReadOnlyCollection<string> followers)
    {
        var anchor = parent.Elements().FirstOrDefault(element =>
            followers.Contains(element.Name.LocalName));
        if (anchor is not null)
            anchor.AddBeforeSelf(child);
        else
            parent.Add(child);
    }

    private static void SetBackground(XDocument document, string hex)
    {
        var common = document.Root?.Element(PptFileOoxml.P + "cSld")
                     ?? throw new NativeRuntimeException("幻灯片缺少 cSld");
        common.Element(PptFileOoxml.P + "bg")?.Remove();
        common.AddFirst(new XElement(PptFileOoxml.P + "bg",
            new XElement(PptFileOoxml.P + "bgPr",
                PptFileOoxml.SolidFill(hex),
                new XElement(PptFileOoxml.A + "effectLst"))));
    }

    private static void SetNotes(SlidePart slidePart, string text)
    {
        var notesPart = slidePart.NotesSlidePart;
        if (notesPart?.NotesSlide is null)
        {
            PptFileOoxml.AddNotes(slidePart, text);
            notesPart = slidePart.NotesSlidePart;
            if (notesPart?.NotesSlide is null)
                throw new NativeRuntimeException("无法创建备注页");
        }
        var document = XDocument.Parse(notesPart.NotesSlide.OuterXml);
        var tree = PptFileOoxml.SlideTree(document);
        var body = PptFileOoxml.ShapeElements(tree).FirstOrDefault(shape =>
        {
            var info = PptFileOoxml.PlaceholderInfo(shape);
            return info?.Type is { } type
                   && string.Equals(type, "body", StringComparison.OrdinalIgnoreCase);
        });
        if (body is null)
        {
            var id = PptFileContent.MaxShapeId(document) + 1;
            body = new XElement(PptFileOoxml.P + "sp",
                new XElement(PptFileOoxml.P + "nvSpPr",
                    new XElement(PptFileOoxml.P + "cNvPr",
                        new XAttribute("id", id),
                        new XAttribute("name", $"Notes Placeholder {id}")),
                    new XElement(PptFileOoxml.P + "cNvSpPr"),
                    new XElement(PptFileOoxml.P + "nvPr",
                        new XElement(PptFileOoxml.P + "ph",
                            new XAttribute("type", "body"),
                            new XAttribute("idx", "3")))),
                new XElement(PptFileOoxml.P + "spPr"),
                new XElement(PptFileOoxml.P + "txBody",
                    new XElement(PptFileOoxml.A + "bodyPr"),
                    new XElement(PptFileOoxml.A + "lstStyle"),
                    new XElement(PptFileOoxml.A + "p")));
            tree.Add(body);
        }
        var textBody = body.Element(PptFileOoxml.P + "txBody");
        if (textBody is null)
        {
            textBody = new XElement(PptFileOoxml.P + "txBody",
                new XElement(PptFileOoxml.A + "bodyPr"),
                new XElement(PptFileOoxml.A + "lstStyle"));
            body.Add(textBody);
        }
        SetTextBody(textBody, text);
        notesPart.NotesSlide = new P.NotesSlide(
            document.Root!.ToString(SaveOptions.DisableFormatting));
    }
}
