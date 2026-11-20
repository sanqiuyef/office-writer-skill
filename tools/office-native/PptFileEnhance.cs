using System.Globalization;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using XPresentation = DocumentFormat.OpenXml.Presentation;

namespace OfficeNative;

/// <summary>配色、布局审计/重排、占位符清理。</summary>
internal static class PptFileEnhance
{
    private const long LayoutEpsilon = 9144;   // 0.01in
    private const long SafeMargin = 91440;     // 0.10in
    private const long MinShapeGap = 137160;   // 0.15in

    private readonly record struct PptBox(long X, long Y, long Width, long Height);

    // ── 配色方案 ────────────────────────────────────────────────

    public static Dictionary<string, object?> ApplyColorScheme(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        NativeCommon.EnsureFile(filePath, ".pptx");
        var schemeName = NativeCommon.OptString(p, "scheme_name") ?? "professional-blue";
        if (!PptModels.ColorSchemes.TryGetValue(schemeName, out var scheme))
        {
            var names = string.Join(", ", PptModels.ColorSchemes.Keys.OrderBy(
                name => name, StringComparer.Ordinal));
            throw new NativeUsageException($"未知配色方案: {schemeName}；可用: {names}");
        }

        var fullPath = Path.GetFullPath(filePath);
        var work = PptFileOoxml.WorkPath(fullPath);
        try
        {
            File.Copy(fullPath, work, overwrite: true);
            var slidesUpdated = 0;
            using (var document = PptFileOoxml.Open(work, true))
            {
                var presentationPart = PptFileOoxml.MainPart(document);
                foreach (var slidePart in PptFileOoxml.SlideParts(presentationPart))
                {
                    var slideDocument = PptFileOoxml.SlideDocument(slidePart);
                    SetBackground(slideDocument, scheme.BackgroundColor);

                    foreach (var shape in PptFileOoxml.ShapeElements(slideDocument))
                    {
                        if (shape.Name == PptFileOoxml.P + "graphicFrame"
                            && TableOf(shape) is { } table)
                        {
                            ApplyTableScheme(table, scheme);
                            continue;
                        }
                        var textBody = shape.Element(PptFileOoxml.P + "txBody");
                        if (textBody is not null)
                        {
                            var isTitle = PptFileOoxml.PlaceholderInfo(shape) is { } info
                                          && info.Type is "title" or "ctrTitle";
                            foreach (var run in textBody.Descendants(PptFileOoxml.A + "r"))
                            {
                                if (isTitle)
                                    ApplyRunStyle(run, scheme.TitleColor,
                                        scheme.TitleFont, scheme.TitleSize);
                                else
                                    ApplyRunStyle(run, scheme.BodyColor,
                                        scheme.BodyFont, scheme.BodySize);
                            }
                        }
                        else if (shape.Name == PptFileOoxml.P + "sp")
                        {
                            SetShapeSolidFill(shape, scheme.ShapeFillColor);
                        }
                    }
                    PptFileOoxml.SaveSlide(slidePart, slideDocument);
                    slidesUpdated++;
                }
                document.Save();
            }
            PptFileOoxml.ReplaceWithBackup(work, fullPath);
            return new Dictionary<string, object?>
            {
                ["success"] = true,
                ["file_path"] = fullPath,
                ["scheme"] = schemeName,
                ["slides_updated"] = slidesUpdated,
                ["message"] = $"已对 {slidesUpdated} 张幻灯片应用「{scheme.DisplayName}」配色方案",
            };
        }
        finally
        {
            PptFileOoxml.EnsureCleanup(work);
        }
    }

    private static XElement? TableOf(XElement graphicFrame) =>
        graphicFrame.Element(PptFileOoxml.A + "graphic")
            ?.Element(PptFileOoxml.A + "graphicData")
            ?.Element(PptFileOoxml.A + "tbl");

    private static void ApplyTableScheme(XElement table, PptColorScheme scheme)
    {
        var rowIndex = 0;
        foreach (var row in table.Elements(PptFileOoxml.A + "tr"))
        {
            foreach (var cell in row.Elements(PptFileOoxml.A + "tc"))
            {
                var textBody = cell.Element(PptFileOoxml.A + "txBody");
                if (textBody is not null)
                {
                    foreach (var run in textBody.Descendants(PptFileOoxml.A + "r"))
                    {
                        if (rowIndex == 0)
                            ApplyRunStyle(run, "FFFFFF", null, null);
                        else
                            ApplyRunStyle(run, scheme.BodyColor,
                                scheme.BodyFont, scheme.BodySize);
                    }
                }
                if (rowIndex == 0)
                {
                    var cellProperties = cell.Element(PptFileOoxml.A + "tcPr");
                    if (cellProperties is null)
                    {
                        cellProperties = new XElement(PptFileOoxml.A + "tcPr");
                        cell.Add(cellProperties);
                    }
                    RemoveFill(cellProperties);
                    InsertBeforeFirst(cellProperties,
                        PptFileOoxml.SolidFill(scheme.TableHeaderColor),
                        "headers", "extLst");
                }
            }
            rowIndex++;
        }
    }

    private static void SetBackground(XDocument document, string color)
    {
        var common = document.Root?.Element(PptFileOoxml.P + "cSld");
        if (common is null)
            return;
        common.Element(PptFileOoxml.P + "bg")?.Remove();
        common.AddFirst(new XElement(PptFileOoxml.P + "bg",
            new XElement(PptFileOoxml.P + "bgPr",
                PptFileOoxml.SolidFill(color),
                new XElement(PptFileOoxml.A + "effectLst"))));
    }

    private static void ApplyRunStyle(XElement run, string color, string? font, int? size)
    {
        var properties = run.Element(PptFileOoxml.A + "rPr");
        if (properties is null)
        {
            properties = new XElement(PptFileOoxml.A + "rPr");
            run.AddFirst(properties);
        }
        RemoveFill(properties);
        properties.Element(PptFileOoxml.A + "latin")?.Remove();
        properties.Element(PptFileOoxml.A + "ea")?.Remove();
        InsertBeforeFirst(properties, PptFileOoxml.SolidFill(color),
            "effectLst", "effectDag", "highlight", "uLnTx", "uLn", "uFillTx",
            "uFill", "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver",
            "rtl", "extLst");
        if (!string.IsNullOrWhiteSpace(font))
        {
            InsertBeforeFirst(properties,
                new XElement(PptFileOoxml.A + "latin",
                    new XAttribute("typeface", font)),
                "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst");
            InsertBeforeFirst(properties,
                new XElement(PptFileOoxml.A + "ea",
                    new XAttribute("typeface", font)),
                "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst");
        }
        if (size is not null)
            properties.SetAttributeValue("sz", size.Value * 100);
    }

    private static void SetShapeSolidFill(XElement shape, string color)
    {
        var shapeProperties = shape.Element(PptFileOoxml.P + "spPr");
        if (shapeProperties is null)
            return;
        RemoveFill(shapeProperties);
        InsertBeforeFirst(shapeProperties, PptFileOoxml.SolidFill(color),
            "ln", "effectLst", "effectDag", "scene3d", "sp3d", "extLst");
    }

    private static void InsertBeforeFirst(XElement parent, XElement child,
        params string[] localNames)
    {
        var reference = parent.Elements().FirstOrDefault(element =>
            localNames.Contains(element.Name.LocalName));
        if (reference is not null)
            reference.AddBeforeSelf(child);
        else
            parent.Add(child);
    }

    private static void RemoveFill(XElement properties)
    {
        foreach (var name in new[]
                 {
                     "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill",
                 })
            properties.Element(PptFileOoxml.A + name)?.Remove();
    }

    // ── 布局审计 / 重排 / 间距优化 ──────────────────────────────

    public static Dictionary<string, object?> OptimizeSlideLayout(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        NativeCommon.EnsureFile(filePath, ".pptx");
        var reflow = PptModels.OptBool(p, "reflow", false);
        var fixSpacing = PptModels.OptBool(p, "fix_spacing", false);
        if (reflow)
            return Reflow(filePath);
        if (fixSpacing)
            return FixSpacing(filePath);
        return Audit(filePath);
    }

    private static Dictionary<string, object?> Audit(string filePath)
    {
        var fullPath = Path.GetFullPath(filePath);
        using var document = PptFileOoxml.Open(fullPath, false);
        var presentationPart = PptFileOoxml.MainPart(document);
        var presentation = presentationPart.Presentation;
        var slideWidth = presentation.SlideSize?.Cx?.Value ?? 9144000;
        var slideHeight = presentation.SlideSize?.Cy?.Value ?? 6858000;
        var warnings = new List<Dictionary<string, object?>>();
        var shapeCount = 0;
        var slidesChecked = 0;

        var slideNumber = 0;
        foreach (var slidePart in PptFileOoxml.SlideParts(presentationPart))
        {
            slideNumber++;
            slidesChecked++;
            var slideDocument = PptFileOoxml.SlideDocument(slidePart);
            var context = PptFileOoxml.ContextFor(slidePart);
            var shapes = PptFileOoxml.ShapeElements(slideDocument).ToList();
            var boxes = new List<(string Label, PptBox Box, int Id)>();
            for (var index = 0; index < shapes.Count; index++)
            {
                var shape = shapes[index];
                var box = PptFileOoxml.GetShapeBox(shape, context);
                if (box is null)
                    continue;
                shapeCount++;
                var current = new PptBox(box.Value.X, box.Value.Y,
                    box.Value.Width, box.Value.Height);
                var label = ShapeLabel(shape, index + 1);
                var shapeId = PptFileOoxml.ShapeId(shape);
                var overflows = new List<string>();
                if (current.X < -LayoutEpsilon)
                    overflows.Add($"左边界超出 {InchesText(-current.X)}in");
                if (current.Y < -LayoutEpsilon)
                    overflows.Add($"上边界超出 {InchesText(-current.Y)}in");
                if (current.X + current.Width > slideWidth + LayoutEpsilon)
                    overflows.Add(
                        $"右边界超出 {InchesText(current.X + current.Width - slideWidth)}in");
                if (current.Y + current.Height > slideHeight + LayoutEpsilon)
                    overflows.Add(
                        $"下边界超出 {InchesText(current.Y + current.Height - slideHeight)}in");
                if (overflows.Count > 0)
                {
                    warnings.Add(new Dictionary<string, object?>
                    {
                        ["type"] = "out_of_bounds",
                        ["slide"] = slideNumber,
                        ["shape"] = label,
                        ["shape_id"] = shapeId,
                        ["detail"] = string.Join("；", overflows),
                    });
                }
                boxes.Add((label, current, shapeId));
            }

            for (var first = 0; first < boxes.Count; first++)
            {
                for (var second = first + 1; second < boxes.Count; second++)
                {
                    var a = boxes[first];
                    var b = boxes[second];
                    var overlapX = Math.Min(a.Box.X + a.Box.Width, b.Box.X + b.Box.Width)
                                   - Math.Max(a.Box.X, b.Box.X);
                    var overlapY = Math.Min(a.Box.Y + a.Box.Height, b.Box.Y + b.Box.Height)
                                   - Math.Max(a.Box.Y, b.Box.Y);
                    if (overlapX <= LayoutEpsilon || overlapY <= LayoutEpsilon)
                        continue;
                    warnings.Add(new Dictionary<string, object?>
                    {
                        ["type"] = "overlap",
                        ["slide"] = slideNumber,
                        ["shapes"] = new List<string> { a.Label, b.Label },
                        ["shape_ids"] = new List<int> { a.Id, b.Id },
                        ["detail"] = $"重叠约 {InchesText(overlapX)}×"
                                     + $"{InchesText(overlapY)}in",
                    });
                }
            }
        }

        var message = warnings.Count > 0
            ? $"布局检查完成：检查 {slidesChecked} 张幻灯片、{shapeCount} 个形状，"
              + $"发现 {warnings.Count} 条警告（未自动重排；自动重排用 --reflow，"
              + "纵向间距优化用 --fix-spacing）"
            : $"布局检查完成：检查 {slidesChecked} 张幻灯片、{shapeCount} 个形状，"
              + "未发现越界或重叠";
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["file_path"] = fullPath,
            ["slides_checked"] = slidesChecked,
            ["shape_count"] = shapeCount,
            ["warning_count"] = warnings.Count,
            ["warnings"] = warnings,
            ["message"] = message,
        };
    }

    private static Dictionary<string, object?> Reflow(string filePath)
    {
        var fullPath = Path.GetFullPath(filePath);
        var work = PptFileOoxml.WorkPath(fullPath);
        var reflowEntries = new List<Dictionary<string, object?>>();
        var unplaceable = new List<Dictionary<string, object?>>();
        var slidesChecked = 0;
        var slidesReflowed = 0;
        var shapeCount = 0;
        var changedCount = 0;
        try
        {
            File.Copy(fullPath, work, overwrite: true);
            using (var document = PptFileOoxml.Open(work, true))
            {
                var presentationPart = PptFileOoxml.MainPart(document);
                var presentation = presentationPart.Presentation;
                var slideWidth = presentation.SlideSize?.Cx?.Value ?? 9144000;
                var slideHeight = presentation.SlideSize?.Cy?.Value ?? 6858000;

                var slideNumber = 0;
                foreach (var slidePart in PptFileOoxml.SlideParts(presentationPart))
                {
                    slideNumber++;
                    slidesChecked++;
                    var slideDocument = PptFileOoxml.SlideDocument(slidePart);
                    var context = PptFileOoxml.ContextFor(slidePart);
                    var shapes = PptFileOoxml.ShapeElements(slideDocument).ToList();
                    var entries = new List<(XElement Shape, int Index, PptBox Box)>();
                    for (var index = 0; index < shapes.Count; index++)
                    {
                        var box = PptFileOoxml.GetShapeBox(shapes[index], context);
                        if (box is null)
                            continue;
                        shapeCount++;
                        entries.Add((shapes[index], index + 1,
                            new PptBox(box.Value.X, box.Value.Y,
                                box.Value.Width, box.Value.Height)));
                    }
                    if (entries.Count == 0)
                        continue;

                    var fitted = new Dictionary<XElement, PptBox>();
                    var reasons = new Dictionary<XElement, string>();
                    foreach (var entry in entries)
                    {
                        if (IsOutOfBounds(entry.Box, slideWidth, slideHeight))
                        {
                            fitted[entry.Shape] = FitBox(entry.Box, slideWidth, slideHeight);
                            reasons[entry.Shape] = "out_of_bounds";
                        }
                        else
                        {
                            fitted[entry.Shape] = entry.Box;
                        }
                    }

                    var ordered = entries
                        .OrderBy(entry => fitted[entry.Shape].Y)
                        .ThenBy(entry => fitted[entry.Shape].X)
                        .ThenBy(entry => PptFileOoxml.ShapeId(entry.Shape))
                        .ToList();
                    var placed = new List<PptBox>();
                    var final = new Dictionary<XElement, PptBox>();
                    foreach (var entry in ordered)
                    {
                        var current = fitted[entry.Shape];
                        if (placed.Any(other => BoxesOverlap(current, other)))
                        {
                            var reason = reasons.TryGetValue(entry.Shape, out var existing)
                                         && existing == "out_of_bounds"
                                ? "out_of_bounds+overlap"
                                : "overlap";
                            var candidate = FindFreeBox(current, placed,
                                slideWidth, slideHeight);
                            if (candidate is null)
                            {
                                unplaceable.Add(new Dictionary<string, object?>
                                {
                                    ["type"] = "unplaceable",
                                    ["slide"] = slideNumber,
                                    ["shape"] = ShapeLabel(entry.Shape, entry.Index),
                                    ["shape_id"] = PptFileOoxml.ShapeId(entry.Shape),
                                    ["detail"] = "无法在保留 0.1in 边距和 0.15in 间距的前提下"
                                                 + "放置，保持原状",
                                });
                                reasons[entry.Shape] = "unplaceable";
                                final[entry.Shape] = entry.Box;
                            }
                            else
                            {
                                reasons[entry.Shape] = reason;
                                final[entry.Shape] = candidate.Value;
                            }
                        }
                        else
                        {
                            final[entry.Shape] = current;
                        }
                        placed.Add(final[entry.Shape]);
                    }

                    var changedOnSlide = 0;
                    foreach (var entry in entries)
                    {
                        var newBox = final[entry.Shape];
                        if (newBox == entry.Box)
                        {
                            if (reasons.TryGetValue(entry.Shape, out var reason)
                                && reason == "unplaceable")
                            {
                                reflowEntries.Add(new Dictionary<string, object?>
                                {
                                    ["shape_id"] = PptFileOoxml.ShapeId(entry.Shape),
                                    ["name"] = PptFileOoxml.ShapeName(entry.Shape),
                                    ["slide"] = slideNumber,
                                    ["action"] = "unchanged",
                                    ["reason"] = "unplaceable",
                                    ["from"] = BoxAsInches(entry.Box),
                                    ["to"] = BoxAsInches(entry.Box),
                                });
                            }
                            continue;
                        }
                        PptFileOoxml.SetShapeBox(entry.Shape,
                            newBox.X, newBox.Y, newBox.Width, newBox.Height);
                        var moved = newBox.X != entry.Box.X || newBox.Y != entry.Box.Y;
                        var resized = newBox.Width != entry.Box.Width
                                      || newBox.Height != entry.Box.Height;
                        var action = moved && resized
                            ? "move_resize"
                            : resized ? "resize" : "move";
                        reflowEntries.Add(new Dictionary<string, object?>
                        {
                            ["shape_id"] = PptFileOoxml.ShapeId(entry.Shape),
                            ["name"] = PptFileOoxml.ShapeName(entry.Shape),
                            ["slide"] = slideNumber,
                            ["action"] = action,
                            ["reason"] = reasons.GetValueOrDefault(entry.Shape, "overlap"),
                            ["from"] = BoxAsInches(entry.Box),
                            ["to"] = BoxAsInches(newBox),
                        });
                        changedOnSlide++;
                        changedCount++;
                    }
                    if (changedOnSlide > 0)
                    {
                        PptFileOoxml.SaveSlide(slidePart, slideDocument);
                        slidesReflowed++;
                    }
                }
                if (changedCount > 0)
                    document.Save();
            }

            if (changedCount > 0)
                PptFileOoxml.ReplaceWithBackup(work, fullPath);
            else
                PptFileOoxml.EnsureCleanup(work);

            var audit = Audit(fullPath);
            var remaining = audit.GetValueOrDefault("warnings") as List<Dictionary<string, object?>>
                            ?? [];
            var message = changedCount > 0
                ? $"已重排 {changedCount} 个形状（{slidesReflowed} 张幻灯片）"
                : "未发现可安全重排的越界或重叠形状";
            if (remaining.Count > 0)
                message += $"；仍有 {remaining.Count} 条布局警告";
            return new Dictionary<string, object?>
            {
                ["success"] = true,
                ["file_path"] = fullPath,
                ["slides_checked"] = slidesChecked,
                ["slides_reflowed"] = slidesReflowed,
                ["shape_count"] = shapeCount,
                ["reflow_count"] = changedCount,
                ["reflow"] = reflowEntries,
                ["warnings"] = unplaceable,
                ["remaining_warnings"] = remaining,
                ["remaining_warning_count"] = remaining.Count,
                ["message"] = message,
            };
        }
        finally
        {
            PptFileOoxml.EnsureCleanup(work);
        }
    }

    private static Dictionary<string, object?> FixSpacing(string filePath)
    {
        var fullPath = Path.GetFullPath(filePath);
        var work = PptFileOoxml.WorkPath(fullPath);
        try
        {
            File.Copy(fullPath, work, overwrite: true);
            var slidesOptimized = 0;
            using (var document = PptFileOoxml.Open(work, true))
            {
                var presentationPart = PptFileOoxml.MainPart(document);
                var slideHeight = presentationPart.Presentation.SlideSize?.Cy?.Value
                                  ?? 6858000;
                const long topMargin = 731520;    // 0.8in
                const long bottomMargin = 457200; // 0.5in
                foreach (var slidePart in PptFileOoxml.SlideParts(presentationPart))
                {
                    var slideDocument = PptFileOoxml.SlideDocument(slidePart);
                    var context = PptFileOoxml.ContextFor(slidePart);
                    var shapes = PptFileOoxml.ShapeElements(slideDocument)
                        .Where(shape => shape.Name == PptFileOoxml.P + "sp")
                        .Select(shape =>
                        {
                            var textBody = shape.Element(PptFileOoxml.P + "txBody");
                            var box = PptFileOoxml.GetShapeBox(shape, context);
                            return (Shape: shape,
                                Text: PptFileOoxml.TxBodyText(textBody),
                                Box: box);
                        })
                        .Where(entry => entry.Box is not null
                                        && entry.Text.Trim().Length > 0)
                        .OrderBy(entry => entry.Box!.Value.Y)
                        .ToList();
                    if (shapes.Count == 0)
                        continue;
                    if (shapes.Count <= 1)
                    {
                        slidesOptimized++;
                        continue;
                    }
                    var usableHeight = slideHeight - topMargin - bottomMargin;
                    var segmentHeight = usableHeight / (double)shapes.Count;
                    for (var index = 0; index < shapes.Count; index++)
                    {
                        var (shape, text, box) = shapes[index];
                        var newTop = topMargin + (long)Math.Round(segmentHeight * index);
                        var lineCount = Math.Max(text.Split('\n')
                            .Sum(line => line.Split('\v').Length), 1);
                        var newHeight = Math.Max(548640,
                            304800L * lineCount + 101600);
                        PptFileOoxml.SetShapeBox(shape,
                            box!.Value.X, newTop, box.Value.Width, newHeight);
                    }
                    PptFileOoxml.SaveSlide(slidePart, slideDocument);
                    slidesOptimized++;
                }
                document.Save();
            }
            PptFileOoxml.ReplaceWithBackup(work, fullPath);
            return new Dictionary<string, object?>
            {
                ["success"] = true,
                ["file_path"] = fullPath,
                ["slides_optimized"] = slidesOptimized,
                ["message"] = $"已优化 {slidesOptimized} 张幻灯片的布局",
            };
        }
        finally
        {
            PptFileOoxml.EnsureCleanup(work);
        }
    }

    // ── 占位符清理 ──────────────────────────────────────────────

    public static Dictionary<string, object?> CleanUnusedPlaceholders(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        NativeCommon.EnsureFile(filePath, ".pptx");
        var fullPath = Path.GetFullPath(filePath);
        var work = PptFileOoxml.WorkPath(fullPath);
        try
        {
            File.Copy(fullPath, work, overwrite: true);
            var removed = 0;
            using (var document = PptFileOoxml.Open(work, true))
            {
                var presentationPart = PptFileOoxml.MainPart(document);
                foreach (var slidePart in PptFileOoxml.SlideParts(presentationPart))
                {
                    var slideDocument = PptFileOoxml.SlideDocument(slidePart);
                    var toRemove = PptFileOoxml.ShapeElements(slideDocument)
                        .Where(shape =>
                        {
                            if (shape.Name != PptFileOoxml.P + "sp"
                                || !PptFileOoxml.IsPlaceholder(shape))
                                return false;
                            var textBody = shape.Element(PptFileOoxml.P + "txBody");
                            // 页码/日期等域即使没有普通文字也视为有内容，避免误删。
                            return !PptFileOoxml.TxBodyHasField(textBody)
                                   && string.IsNullOrWhiteSpace(
                                       PptFileOoxml.TxBodyText(textBody));
                        })
                        .ToList();
                    foreach (var shape in toRemove)
                    {
                        shape.Remove();
                        removed++;
                    }
                    if (toRemove.Count > 0)
                        PptFileOoxml.SaveSlide(slidePart, slideDocument);
                }
                document.Save();
            }
            PptFileOoxml.ReplaceWithBackup(work, fullPath);
            return new Dictionary<string, object?>
            {
                ["success"] = true,
                ["file_path"] = fullPath,
                ["placeholders_removed"] = removed,
                ["message"] = $"已清理 {removed} 个未使用的占位符",
            };
        }
        finally
        {
            PptFileOoxml.EnsureCleanup(work);
        }
    }

    // ── 布局几何算法（与 ppt_writer.py 对齐）────────────────────

    private static bool IsOutOfBounds(PptBox box, long slideWidth, long slideHeight) =>
        box.X < -LayoutEpsilon || box.Y < -LayoutEpsilon
        || box.X + box.Width > slideWidth + LayoutEpsilon
        || box.Y + box.Height > slideHeight + LayoutEpsilon;

    private static PptBox ClampBox(PptBox box, long slideWidth, long slideHeight)
    {
        var x = Math.Max(SafeMargin,
            Math.Min(box.X, slideWidth - SafeMargin - box.Width));
        var y = Math.Max(SafeMargin,
            Math.Min(box.Y, slideHeight - SafeMargin - box.Height));
        return new PptBox(x, y, box.Width, box.Height);
    }

    private static bool BoxInside(PptBox box, long slideWidth, long slideHeight) =>
        box.X >= SafeMargin - LayoutEpsilon
        && box.Y >= SafeMargin - LayoutEpsilon
        && box.X + box.Width <= slideWidth - SafeMargin + LayoutEpsilon
        && box.Y + box.Height <= slideHeight - SafeMargin + LayoutEpsilon;

    private static PptBox FitBox(PptBox box, long slideWidth, long slideHeight)
    {
        var availableWidth = Math.Max(1, slideWidth - 2 * SafeMargin);
        var availableHeight = Math.Max(1, slideHeight - 2 * SafeMargin);
        var scale = Math.Min(1.0, Math.Min(availableWidth / (double)box.Width,
            availableHeight / (double)box.Height));
        var width = box.Width;
        var height = box.Height;
        if (scale < 1.0)
        {
            width = Math.Max(1, (long)(box.Width * scale));
            height = Math.Max(1, (long)(box.Height * scale));
        }
        return ClampBox(new PptBox(box.X, box.Y, width, height),
            slideWidth, slideHeight);
    }

    private static bool BoxesOverlap(PptBox first, PptBox second)
    {
        var overlapX = Math.Min(first.X + first.Width, second.X + second.Width)
                       - Math.Max(first.X, second.X);
        var overlapY = Math.Min(first.Y + first.Height, second.Y + second.Height)
                       - Math.Max(first.Y, second.Y);
        return overlapX > LayoutEpsilon && overlapY > LayoutEpsilon;
    }

    private static List<PptBox> CandidateBoxes(PptBox box, List<PptBox> placed)
    {
        var candidates = new List<PptBox> { box };
        var overlapping = placed.Where(other => BoxesOverlap(box, other)).ToList();
        if (overlapping.Count > 0)
        {
            var down = overlapping.Max(other => other.Y + other.Height) + MinShapeGap;
            var right = overlapping.Max(other => other.X + other.Width) + MinShapeGap;
            candidates.Add(new PptBox(box.X, down, box.Width, box.Height));
            candidates.Add(new PptBox(right, box.Y, box.Width, box.Height));
            candidates.Add(new PptBox(right, down, box.Width, box.Height));
        }
        foreach (var other in placed)
        {
            candidates.Add(new PptBox(box.X, other.Y + other.Height + MinShapeGap,
                box.Width, box.Height));
            candidates.Add(new PptBox(other.X + other.Width + MinShapeGap, box.Y,
                box.Width, box.Height));
        }
        foreach (var factor in new[] { 0.9, 0.8, 0.7, 0.6, 0.5 })
        {
            var scaledWidth = Math.Max(1, (long)(box.Width * factor));
            var scaledHeight = Math.Max(1, (long)(box.Height * factor));
            candidates.Add(new PptBox(box.X, box.Y, scaledWidth, scaledHeight));
            foreach (var other in placed)
            {
                candidates.Add(new PptBox(box.X,
                    other.Y + other.Height + MinShapeGap, scaledWidth, scaledHeight));
                candidates.Add(new PptBox(other.X + other.Width + MinShapeGap, box.Y,
                    scaledWidth, scaledHeight));
            }
        }
        return candidates;
    }

    private static PptBox? FindFreeBox(PptBox box, List<PptBox> placed,
        long slideWidth, long slideHeight)
    {
        PptBox? best = null;
        double? bestScore = null;
        var originalArea = box.Width * (double)box.Height;
        foreach (var candidate in CandidateBoxes(box, placed))
        {
            var clamped = ClampBox(candidate, slideWidth, slideHeight);
            if (clamped == box)
                continue;
            if (!BoxInside(clamped, slideWidth, slideHeight))
                continue;
            if (placed.Any(other => BoxesOverlap(clamped, other)))
                continue;
            var distance = Math.Abs(clamped.X - box.X) + Math.Abs(clamped.Y - box.Y);
            var areaLoss = 1.0 - clamped.Width * (double)clamped.Height / originalArea;
            var score = distance + areaLoss * 914400;
            if (bestScore is null || score < bestScore)
            {
                best = clamped;
                bestScore = score;
            }
        }
        return best;
    }

    private static Dictionary<string, object?> BoxAsInches(PptBox box) => new()
    {
        ["left"] = Math.Round(PptFileOoxml.ToInches(box.X), 3),
        ["top"] = Math.Round(PptFileOoxml.ToInches(box.Y), 3),
        ["width"] = Math.Round(PptFileOoxml.ToInches(box.Width), 3),
        ["height"] = Math.Round(PptFileOoxml.ToInches(box.Height), 3),
    };

    private static string ShapeLabel(XElement shape, int index)
    {
        var name = PptFileOoxml.ShapeName(shape);
        return name.Length > 0 ? $"#{index} {name}" : $"#{index}";
    }

    private static string InchesText(long emu) =>
        PptFileOoxml.ToInches(emu).ToString("0.00", CultureInfo.InvariantCulture);
}
