using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XDrawing = DocumentFormat.OpenXml.Drawing;
using XPresentation = DocumentFormat.OpenXml.Presentation;

namespace OfficeNative;

/// <summary>PowerPoint 文件级底层：包/版式/占位符继承/几何读取/原子保存。</summary>
internal static class PptFileOoxml
{
    public static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    public static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    public static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    public static readonly XNamespace P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";

    public const string SectionExtensionUri = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
    public const string DefaultSectionName = "默认节";

    public const long EmuPerInch = 914400;
    public const long EmuPerPoint = 12700;

    public static long Inches(double value) => (long)Math.Round(value * EmuPerInch);
    public static long Points(double value) => (long)Math.Round(value * EmuPerPoint);
    public static double ToInches(long emu) => emu / (double)EmuPerInch;

    public static string NormalizeHex(string? color)
    {
        var text = (color ?? "").Trim().TrimStart('#');
        return text.Length == 6 ? text.ToUpperInvariant() : "000000";
    }

    public static XElement El(XNamespace ns, string name, params object?[] content) =>
        new(ns + name, content);

    public static XElement SolidFill(string hex) =>
        new(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", NormalizeHex(hex))));

    // ── 包生命周期 ───────────────────────────────────────────────

    public static string WorkPath(string targetPath)
    {
        var full = Path.GetFullPath(targetPath);
        var dir = Path.GetDirectoryName(full)
                  ?? throw new NativeRuntimeException($"无法解析输出目录: {targetPath}");
        Directory.CreateDirectory(dir);
        return Path.Combine(dir,
            $".{Path.GetFileNameWithoutExtension(full)}-{Guid.NewGuid():N}.pptx");
    }

    /// <summary>写回目标；目标已存在且非空时先备份。返回备份路径。</summary>
    public static string? ReplaceWithBackup(string workPath, string targetPath)
    {
        var full = Path.GetFullPath(targetPath);
        string? backup = null;
        try
        {
            if (File.Exists(full))
            {
                if (new FileInfo(full).Length > 0)
                    backup = NativeCommon.Backup(full);
                try
                {
                    File.Replace(workPath, full, null, ignoreMetadataErrors: true);
                }
                catch (IOException)
                {
                    File.Copy(workPath, full, overwrite: true);
                    File.Delete(workPath);
                }
                catch (PlatformNotSupportedException)
                {
                    File.Copy(workPath, full, overwrite: true);
                    File.Delete(workPath);
                }
            }
            else
            {
                File.Move(workPath, full);
            }
        }
        finally
        {
            if (File.Exists(workPath))
            {
                try { File.Delete(workPath); } catch { /* temp cleanup */ }
            }
        }
        return backup;
    }

    public static void EnsureCleanup(string workPath)
    {
        if (File.Exists(workPath))
        {
            try { File.Delete(workPath); } catch { /* temp cleanup */ }
        }
    }

    public static PresentationDocument Open(string path, bool editable = true) =>
        PresentationDocument.Open(path, editable);

    public static PresentationPart MainPart(PresentationDocument document) =>
        document.PresentationPart
        ?? throw new NativeRuntimeException("演示文稿缺少 presentation part");

    /// <summary>.potx 工作副本：把 template.main 内容类型改写为 presentation.main，便于编辑后另存 .pptx。</summary>
    public static void NormalizePotxContentType(string path)
    {
        using var archive = System.IO.Compression.ZipFile.Open(
            path, System.IO.Compression.ZipArchiveMode.Update);
        var entry = archive.GetEntry("[Content_Types].xml");
        if (entry is null)
            return;
        string text;
        using (var reader = new StreamReader(entry.Open()))
            text = reader.ReadToEnd();
        var replaced = text
            .Replace(
                "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
                "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml")
            .Replace(
                "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml",
                "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml");
        if (replaced == text)
            return;
        entry.Delete();
        var created = archive.CreateEntry("[Content_Types].xml");
        using var writer = new StreamWriter(created.Open(), new System.Text.UTF8Encoding(false));
        writer.Write(replaced);
    }

    // ── 幻灯片遍历 ───────────────────────────────────────────────

    public static List<SlidePart> SlideParts(PresentationPart part)
    {
        var list = new List<SlidePart>();
        var slideIds = part.Presentation.SlideIdList;
        if (slideIds is null)
            return list;
        foreach (var slideId in slideIds.Elements<XPresentation.SlideId>())
        {
            var rid = slideId.RelationshipId?.Value;
            if (string.IsNullOrEmpty(rid))
                continue;
            if (part.GetPartById(rid) is SlidePart slidePart)
                list.Add(slidePart);
        }
        return list;
    }

    public static List<SlideLayoutPart> LayoutParts(PresentationPart part)
    {
        var master = part.SlideMasterParts.FirstOrDefault();
        var list = new List<SlideLayoutPart>();
        if (master?.SlideMaster?.SlideLayoutIdList is null)
            return list;
        foreach (var layoutId in master.SlideMaster.SlideLayoutIdList.Elements<XPresentation.SlideLayoutId>())
        {
            var rid = layoutId.RelationshipId?.Value;
            if (string.IsNullOrEmpty(rid))
                continue;
            if (master.GetPartById(rid) is SlideLayoutPart layoutPart)
                list.Add(layoutPart);
        }
        return list;
    }

    public static string LayoutName(SlideLayoutPart part) =>
        part.SlideLayout?.CommonSlideData?.Name?.Value ?? "";

    public static XDocument SlideDocument(SlidePart part) =>
        XDocument.Parse(part.Slide.OuterXml);

    public static void SaveSlide(SlidePart part, XDocument document) =>
        part.Slide = new XPresentation.Slide(
            document.Root!.ToString(SaveOptions.DisableFormatting));

    public static XElement SlideTree(XDocument document) =>
        document.Root?.Element(P + "cSld")?.Element(P + "spTree")
        ?? throw new NativeRuntimeException("幻灯片缺少 spTree");

    public static IEnumerable<XElement> ShapeElements(XDocument document) =>
        ShapeElements(SlideTree(document));

    public static IEnumerable<XElement> ShapeElements(XElement tree) =>
        tree.Elements().Where(element => element.Name == P + "sp"
            || element.Name == P + "pic"
            || element.Name == P + "graphicFrame"
            || element.Name == P + "cxnSp"
            || element.Name == P + "grpSp"
            || element.Name == P + "contentPart");

    public static int SlideCount(PresentationPart part) => SlideParts(part).Count;

    // ── 新建空白演示文稿（默认母版 + 9 种版式）────────────────────

    public static void CreateScratch(string workPath, double widthInches = 10,
        double heightInches = 7.5)
    {
        using var document = PresentationDocument.Create(
            workPath, PresentationDocumentType.Presentation);
        var presentationPart = document.PresentationPart ?? document.AddPresentationPart();

        var masterPart = presentationPart.AddNewPart<SlideMasterPart>();
        var themePart = masterPart.AddNewPart<ThemePart>();
        themePart.Theme = new XDrawing.Theme(ThemeXml);

        var layoutRefs = new List<(uint Id, string RelationshipId)>();
        for (var index = 0; index < PptModels.LayoutNames.Length; index++)
        {
            var layoutPart = masterPart.AddNewPart<SlideLayoutPart>();
            layoutPart.AddPart(masterPart);
            layoutPart.SlideLayout = new XPresentation.SlideLayout(
                BuildLayoutElement(index).ToString(SaveOptions.DisableFormatting));
            layoutRefs.Add((2147483649U + (uint)index,
                masterPart.GetIdOfPart(layoutPart)));
        }

        masterPart.SlideMaster = new XPresentation.SlideMaster(
            BuildMasterElement(layoutRefs).ToString(SaveOptions.DisableFormatting));

        var presentationXml = BuildPresentationElement(
                presentationPart, masterPart, widthInches, heightInches)
            .ToString(SaveOptions.DisableFormatting);
        presentationPart.Presentation = new XPresentation.Presentation(presentationXml);
        EnsureNotesMaster(presentationPart);
        document.Save();
    }

    public static void SetSlideSize(PresentationPart part, double widthInches,
        double heightInches)
    {
        var presentation = part.Presentation;
        var size = presentation.SlideSize ?? new XPresentation.SlideSize();
        size.Cx = (int)Inches(widthInches);
        size.Cy = (int)Inches(heightInches);
        presentation.SlideSize = size;
    }

    private static XElement BuildPresentationElement(
        PresentationPart presentationPart, SlideMasterPart masterPart,
        double widthInches, double heightInches)
    {
        return new XElement(P + "presentation",
            new XAttribute(XNamespace.Xmlns + "a", A),
            new XAttribute(XNamespace.Xmlns + "r", R),
            new XAttribute(XNamespace.Xmlns + "p", P),
            new XAttribute("saveSubsetFonts", "1"),
            new XAttribute("autoCompressPictures", "0"),
            new XElement(P + "sldMasterIdLst",
                new XElement(P + "sldMasterId",
                    new XAttribute("id", "2147483648"),
                    new XAttribute(R + "id", presentationPart.GetIdOfPart(masterPart)))),
            new XElement(P + "sldSz",
                new XAttribute("cx", Inches(widthInches).ToString(CultureInfo.InvariantCulture)),
                new XAttribute("cy", Inches(heightInches).ToString(CultureInfo.InvariantCulture))),
            new XElement(P + "notesSz",
                new XAttribute("cx", "6858000"),
                new XAttribute("cy", "9144000")),
            DefaultTextStyle());
    }

    private static XElement DefaultTextStyle()
    {
        var levels = new List<XElement>();
        for (var level = 1; level <= 9; level++)
        {
            levels.Add(new XElement(A + $"lvl{level}pPr",
                new XAttribute("marL", (level - 1) * 457200),
                new XAttribute("algn", "l"),
                new XElement(A + "defRPr",
                    new XAttribute("sz", "1800"),
                    new XElement(A + "solidFill",
                        new XElement(A + "schemeClr", new XAttribute("val", "tx1"))),
                    new XElement(A + "latin", new XAttribute("typeface", "+mn-lt")),
                    new XElement(A + "ea", new XAttribute("typeface", "+mn-ea")),
                    new XElement(A + "cs", new XAttribute("typeface", "+mn-cs")))));
        }
        return new XElement(P + "defaultTextStyle",
            new XElement(A + "defPPr",
                new XElement(A + "defRPr", new XAttribute("lang", "en-US"))),
            levels);
    }

    private static XElement GroupProperties() =>
        new(P + "nvGrpSpPr",
            new XElement(P + "cNvPr", new XAttribute("id", "1"), new XAttribute("name", "")),
            new XElement(P + "cNvGrpSpPr"),
            new XElement(P + "nvPr"));

    private static XElement GroupTransform() =>
        new(P + "grpSpPr",
            new XElement(A + "xfrm",
                new XElement(A + "off", new XAttribute("x", "0"), new XAttribute("y", "0")),
                new XElement(A + "ext", new XAttribute("cx", "0"), new XAttribute("cy", "0")),
                new XElement(A + "chOff", new XAttribute("x", "0"), new XAttribute("y", "0")),
                new XElement(A + "chExt", new XAttribute("cx", "0"), new XAttribute("cy", "0"))));

    private static XElement PlaceholderShape(uint id, string name, string type, int? idx,
        double x, double y, double width, double height, string? size = null)
    {
        var ph = new XElement(P + "ph", new XAttribute("type", type));
        if (size is not null)
            ph.SetAttributeValue("sz", size);
        if (idx is not null)
            ph.SetAttributeValue("idx", idx.Value);
        return new XElement(P + "sp",
            new XElement(P + "nvSpPr",
                new XElement(P + "cNvPr", new XAttribute("id", id), new XAttribute("name", name)),
                new XElement(P + "cNvSpPr",
                    new XElement(A + "spLocks", new XAttribute("noGrp", "1"))),
                new XElement(P + "nvPr", ph)),
            new XElement(P + "spPr",
                XfrmElement(Inches(x), Inches(y), Inches(width), Inches(height)),
                new XElement(A + "prstGeom", new XAttribute("prst", "rect"),
                    new XElement(A + "avLst"))),
            new XElement(P + "txBody",
                new XElement(A + "bodyPr"),
                new XElement(A + "lstStyle"),
                new XElement(A + "p")));
    }

    private static XElement XfrmElement(long x, long y, long width, long height) =>
        new(A + "xfrm",
            new XElement(A + "off",
                new XAttribute("x", x), new XAttribute("y", y)),
            new XElement(A + "ext",
                new XAttribute("cx", width), new XAttribute("cy", height)));

    private static XElement BuildMasterElement(
        IEnumerable<(uint Id, string RelationshipId)> layouts)
    {
        var layoutIds = layouts.Select(layout => new XElement(P + "sldLayoutId",
            new XAttribute("id", layout.Id),
            new XAttribute(R + "id", layout.RelationshipId)));
        return new XElement(P + "sldMaster",
            new XAttribute(XNamespace.Xmlns + "a", A),
            new XAttribute(XNamespace.Xmlns + "r", R),
            new XAttribute(XNamespace.Xmlns + "p", P),
            new XElement(P + "cSld",
                new XElement(P + "bg",
                    new XElement(P + "bgRef", new XAttribute("idx", "1001"),
                        new XElement(A + "schemeClr", new XAttribute("val", "bg1")))),
                new XElement(P + "spTree",
                    GroupProperties(),
                    GroupTransform(),
                    PlaceholderShape(2, "Title Placeholder 1", "title", null,
                        0.5, 0.3, 9.0, 1.25),
                    PlaceholderShape(3, "Text Placeholder 2", "body", 1,
                        0.5, 1.75, 9.0, 4.95))),
            new XElement(P + "clrMap",
                new XAttribute("bg1", "lt1"), new XAttribute("tx1", "dk1"),
                new XAttribute("bg2", "lt2"), new XAttribute("tx2", "dk2"),
                new XAttribute("accent1", "accent1"), new XAttribute("accent2", "accent2"),
                new XAttribute("accent3", "accent3"), new XAttribute("accent4", "accent4"),
                new XAttribute("accent5", "accent5"), new XAttribute("accent6", "accent6"),
                new XAttribute("hlink", "hlink"), new XAttribute("folHlink", "folHlink")),
            new XElement(P + "sldLayoutIdLst", layoutIds),
            new XElement(P + "txStyles",
                TitleStyle(), BodyStyle(), OtherStyle()));
    }

    private static XElement TitleStyle() =>
        new(P + "titleStyle",
            new XElement(A + "lvl1pPr", new XAttribute("algn", "ctr"),
                new XElement(A + "defRPr", new XAttribute("sz", "4400"),
                    new XElement(A + "solidFill",
                        new XElement(A + "schemeClr", new XAttribute("val", "tx1"))),
                    new XElement(A + "latin", new XAttribute("typeface", "+mj-lt")),
                    new XElement(A + "ea", new XAttribute("typeface", "+mj-ea")),
                    new XElement(A + "cs", new XAttribute("typeface", "+mj-cs")))));

    private static XElement BodyStyle()
    {
        var levels = new List<XElement>();
        var sizes = new[] { 3200, 2800, 2400, 2000, 2000 };
        var chars = new[] { "•", "–", "•", "–", "»" };
        for (var i = 0; i < sizes.Length; i++)
        {
            var level = i + 1;
            levels.Add(new XElement(A + $"lvl{level}pPr",
                new XAttribute("marL", level == 1 ? 342900 : 342900 + (level - 1) * 400050),
                new XAttribute("indent", level == 1 ? -342900 : -228600),
                new XAttribute("algn", "l"),
                new XElement(A + "buFont", new XAttribute("typeface", "Arial")),
                new XElement(A + "buChar", new XAttribute("char", chars[i])),
                new XElement(A + "defRPr", new XAttribute("sz", sizes[i]),
                    new XElement(A + "solidFill",
                        new XElement(A + "schemeClr", new XAttribute("val", "tx1"))),
                    new XElement(A + "latin", new XAttribute("typeface", "+mn-lt")),
                    new XElement(A + "ea", new XAttribute("typeface", "+mn-ea")),
                    new XElement(A + "cs", new XAttribute("typeface", "+mn-cs")))));
        }
        return new XElement(P + "bodyStyle", levels);
    }

    private static XElement OtherStyle()
    {
        var levels = new List<XElement>();
        for (var level = 1; level <= 9; level++)
        {
            levels.Add(new XElement(A + $"lvl{level}pPr",
                new XAttribute("marL", (level - 1) * 457200),
                new XAttribute("algn", "l"),
                new XElement(A + "defRPr", new XAttribute("sz", "1800"),
                    new XElement(A + "solidFill",
                        new XElement(A + "schemeClr", new XAttribute("val", "tx1"))),
                    new XElement(A + "latin", new XAttribute("typeface", "+mn-lt")),
                    new XElement(A + "ea", new XAttribute("typeface", "+mn-ea")),
                    new XElement(A + "cs", new XAttribute("typeface", "+mn-cs")))));
        }
        return new XElement(P + "otherStyle", levels);
    }

    private static XElement BuildLayoutElement(int index)
    {
        var name = PptModels.LayoutNames[index];
        var type = index switch
        {
            0 => "title",
            1 => "obj",
            2 => "secHead",
            3 => "twoObj",
            4 => "twoTxTwoObj",
            5 => "titleOnly",
            6 => "blank",
            7 => "objTx",
            _ => "picTx",
        };
        var placeholders = new List<XElement>();
        switch (index)
        {
            case 0:
                // 与 python-pptx 默认模板 Title Slide 一致。
                placeholders.Add(PlaceholderShape(2, "Title 1", "ctrTitle", null,
                    0.75, 2.3299, 8.5, 1.6076));
                placeholders.Add(PlaceholderShape(3, "Subtitle 2", "subTitle", 1,
                    1.5, 4.25, 7.0, 1.9167));
                break;
            case 2:
                // 与 python-pptx 默认模板 Section Header 一致（标题在下、正文在上）。
                placeholders.Add(PlaceholderShape(2, "Title 1", "title", null,
                    0.7899, 4.8194, 8.5, 1.4896));
                placeholders.Add(PlaceholderShape(3, "Text Placeholder 2", "body", 1,
                    0.7899, 3.1788, 8.5, 1.6406));
                break;
            case 3:
                placeholders.Add(PlaceholderShape(2, "Title 1", "title", null,
                    0.5, 0.3, 9.0, 1.25));
                placeholders.Add(PlaceholderShape(3, "Content Placeholder 2", "obj", 1,
                    0.5, 1.75, 4.4167, 4.9497, "half"));
                placeholders.Add(PlaceholderShape(4, "Content Placeholder 3", "obj", 2,
                    5.0833, 1.75, 4.4167, 4.9497, "half"));
                break;
            case 4:
                // 与 python-pptx 默认模板 Comparison 一致（左右各一组标题+内容）。
                placeholders.Add(PlaceholderShape(2, "Title 1", "title", null,
                    0.5, 0.3, 9.0, 1.25));
                placeholders.Add(PlaceholderShape(3, "Text Placeholder 2", "body", 1,
                    0.5, 1.6788, 4.4184, 0.6997, "half"));
                placeholders.Add(PlaceholderShape(4, "Content Placeholder 3", "obj", 2,
                    0.5, 2.3785, 4.4184, 4.3212, "half"));
                placeholders.Add(PlaceholderShape(5, "Text Placeholder 4", "body", 3,
                    5.0799, 1.6788, 4.4201, 0.6997, "half"));
                placeholders.Add(PlaceholderShape(6, "Content Placeholder 5", "obj", 4,
                    5.0799, 2.3785, 4.4201, 4.3212, "half"));
                break;
            case 5:
                placeholders.Add(PlaceholderShape(2, "Title 1", "title", null,
                    0.5, 0.3, 9.0, 1.25));
                break;
            case 6:
                break;
            case 7:
                // 与 python-pptx 默认模板 Content with Caption 一致。
                placeholders.Add(PlaceholderShape(2, "Title 1", "title", null,
                    0.5, 0.2986, 3.2899, 1.2708));
                placeholders.Add(PlaceholderShape(3, "Content Placeholder 2", "obj", 1,
                    3.9097, 0.2986, 5.5903, 6.4010));
                placeholders.Add(PlaceholderShape(4, "Text Placeholder 3", "body", 2,
                    0.5, 1.5694, 3.2899, 5.1302));
                break;
            case 8:
                // 与 python-pptx 默认模板 Picture with Caption 一致。
                placeholders.Add(PlaceholderShape(2, "Title 1", "title", null,
                    1.9601, 5.25, 6.0, 0.6198));
                placeholders.Add(PlaceholderShape(3, "Picture Placeholder 2", "pic", 1,
                    1.9601, 0.6701, 6.0, 4.5));
                placeholders.Add(PlaceholderShape(4, "Text Placeholder 3", "body", 2,
                    1.9601, 5.8698, 6.0, 0.8802));
                break;
            default:
                placeholders.Add(PlaceholderShape(2, "Title 1", "title", null,
                    0.5, 0.3, 9.0, 1.25));
                placeholders.Add(PlaceholderShape(3, "Content Placeholder 2", "obj", 1,
                    0.5, 1.75, 9.0, 4.95));
                break;
        }
        return new XElement(P + "sldLayout",
            new XAttribute(XNamespace.Xmlns + "a", A),
            new XAttribute(XNamespace.Xmlns + "r", R),
            new XAttribute(XNamespace.Xmlns + "p", P),
            new XAttribute("type", type),
            new XAttribute("preserve", "1"),
            new XElement(P + "cSld",
                new XAttribute("name", name),
                new XElement(P + "spTree",
                    GroupProperties(), GroupTransform(), placeholders)),
            new XElement(P + "clrMapOvr",
                new XElement(A + "masterClrMapping")));
    }

    // ── 页脚 / 页码 ──────────────────────────────────────────────

    public const double FooterBoxWidthInches = 8.0;
    public const double FooterBoxHeightInches = 0.35;
    public const double FooterBoxTopInches = 7.0;
    public const double FooterBottomMarginInches = 0.5;
    public const double SlideNumberBoxWidthInches = 1.2;
    public const double SlideNumberRightMarginInches = 0.5;
    public const int FooterFontSizePoints = 12;
    public const string FooterGrayHex = "808080";

    public readonly record struct FooterGeometry(
        double FooterLeft, double FooterWidth, double NumberLeft, double NumberWidth, double Top);

    /// <summary>页脚/页码文本框几何（英寸）；与 ppt_writer._footer_geometry 对齐。</summary>
    public static FooterGeometry FooterGeometryFor(double slideWidthInches,
        double slideHeightInches)
    {
        var top = Math.Min(FooterBoxTopInches,
            Math.Max(0.0, slideHeightInches - FooterBottomMarginInches));
        var footerWidth = Math.Min(FooterBoxWidthInches,
            Math.Max(1.0, slideWidthInches - 1.0));
        var numberWidth = Math.Min(SlideNumberBoxWidthInches,
            Math.Max(0.6, slideWidthInches - 2.0));
        return new FooterGeometry(
            FooterLeft: (slideWidthInches - footerWidth) / 2,
            FooterWidth: footerWidth,
            NumberLeft: Math.Max(0.2,
                slideWidthInches - numberWidth - SlideNumberRightMarginInches),
            NumberWidth: numberWidth,
            Top: top);
    }

    /// <summary>
    /// 在 spTree 中查找指定类型的占位符（footer/sldNum）。版式占位符可能没有 spPr 几何，
    /// 需要向版式/母版回溯，否则拿不到可复制的 xfrm。
    /// </summary>
    public static XElement? FindPlaceholderOnSlide(XElement tree, string type)
    {
        return ShapeElements(tree).FirstOrDefault(shape =>
        {
            var info = PlaceholderInfo(shape);
            return info is not null && info.Value.Type is not null
                   && string.Equals(info.Value.Type, type, StringComparison.OrdinalIgnoreCase);
        });
    }

    /// <summary>
    /// 写页脚文字：优先填充幻灯片上已有的 footer 占位符；没有则从版式克隆，
    /// 再没有则按统一几何新建文本框。不触碰标题/正文等已有占位符。
    /// </summary>
    public static void ApplyFooter(XDocument slide, string text,
        SlideContext context, FooterGeometry geometry, uint newId)
    {
        var tree = SlideTree(slide);
        var placeholder = FindPlaceholderOnSlide(tree, "ftr")
                          ?? ClonePlaceholderFromLayout(slide, "ftr", context, newId);
        if (placeholder is not null)
        {
            SetPlaceholderText(placeholder, text);
            return;
        }
        tree.Add(BuildFooterTextbox(newId, text, geometry.FooterLeft, geometry.FooterWidth,
            geometry.Top, "ctr"));
    }

    /// <summary>写页码：占位符优先；否则新建右下角文本框，内容是 slidenum 域。</summary>
    public static void ApplySlideNumber(XDocument slide,
        int slideNumber, SlideContext context, FooterGeometry geometry, uint newId)
    {
        var tree = SlideTree(slide);
        var placeholder = FindPlaceholderOnSlide(tree, "sldNum")
                          ?? ClonePlaceholderFromLayout(slide, "sldNum", context, newId);
        if (placeholder is not null)
        {
            SetPlaceholderField(placeholder, slideNumber);
            return;
        }
        var box = BuildFooterTextbox(newId, null, geometry.NumberLeft, geometry.NumberWidth,
            geometry.Top, "r");
        var paragraph = box.Element(P + "txBody")!.Elements(A + "p").Last();
        paragraph.Add(BuildSlideNumberField(slideNumber));
        tree.Add(box);
    }

    /// <summary>从版式复制 footer/sldNum 占位符到幻灯片；版式没有则返回 null。</summary>
    private static XElement? ClonePlaceholderFromLayout(XDocument slide,
        string type, SlideContext context, uint newId)
    {
        if (context.Layout is null)
            return null;
        var layoutShape = ShapeElements(context.Layout)
            .FirstOrDefault(shape =>
            {
                var info = PlaceholderInfo(shape);
                return info is not null && string.Equals(info.Value.Type, type,
                    StringComparison.OrdinalIgnoreCase);
            });
        if (layoutShape is null)
            return null;

        var info = PlaceholderInfo(layoutShape)!.Value;
        var tree = SlideTree(slide);
        var nextId = Math.Max(newId, 2);
        var sourceXfrm = layoutShape.Element(P + "spPr")?.Element(A + "xfrm");
        if (sourceXfrm is null && context.Master is not null)
        {
            sourceXfrm = ShapeElements(context.Master)
                .Where(shape => PlaceholderInfo(shape)?.Type is { } masterType
                    && string.Equals(masterType, type, StringComparison.OrdinalIgnoreCase))
                .Select(shape => shape.Element(P + "spPr")?.Element(A + "xfrm"))
                .FirstOrDefault(xfrm => xfrm is not null);
        }
        var shapeProperties = new XElement(P + "spPr");
        if (sourceXfrm is not null)
            shapeProperties.Add(new XElement(sourceXfrm));

        var sourceTextBody = layoutShape.Element(P + "txBody");
        var textBody = new XElement(P + "txBody");
        var sourceBodyProperties = sourceTextBody?.Element(A + "bodyPr");
        textBody.Add(sourceBodyProperties is null
            ? new XElement(A + "bodyPr")
            : new XElement(sourceBodyProperties));
        var sourceListStyle = sourceTextBody?.Element(A + "lstStyle");
        if (sourceListStyle is not null)
            textBody.Add(new XElement(sourceListStyle));
        textBody.Add(new XElement(A + "p"));

        var name = ShapeName(layoutShape);
        var clonePh = new XElement(P + "ph",
            new XAttribute("type", type),
            new XAttribute("idx", info.Idx.ToString(CultureInfo.InvariantCulture)));
        // 版式常带 sz/orient（如 sz="quarter"），一并保留以维持 PowerPoint 的占位符语义。
        var layoutPh = layoutShape.Element(P + "nvSpPr")?.Element(P + "nvPr")
            ?.Element(P + "ph");
        foreach (var attributeName in new[] { "sz", "orient" })
        {
            var value = (string?)layoutPh?.Attribute(attributeName);
            if (!string.IsNullOrWhiteSpace(value))
                clonePh.SetAttributeValue(attributeName, value);
        }
        var clone = new XElement(P + "sp",
            new XElement(P + "nvSpPr",
                new XElement(P + "cNvPr",
                    new XAttribute("id", nextId),
                    new XAttribute("name", string.IsNullOrWhiteSpace(name)
                        ? $"Placeholder {nextId}"
                        : name)),
                new XElement(P + "cNvSpPr",
                    new XElement(A + "spLocks", new XAttribute("noGrp", "1"))),
                new XElement(P + "nvPr", clonePh)),
            shapeProperties,
            textBody);
        tree.Add(clone);
        return clone;
    }

    private static void SetPlaceholderText(XElement shape, string text)
    {
        var textBody = shape.Element(P + "txBody");
        if (textBody is null)
        {
            textBody = new XElement(P + "txBody",
                new XElement(A + "bodyPr"),
                new XElement(A + "lstStyle"));
            shape.Add(textBody);
        }
        foreach (var paragraph in textBody.Elements(A + "p").ToList())
            paragraph.Remove();
        var textParagraph = new XElement(A + "p",
            new XElement(A + "pPr", new XAttribute("algn", "ctr")),
            new XElement(A + "r", new XElement(A + "t", text)));
        textBody.Add(textParagraph);
    }

    private static void SetPlaceholderField(XElement shape, int slideNumber)
    {
        var textBody = shape.Element(P + "txBody");
        if (textBody is null)
        {
            textBody = new XElement(P + "txBody",
                new XElement(A + "bodyPr"),
                new XElement(A + "lstStyle"));
            shape.Add(textBody);
        }
        foreach (var paragraph in textBody.Elements(A + "p").ToList())
            paragraph.Remove();
        textBody.Add(new XElement(A + "p",
            new XElement(A + "pPr", new XAttribute("algn", "r")),
            BuildSlideNumberField(slideNumber)));
    }

    /// <summary>DrawingML 幻灯片编号域（缓存文本为当前页码）。</summary>
    public static XElement BuildSlideNumberField(int slideNumber) =>
        new(A + "fld",
            new XAttribute("id", "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}"),
            new XAttribute("type", "slidenum"),
            new XElement(A + "rPr",
                new XAttribute("lang", "zh-CN"),
                new XAttribute("sz", FooterFontSizePoints * 100),
                new XElement(A + "solidFill",
                    new XElement(A + "srgbClr", new XAttribute("val", FooterGrayHex))),
                new XElement(A + "latin", new XAttribute("typeface", "+mn-lt")),
                new XElement(A + "ea", new XAttribute("typeface", "+mn-ea"))),
            new XElement(A + "t", slideNumber.ToString(CultureInfo.InvariantCulture)));

    private static XElement BuildFooterTextbox(uint id, string? text,
        double left, double width, double top, string alignment)
    {
        // algn 属于 a:pPr；直接写在 a:p 上 Schema 不认（会报 'algn' attribute is not declared）。
        var paragraph = new XElement(A + "p",
            new XElement(A + "pPr", new XAttribute("algn", alignment)));
        if (!string.IsNullOrEmpty(text))
        {
            paragraph.Add(new XElement(A + "r",
                new XElement(A + "rPr",
                    new XAttribute("lang", "zh-CN"),
                    new XAttribute("sz", FooterFontSizePoints * 100),
                    new XElement(A + "solidFill",
                        new XElement(A + "srgbClr", new XAttribute("val", FooterGrayHex))),
                    new XElement(A + "latin", new XAttribute("typeface", "+mn-lt")),
                    new XElement(A + "ea", new XAttribute("typeface", "+mn-ea"))),
                new XElement(A + "t", text)));
        }
        return new XElement(P + "sp",
            new XElement(P + "nvSpPr",
                new XElement(P + "cNvPr",
                    new XAttribute("id", id), new XAttribute("name", $"TextBox {id}")),
                new XElement(P + "cNvSpPr", new XAttribute("txBox", "1")),
                new XElement(P + "nvPr")),
            new XElement(P + "spPr",
                XfrmElement(Inches(left), Inches(top), Inches(width),
                    Inches(FooterBoxHeightInches)),
                new XElement(A + "prstGeom", new XAttribute("prst", "rect"),
                    new XElement(A + "avLst")),
                new XElement(A + "noFill")),
            new XElement(P + "txBody",
                // 不写 spAutoFit：保留声明的 0.35in 高度，避免 PowerPoint 打开后按文字缩框。
                new XElement(A + "bodyPr", new XAttribute("wrap", "square")),
                new XElement(A + "lstStyle"),
                paragraph));
    }

    // ── 节（p14:sectionLst）─────────────────────────────────────

    /// <summary>
    /// 把 sections 换算为 (name, first, last) 的 1 基闭区间。
    /// 每节覆盖 slide_start 到下一节 slide_start 前一张，末节到最后一张；
    /// 第一节 slide_start&gt;1 时，前面的幻灯片归入「默认节」。
    /// </summary>
    public static List<(string Name, int First, int Last)> SectionRanges(
        IReadOnlyList<PptSectionDef> sections, int slideCount)
    {
        var ranges = new List<(string Name, int First, int Last)>();
        if (sections.Count == 0)
            return ranges;
        if (slideCount <= 0)
            throw new NativeRuntimeException("演示文稿没有幻灯片，无法写入节");
        if (sections.Count > slideCount)
            throw new NativeUsageException(
                $"sections 有 {sections.Count} 个节，超过幻灯片总数 {slideCount}");

        var previous = 0;
        for (var index = 0; index < sections.Count; index++)
        {
            var section = sections[index];
            var name = (section.Name ?? "").Trim();
            if (name.Length == 0)
                throw new NativeUsageException($"sections[{index}] 的 name 不能为空");
            if (section.SlideStart < 1 || section.SlideStart > slideCount)
                throw new NativeUsageException(
                    $"节「{name}」的 slide_start={section.SlideStart} "
                    + $"超出范围 1..{slideCount}");
            if (section.SlideStart <= previous)
                throw new NativeUsageException(
                    $"节「{name}」的 slide_start={section.SlideStart} "
                    + $"未严格升序（前一节为 {previous}）");
            previous = section.SlideStart;
        }

        if (sections[0].SlideStart > 1)
            ranges.Add((DefaultSectionName, 1, sections[0].SlideStart - 1));
        for (var index = 0; index < sections.Count; index++)
        {
            var first = sections[index].SlideStart;
            var last = index + 1 < sections.Count
                ? sections[index + 1].SlideStart - 1
                : slideCount;
            ranges.Add((sections[index].Name.Trim(), first, last));
        }
        return ranges;
    }

    /// <summary>把 sections 写入 presentation.xml 的 p:extLst/p:ext/p14:sectionLst。</summary>
    public static void WriteSections(PresentationPart part,
        IReadOnlyList<PptSectionDef> sections)
    {
        var presentation = part.Presentation;
        var slideIds = presentation.SlideIdList?
            .Elements<XPresentation.SlideId>()
            .Select(slideId => slideId.Id?.Value ?? 0)
            .Where(id => id > 0)
            .ToList() ?? [];
        var ranges = SectionRanges(sections, slideIds.Count);
        if (ranges.Count == 0)
            return;

        var root = XDocument.Parse(presentation.OuterXml).Root
                   ?? throw new NativeRuntimeException("presentation.xml 解析失败");
        var extensionList = root.Element(P + "extLst");
        if (extensionList is null)
        {
            extensionList = new XElement(P + "extLst");
            root.Add(extensionList);
        }
        foreach (var old in extensionList.Elements(P + "ext")
                     .Where(ext => (string?)ext.Attribute("uri") == SectionExtensionUri)
                     .ToList())
            old.Remove();

        var sectionsElement = new XElement(P14 + "sectionLst",
            new XAttribute(XNamespace.Xmlns + "p14", P14));
        foreach (var (name, first, last) in ranges)
        {
            var slideIdList = new XElement(P14 + "sldIdLst");
            foreach (var slideId in slideIds.Skip(first - 1).Take(last - first + 1))
                slideIdList.Add(new XElement(P14 + "sldId",
                    new XAttribute("id", slideId.ToString(CultureInfo.InvariantCulture))));
            sectionsElement.Add(new XElement(P14 + "section",
                new XAttribute("name", name),
                new XAttribute("id", "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}"),
                slideIdList));
        }
        extensionList.Add(new XElement(P + "ext",
            new XAttribute("uri", SectionExtensionUri),
            sectionsElement));

        part.Presentation = new XPresentation.Presentation(
            root.ToString(SaveOptions.DisableFormatting));
    }

    // ── 放映设置（presProps.xml 的 p:showPr）─────────────────────

    /// <summary>
    /// 把放映设置写入 ppt/presProps.xml 的 p:presentationPr/p:showPr。
    /// CT_Presentation 不含 showPr（OpenXml SDK 校验报 invalid child），
    /// PowerPoint 也只从 presProps.xml 读，因此这里改 presProps 部件。
    /// </summary>
    public static void WriteShowSettings(PresentationPart part,
        PptShowSettingsDef settings)
    {
        var slideCount = SlideParts(part).Count;
        var start = settings.RangeStart;
        var end = settings.RangeEnd;
        if (start is not null || end is not null)
        {
            var rangeStart = start ?? 1;
            var rangeEnd = end ?? slideCount;
            if (rangeStart < 1 || rangeEnd < 1)
                throw new NativeUsageException(
                    "show_settings 的 range 起点/终点必须 >= 1（1 起）");
            if (rangeEnd < rangeStart)
                throw new NativeUsageException(
                    $"show_settings 的 range 不合法：range_end={rangeEnd} "
                    + $"小于 range_start={rangeStart}");
            if (rangeEnd > slideCount)
                throw new NativeUsageException(
                    $"show_settings 的 range 终点 {rangeEnd} 超出幻灯片总数 {slideCount}");
        }

        var propertiesPart = part.PresentationPropertiesPart
                             ?? part.AddNewPart<PresentationPropertiesPart>();
        var root = XDocument.Parse(propertiesPart.PresentationProperties?.OuterXml
                                   ?? "<p:presentationPr "
                                   + "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
                                   + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
                                   + "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/>")
            .Root ?? throw new NativeRuntimeException("presProps.xml 解析失败");

        foreach (var old in root.Elements(P + "showPr").ToList())
            old.Remove();

        var showProperties = new XElement(P + "showPr",
            new XAttribute("loop", settings.Loop ? "1" : "0"),
            new XAttribute("showNarration", settings.ShowNarration ? "1" : "0"),
            new XAttribute("showAnimation", settings.ShowAnimation ? "1" : "0"),
            new XAttribute("useTimings", settings.UseTimings ? "1" : "0"),
            new XElement(P + settings.ShowType switch
            {
                "browse" => "browse",
                "kiosk" => "kiosk",
                _ => "present",
            }));
        if (start is not null || end is not null)
        {
            showProperties.Add(new XElement(P + "sldRg",
                new XAttribute("st", (start ?? 1).ToString(CultureInfo.InvariantCulture)),
                new XAttribute("end", (end ?? slideCount)
                    .ToString(CultureInfo.InvariantCulture))));
        }

        var anchor = root.Element(P + "clrMru") ?? root.Element(P + "extLst");
        if (anchor is not null)
            anchor.AddBeforeSelf(showProperties);
        else
            root.Add(showProperties);

        propertiesPart.PresentationProperties = new XPresentation.PresentationProperties(
            root.ToString(SaveOptions.DisableFormatting));
    }

    // ── 备注母版/备注页 ──────────────────────────────────────────

    public static void EnsureNotesMaster(PresentationPart part)
    {
        if (part.NotesMasterPart is not null)
            return;
        var notesMasterPart = part.AddNewPart<NotesMasterPart>();
        var themePart = notesMasterPart.AddNewPart<ThemePart>();
        themePart.Theme = new XDrawing.Theme(ThemeXml);
        notesMasterPart.NotesMaster = new XPresentation.NotesMaster(
            BuildNotesMasterElement().ToString(SaveOptions.DisableFormatting));

        var presentation = part.Presentation;
        var list = presentation.NotesMasterIdList;
        if (list is null)
        {
            list = new XPresentation.NotesMasterIdList();
            var slideIdList = presentation.SlideIdList;
            var slideSize = presentation.SlideSize;
            if (slideIdList is not null)
                presentation.InsertBefore(list, slideIdList);
            else if (slideSize is not null)
                presentation.InsertBefore(list, slideSize);
            else
                presentation.AppendChild(list);
        }
        list.Append(new XPresentation.NotesMasterId
        {
            Id = part.GetIdOfPart(notesMasterPart),
        });
    }

    private static XElement BuildNotesMasterElement()
    {
        var notesBody = new XElement(P + "sp",
            new XElement(P + "nvSpPr",
                new XElement(P + "cNvPr", new XAttribute("id", "4"),
                    new XAttribute("name", "Notes Placeholder 4")),
                new XElement(P + "cNvSpPr",
                    new XElement(A + "spLocks", new XAttribute("noGrp", "1"))),
                new XElement(P + "nvPr",
                    new XElement(P + "ph",
                        new XAttribute("type", "body"),
                        new XAttribute("sz", "quarter"),
                        new XAttribute("idx", "3")))),
            new XElement(P + "spPr",
                XfrmElement(Inches(0.75), Inches(4.75), Inches(6.0), Inches(4.5)),
                new XElement(A + "prstGeom", new XAttribute("prst", "rect"),
                    new XElement(A + "avLst"))),
            new XElement(P + "txBody",
                new XElement(A + "bodyPr"),
                new XElement(A + "lstStyle"),
                new XElement(A + "p")));

        var slideImage = new XElement(P + "sp",
            new XElement(P + "nvSpPr",
                new XElement(P + "cNvPr", new XAttribute("id", "2"),
                    new XAttribute("name", "Slide Image Placeholder 1")),
                new XElement(P + "cNvSpPr",
                    new XElement(A + "spLocks",
                        new XAttribute("noGrp", "1"),
                        new XAttribute("noRot", "1"),
                        new XAttribute("noChangeAspect", "1"))),
                new XElement(P + "nvPr",
                    new XElement(P + "ph",
                        new XAttribute("type", "sldImg"),
                        new XAttribute("idx", "2")))),
            new XElement(P + "spPr",
                XfrmElement(Inches(1.25), Inches(0.75), Inches(5.0), Inches(3.75)),
                new XElement(A + "prstGeom", new XAttribute("prst", "rect"),
                    new XElement(A + "avLst")),
                new XElement(A + "noFill"),
                new XElement(A + "ln", new XAttribute("w", "12700"),
                    new XElement(A + "solidFill",
                        new XElement(A + "prstClr", new XAttribute("val", "black"))))));

        return new XElement(P + "notesMaster",
            new XAttribute(XNamespace.Xmlns + "a", A),
            new XAttribute(XNamespace.Xmlns + "r", R),
            new XAttribute(XNamespace.Xmlns + "p", P),
            new XElement(P + "cSld",
                new XElement(P + "bg",
                    new XElement(P + "bgRef", new XAttribute("idx", "1001"),
                        new XElement(A + "schemeClr", new XAttribute("val", "bg1")))),
                new XElement(P + "spTree",
                    GroupProperties(), GroupTransform(), slideImage, notesBody)),
            new XElement(P + "clrMap",
                new XAttribute("bg1", "lt1"), new XAttribute("tx1", "dk1"),
                new XAttribute("bg2", "lt2"), new XAttribute("tx2", "dk2"),
                new XAttribute("accent1", "accent1"), new XAttribute("accent2", "accent2"),
                new XAttribute("accent3", "accent3"), new XAttribute("accent4", "accent4"),
                new XAttribute("accent5", "accent5"), new XAttribute("accent6", "accent6"),
                new XAttribute("hlink", "hlink"), new XAttribute("folHlink", "folHlink")),
            new XElement(P + "notesStyle",
                new XElement(A + "lvl1pPr",
                    new XAttribute("marL", "0"),
                    new XAttribute("algn", "l"),
                    new XElement(A + "defRPr",
                        new XAttribute("sz", "1200"),
                        new XElement(A + "latin", new XAttribute("typeface", "+mn-lt")),
                        new XElement(A + "ea", new XAttribute("typeface", "+mn-ea")),
                        new XElement(A + "cs", new XAttribute("typeface", "+mn-cs"))))));
    }

    public static void AddNotes(SlidePart slidePart, string text)
    {
        var presentationPart = slidePart.GetParentParts()
            .OfType<PresentationPart>().FirstOrDefault()
            ?? throw new NativeRuntimeException("幻灯片不属于演示文稿");
        EnsureNotesMaster(presentationPart);
        var notesMaster = presentationPart.NotesMasterPart
                         ?? throw new NativeRuntimeException("备注母版创建失败");

        var notesPart = slidePart.AddNewPart<NotesSlidePart>();
        notesPart.AddPart(notesMaster);
        notesPart.NotesSlide = new XPresentation.NotesSlide(
            BuildNotesSlideElement(text).ToString(SaveOptions.DisableFormatting));
    }

    private static XElement BuildNotesSlideElement(string text)
    {
        var body = new XElement(P + "sp",
            new XElement(P + "nvSpPr",
                new XElement(P + "cNvPr", new XAttribute("id", "3"),
                    new XAttribute("name", "Notes Placeholder 2")),
                new XElement(P + "cNvSpPr",
                    new XElement(A + "spLocks", new XAttribute("noGrp", "1"))),
                new XElement(P + "nvPr",
                    new XElement(P + "ph",
                        new XAttribute("type", "body"),
                        new XAttribute("idx", "3"),
                        new XAttribute("sz", "quarter")))),
            new XElement(P + "spPr"),
            new XElement(P + "txBody",
                new XElement(A + "bodyPr"),
                new XElement(A + "lstStyle"),
                PptFileContent.BuildPlainParagraph(text)));

        var image = new XElement(P + "sp",
            new XElement(P + "nvSpPr",
                new XElement(P + "cNvPr", new XAttribute("id", "2"),
                    new XAttribute("name", "Slide Image Placeholder 1")),
                new XElement(P + "cNvSpPr",
                    new XElement(A + "spLocks", new XAttribute("noGrp", "1"))),
                new XElement(P + "nvPr",
                    new XElement(P + "ph",
                        new XAttribute("type", "sldImg"),
                        new XAttribute("idx", "2")))),
            new XElement(P + "spPr"));

        return new XElement(P + "notes",
            new XAttribute(XNamespace.Xmlns + "a", A),
            new XAttribute(XNamespace.Xmlns + "r", R),
            new XAttribute(XNamespace.Xmlns + "p", P),
            new XElement(P + "cSld",
                new XElement(P + "spTree",
                    GroupProperties(), GroupTransform(), image, body)),
            new XElement(P + "clrMapOvr",
                new XElement(A + "masterClrMapping")));
    }

    public static string? ReadNotesText(SlidePart slidePart)
    {
        var notesPart = slidePart.NotesSlidePart;
        if (notesPart?.NotesSlide is null)
            return null;
        var document = XDocument.Parse(notesPart.NotesSlide.OuterXml);
        var body = ShapeElements(document).FirstOrDefault(shape =>
        {
            var info = PlaceholderInfo(shape);
            return info is not null
                   && string.Equals(info.Value.Type, "body", StringComparison.OrdinalIgnoreCase);
        });
        if (body is null)
            return null;
        var text = TxBodyText(body.Element(P + "txBody"));
        return string.IsNullOrWhiteSpace(text) ? null : text.Trim();
    }

    // ── 占位符继承与几何 ────────────────────────────────────────

    public sealed class SlideContext
    {
        public XDocument? Layout { get; init; }
        public XDocument? Master { get; init; }
    }

    public static SlideContext ContextFor(SlidePart slidePart)
    {
        var layoutPart = slidePart.SlideLayoutPart;
        if (layoutPart is null)
            return new SlideContext();
        var layout = layoutPart.SlideLayout is null
            ? null
            : XDocument.Parse(layoutPart.SlideLayout.OuterXml);
        var master = layoutPart.SlideMasterPart?.SlideMaster is { } masterElement
            ? XDocument.Parse(masterElement.OuterXml)
            : null;
        return new SlideContext { Layout = layout, Master = master };
    }

    public static (string? Type, int Idx)? PlaceholderInfo(XElement shape)
    {
        XName[] containers = [P + "nvSpPr", P + "nvPicPr",
            P + "nvGraphicFramePr", P + "nvCxnSpPr", P + "nvGrpSpPr"];
        foreach (var container in containers)
        {
            var ph = shape.Element(container)?.Element(P + "nvPr")?.Element(P + "ph");
            if (ph is null)
                continue;
            var type = (string?)ph.Attribute("type");
            var idxText = (string?)ph.Attribute("idx");
            var idx = idxText is not null && int.TryParse(idxText, out var parsed) ? parsed : 0;
            return (type, idx);
        }
        return null;
    }

    public static bool IsPlaceholder(XElement shape) => PlaceholderInfo(shape) is not null;

    public static XElement? FindPlaceholderShape(XDocument? document, string? type, int idx)
    {
        if (document?.Root is null)
            return null;
        var shapes = ShapeElements(document).ToList();
        return shapes.FirstOrDefault(shape =>
        {
            var info = PlaceholderInfo(shape);
            return info is not null && info.Value.Idx == idx
                   && (type is null || info.Value.Type is null
                       || string.Equals(info.Value.Type, type,
                           StringComparison.OrdinalIgnoreCase));
        }) ?? shapes.FirstOrDefault(shape =>
        {
            var info = PlaceholderInfo(shape);
            return info is not null && info.Value.Idx == idx;
        });
    }

    private static XElement? XfrmOf(XElement shape)
    {
        if (shape.Name == P + "graphicFrame")
            return shape.Element(P + "xfrm");
        if (shape.Name == P + "grpSp")
            return shape.Element(P + "grpSpPr")?.Element(A + "xfrm");
        return shape.Element(P + "spPr")?.Element(A + "xfrm");
    }

    public static (long X, long Y, long Width, long Height)? GetShapeBox(
        XElement shape, SlideContext? context = null)
    {
        var xfrm = XfrmOf(shape);
        if (xfrm is null && context is not null && PlaceholderInfo(shape) is { } info)
        {
            var layoutShape = FindPlaceholderShape(context.Layout, info.Type, info.Idx);
            if (layoutShape is not null)
                xfrm = XfrmOf(layoutShape);
            if (xfrm is null)
            {
                var masterShape = FindPlaceholderShape(context.Master, info.Type, info.Idx);
                if (masterShape is not null)
                    xfrm = XfrmOf(masterShape);
            }
        }
        if (xfrm is null)
            return null;
        var off = xfrm.Element(A + "off");
        var ext = xfrm.Element(A + "ext");
        if (off is null || ext is null)
            return null;
        if (!long.TryParse((string?)off.Attribute("x"), out var x)
            || !long.TryParse((string?)off.Attribute("y"), out var y)
            || !long.TryParse((string?)ext.Attribute("cx"), out var width)
            || !long.TryParse((string?)ext.Attribute("cy"), out var height))
            return null;
        if (width <= 0 || height <= 0)
            return null;
        return (x, y, width, height);
    }

    public static bool SetShapeBox(XElement shape, long x, long y, long width, long height)
    {
        XElement? xfrm;
        if (shape.Name == P + "graphicFrame")
        {
            xfrm = shape.Element(P + "xfrm");
            if (xfrm is null)
            {
                xfrm = new XElement(P + "xfrm");
                shape.AddFirst(xfrm);
            }
        }
        else if (shape.Name == P + "grpSp")
        {
            var groupProperties = shape.Element(P + "grpSpPr");
            if (groupProperties is null)
                return false;
            xfrm = groupProperties.Element(A + "xfrm");
            if (xfrm is null)
            {
                xfrm = new XElement(A + "xfrm");
                groupProperties.AddFirst(xfrm);
            }
        }
        else
        {
            var shapeProperties = shape.Element(P + "spPr");
            if (shapeProperties is null)
                return false;
            xfrm = shapeProperties.Element(A + "xfrm");
            if (xfrm is null)
            {
                xfrm = new XElement(A + "xfrm");
                shapeProperties.AddFirst(xfrm);
            }
        }
        var off = xfrm.Element(A + "off");
        var ext = xfrm.Element(A + "ext");
        if (off is null)
        {
            off = new XElement(A + "off");
            xfrm.AddFirst(off);
        }
        if (ext is null)
        {
            ext = new XElement(A + "ext");
            off.AddAfterSelf(ext);
        }
        off.SetAttributeValue("x", x);
        off.SetAttributeValue("y", y);
        ext.SetAttributeValue("cx", width);
        ext.SetAttributeValue("cy", height);
        return true;
    }

    public static string ShapeName(XElement shape)
    {
        XName[] containers = [P + "nvSpPr", P + "nvPicPr",
            P + "nvGraphicFramePr", P + "nvCxnSpPr", P + "nvGrpSpPr"];
        foreach (var container in containers)
        {
            var name = shape.Element(container)?.Element(P + "cNvPr")?.Attribute("name")?.Value;
            if (name is not null)
                return name.Trim();
        }
        return "";
    }

    public static int ShapeId(XElement shape)
    {
        XName[] containers = [P + "nvSpPr", P + "nvPicPr",
            P + "nvGraphicFramePr", P + "nvCxnSpPr", P + "nvGrpSpPr"];
        foreach (var container in containers)
        {
            var id = shape.Element(container)?.Element(P + "cNvPr")?.Attribute("id")?.Value;
            if (id is not null && int.TryParse(id, out var parsed))
                return parsed;
        }
        return 0;
    }

    public static XElement? CnvPr(XElement shape)
    {
        XName[] containers = [P + "nvSpPr", P + "nvPicPr",
            P + "nvGraphicFramePr", P + "nvCxnSpPr", P + "nvGrpSpPr"];
        var container = containers.Select(name => shape.Element(name))
            .FirstOrDefault(element => element is not null);
        return container?.Element(P + "cNvPr");
    }

    // ── 文本读取 ─────────────────────────────────────────────────

    /// <summary>
    /// 按文档顺序拼接 a:r / a:fld 文本与 a:br 换行；a:br 映射为 \v（与 python-pptx
    /// TextFrame.text 语义一致），段落之间由调用方用 \n 连接。
    /// </summary>
    public static string ParagraphText(XElement? paragraph)
    {
        if (paragraph is null)
            return "";
        var builder = new System.Text.StringBuilder();
        foreach (var child in paragraph.Elements())
        {
            if (child.Name == A + "r" || child.Name == A + "fld")
            {
                foreach (var text in child.Elements(A + "t"))
                    builder.Append(text.Value);
            }
            else if (child.Name == A + "br")
            {
                builder.Append('\v');
            }
        }
        return builder.ToString();
    }

    public static bool TxBodyHasField(XElement? textBody) =>
        textBody?.Descendants(A + "fld").Any() == true;

    public static string TxBodyText(XElement? textBody)
    {
        if (textBody is null)
            return "";
        return string.Join("\n", textBody.Elements(A + "p").Select(ParagraphText));
    }

    public static string TextFrameText(XElement shape)
    {
        var textBody = shape.Element(P + "txBody");
        // 目前唯一调用方是 batch replace；整体重建文本框会丢字段，明确拒绝而不是静默丢失。
        if (TxBodyHasField(textBody))
            throw new BatchEditException(
                "unsupported_operation",
                "目标形状文本包含字段（a:fld），replace 会丢失字段；请先转换为普通文本",
                suggestion: "把字段改为普通文字后再执行 replace");
        return TxBodyText(textBody);
    }

    public static int ParagraphLevel(XElement? paragraph) =>
        int.TryParse((string?)paragraph?.Element(A + "pPr")?.Attribute("lvl"), out var level)
            ? level
            : 0;

    public static XElement? TextBodyOf(XElement shape) => shape.Element(P + "txBody");

    /// <summary>遍历一个 slide/notes 文档中的文本帧（文本框、表格单元格）。</summary>
    public static IEnumerable<XElement> TextBodiesInDocument(XDocument document)
    {
        foreach (var shape in ShapeElements(document))
        {
            var textBody = shape.Element(P + "txBody");
            if (textBody is not null)
                yield return textBody;
            if (shape.Name == P + "graphicFrame")
            {
                var table = shape.Element(A + "graphic")?.Element(A + "graphicData")
                    ?.Element(A + "tbl");
                if (table is not null)
                {
                    foreach (var cell in table.Descendants(A + "tc"))
                    {
                        var cellBody = cell.Element(A + "txBody");
                        if (cellBody is not null)
                            yield return cellBody;
                    }
                }
            }
        }
    }

    /// <summary>遍历所有文本帧（幻灯片文本框、表格单元格、备注）。</summary>
    public static IEnumerable<XElement> AllTextBodies(SlidePart slidePart)
    {
        foreach (var textBody in TextBodiesInDocument(SlideDocument(slidePart)))
            yield return textBody;
        var notesPart = slidePart.NotesSlidePart;
        if (notesPart?.NotesSlide is not null)
        {
            var notesDocument = XDocument.Parse(notesPart.NotesSlide.OuterXml);
            foreach (var textBody in TextBodiesInDocument(notesDocument))
                yield return textBody;
        }
    }

    // ── 主题（默认）─────────────────────────────────────────────

    private const string ThemeXml = """
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
      <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
      <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface="宋体"/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface="宋体"/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"><a:shade val="50000"/></a:schemeClr></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>
""";
}
