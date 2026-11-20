using System.Globalization;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeNative;

/// <summary>DrawingML 锚点、图片与旧式批注 VML 的底层写入。</summary>
internal static class ExcelFileDrawing
{
    public static readonly XNamespace Xdr =
        "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    public static readonly XNamespace A =
        "http://schemas.openxmlformats.org/drawingml/2006/main";
    public static readonly XNamespace ROwn =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public static readonly XNamespace C =
        "http://schemas.openxmlformats.org/drawingml/2006/chart";

    private const long EmuPerPixel = 9525;
    private const double EmuPerCentimeter = 360000;

    /// <summary>图片/图表尺寸上限（cm），避免非法或极端值让 EMU 计算溢出。</summary>
    public const double MaxDrawingSizeCm = 1000;
    private const long MaxEmu = (long)(MaxDrawingSizeCm * EmuPerCentimeter);

    public static long CmToEmu(double centimeters, string label)
    {
        if (!double.IsFinite(centimeters)
            || centimeters <= 0
            || centimeters > MaxDrawingSizeCm)
        {
            throw new NativeRuntimeException(
                $"{label} 必须是 (0, {MaxDrawingSizeCm}] cm 之间的有限正数: {centimeters}");
        }
        try
        {
            var emu = checked((long)Math.Round(centimeters * EmuPerCentimeter));
            if (emu <= 0 || emu > MaxEmu)
                throw new OverflowException();
            return emu;
        }
        catch (OverflowException)
        {
            throw new NativeRuntimeException($"{label} 换算为 EMU 后超出允许范围: {centimeters} cm");
        }
    }

    public static string EnsureDrawingsPart(WorksheetPart worksheetPart)
    {
        var part = worksheetPart.DrawingsPart
                   ?? worksheetPart.AddNewPart<DrawingsPart>();
        return worksheetPart.GetIdOfPart(part);
    }

    public static XDocument LoadDrawing(DrawingsPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        if (stream.Length > 0)
        {
            using var reader = new StreamReader(stream, Encoding.UTF8, true);
            return XDocument.Load(reader);
        }
        return new XDocument(new XElement(Xdr + "wsDr",
            new XAttribute(XNamespace.Xmlns + "xdr", Xdr.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "a", A.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "r", ROwn.NamespaceName)));
    }

    public static void SaveDrawing(DrawingsPart part, XDocument document)
    {
        using var buffer = new MemoryStream();
        document.Save(buffer, SaveOptions.DisableFormatting);
        buffer.Position = 0;
        part.FeedData(buffer);
    }

    public static int NextShapeId(XDocument document)
    {
        var maxId = 0;
        foreach (var element in document.Descendants())
        {
            if (element.Name.LocalName != "cNvPr")
                continue;
            var value = element.Attribute("id")?.Value;
            if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id)
                && id > maxId)
                maxId = id;
        }
        return maxId + 1;
    }

    /// <summary>按 CT_Drawing 顺序插入锚点：twoCellAnchor → oneCellAnchor → absoluteAnchor → extLst。</summary>
    public static void AddDrawingAnchor(XDocument document, XElement anchor)
    {
        var root = document.Root
                   ?? throw new NativeRuntimeException("drawing 缺少根节点");
        var rank = AnchorRank(anchor.Name.LocalName);
        var before = root.Elements().FirstOrDefault(element =>
            AnchorRank(element.Name.LocalName) > rank);
        if (before is not null)
            before.AddBeforeSelf(anchor);
        else
            root.Add(anchor);
    }

    private static int AnchorRank(string localName) => localName switch
    {
        "twoCellAnchor" => 0,
        "oneCellAnchor" => 1,
        "absoluteAnchor" => 2,
        "extLst" => 3,
        _ => 4,
    };

    public static XElement Marker(XName name, int col, int row) => new(name,
        new XElement(Xdr + "col", col.ToString(CultureInfo.InvariantCulture)),
        new XElement(Xdr + "colOff", "0"),
        new XElement(Xdr + "row", row.ToString(CultureInfo.InvariantCulture)),
        new XElement(Xdr + "rowOff", "0"));

    public static XElement ClientData() => new(Xdr + "clientData");

    public static void AddImage(WorksheetPart worksheetPart, ExcelImageDef image)
    {
        var path = image.Path;
        if (!File.Exists(path))
            throw new NativeRuntimeException($"图片文件不存在: {path}");
        if (!ExcelA1.TryParseCell(image.Anchor, out var col, out var row))
            throw new NativeUsageException($"无效的图片锚点单元格: {image.Anchor}");

        var drawingsPart = worksheetPart.DrawingsPart
                           ?? worksheetPart.AddNewPart<DrawingsPart>();
        var document = LoadDrawing(drawingsPart);
        var shapeId = NextShapeId(document);

        var emu = ComputeImageSize(image, path);
        var contentType = ImageContentType(path);
        var imagePart = drawingsPart.AddNewPart<ImagePart>(contentType);
        using (var stream = File.OpenRead(path))
            imagePart.FeedData(stream);
        var relId = drawingsPart.GetIdOfPart(imagePart);

        var anchor = new XElement(Xdr + "oneCellAnchor",
            Marker(Xdr + "from", col - 1, row - 1),
            new XElement(Xdr + "ext",
                new XAttribute("cx", emu.Width.ToString(CultureInfo.InvariantCulture)),
                new XAttribute("cy", emu.Height.ToString(CultureInfo.InvariantCulture))),
            new XElement(Xdr + "pic",
                new XElement(Xdr + "nvPicPr",
                    new XElement(Xdr + "cNvPr",
                        new XAttribute("id", shapeId.ToString(CultureInfo.InvariantCulture)),
                        new XAttribute("name", $"Picture {shapeId}")),
                    new XElement(Xdr + "cNvPicPr",
                        new XElement(A + "picLocks", new XAttribute("noChangeAspect", "1")))),
                new XElement(Xdr + "blipFill",
                    new XElement(A + "blip", new XAttribute(ROwn + "embed", relId)),
                    new XElement(A + "stretch", new XElement(A + "fillRect"))),
                new XElement(Xdr + "spPr",
                    new XElement(A + "xfrm",
                        new XElement(A + "off", new XAttribute("x", "0"), new XAttribute("y", "0")),
                        new XElement(A + "ext",
                            new XAttribute("cx", emu.Width.ToString(CultureInfo.InvariantCulture)),
                            new XAttribute("cy", emu.Height.ToString(CultureInfo.InvariantCulture)))),
                    new XElement(A + "prstGeom", new XAttribute("prst", "rect"),
                        new XElement(A + "avLst")))),
            ClientData());

        AddDrawingAnchor(document, anchor);
        SaveDrawing(drawingsPart, document);
    }

    /// <summary>写入旧式批注 VML 绘图，返回 legacyDrawing 关系 id。</summary>
    public static string AddCommentsVml(
        WorksheetPart worksheetPart, IReadOnlyList<ExcelCommentDef> comments)
    {
        var vmlPart = worksheetPart.AddNewPart<VmlDrawingPart>();
        var builder = new StringBuilder();
        builder.Append("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"");
        builder.Append(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"");
        builder.Append(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
        builder.Append("<o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout>");
        builder.Append("<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\"");
        builder.Append(" path=\"m,l,21600r21600,l21600,xe\">");
        builder.Append("<v:stroke joinstyle=\"miter\"/>");
        builder.Append("<v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>");
        builder.Append("</v:shapetype>");

        var index = 0;
        foreach (var comment in comments)
        {
            if (!ExcelA1.TryParseCell(comment.Cell, out var col, out var row))
                continue;
            var shapeId = 1025 + index;
            builder.Append("<v:shape id=\"_x0000_s");
            builder.Append(shapeId.ToString(CultureInfo.InvariantCulture));
            builder.Append("\" type=\"#_x0000_t202\"");
            builder.Append(" style=\"position:absolute;margin-left:59.25pt;margin-top:1.5pt;");
            builder.Append("width:108pt;height:59.25pt;z-index:");
            builder.Append((index + 1).ToString(CultureInfo.InvariantCulture));
            builder.Append(";visibility:hidden\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\">");
            builder.Append("<v:fill color2=\"#ffffe1\"/>");
            builder.Append("<v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>");
            builder.Append("<v:path o:connecttype=\"none\"/>");
            builder.Append("<v:textbox style=\"mso-direction-alt:auto\">");
            builder.Append("<div style=\"text-align:left\"></div></v:textbox>");
            builder.Append("<x:ClientData ObjectType=\"Note\">");
            builder.Append("<x:MoveWithCells/><x:SizeWithCells/>");
            // VML Anchor 使用 0-based 列/行；四角裁剪到 16384×1048576 网格内。
            var columnIndex = col - 1;
            var rowIndex = row - 1;
            var rightColumn = Math.Min(columnIndex + 2, 16383);
            var bottomRow = Math.Min(rowIndex + 4, 1048575);
            var anchor = string.Create(CultureInfo.InvariantCulture,
                $"{columnIndex}, 15, {rowIndex}, 2, {rightColumn}, 15, {bottomRow}, 4");
            builder.Append("<x:Anchor>");
            builder.Append(anchor);
            builder.Append("</x:Anchor>");
            builder.Append("<x:AutoFill>False</x:AutoFill>");
            builder.Append("<x:Row>");
            builder.Append((row - 1).ToString(CultureInfo.InvariantCulture));
            builder.Append("</x:Row><x:Column>");
            builder.Append((col - 1).ToString(CultureInfo.InvariantCulture));
            builder.Append("</x:Column></x:ClientData></v:shape>");
            index++;
        }
        builder.Append("</xml>");

        using var buffer = new MemoryStream(Encoding.UTF8.GetBytes(builder.ToString()));
        vmlPart.FeedData(buffer);
        return worksheetPart.GetIdOfPart(vmlPart);
    }

    public static void EnsureDrawingElement(S.Worksheet worksheet, string relId)
    {
        if (worksheet.GetFirstChild<S.Drawing>() is not null)
            return;
        var drawing = new S.Drawing { Id = relId };
        var anchor = worksheet.ChildElements.FirstOrDefault(element =>
            element.LocalName is "legacyDrawing" or "legacyDrawingHF" or "picture"
                or "oleObjects" or "controls" or "webPublishItems" or "tableParts" or "extLst");
        if (anchor is not null)
            worksheet.InsertBefore(drawing, anchor);
        else
            worksheet.Append(drawing);
    }

    // ── 图片尺寸 ─────────────────────────────────────────────────

    private static (long Width, long Height) ComputeImageSize(ExcelImageDef image, string path)
    {
        var pixels = TryReadImagePixelSize(path, out var pixelWidth, out var pixelHeight)
            ? (Width: pixelWidth, Height: pixelHeight)
            : (Width: 0d, Height: 0d);

        if (image.Width is not null && image.Height is not null)
        {
            return (
                CmToEmu(image.Width.Value, "image.width"),
                CmToEmu(image.Height.Value, "image.height"));
        }
        if (image.Width is not null)
        {
            var width = CmToEmu(image.Width.Value, "image.width");
            var height = pixels.Width > 0 && pixels.Height > 0
                ? ScaleEmu(width, pixels.Height, pixels.Width, "image.height")
                : width;
            return (width, height);
        }
        if (image.Height is not null)
        {
            var height = CmToEmu(image.Height.Value, "image.height");
            var width = pixels.Width > 0 && pixels.Height > 0
                ? ScaleEmu(height, pixels.Width, pixels.Height, "image.width")
                : height;
            return (width, height);
        }
        if (pixels.Width <= 0 || pixels.Height <= 0)
            throw new NativeRuntimeException(
                $"无法读取图片尺寸，请显式提供 width/height: {path}");
        return (
            PixelsToEmu(pixels.Width, path),
            PixelsToEmu(pixels.Height, path));
    }

    private static long ScaleEmu(long source, double numerator, double denominator, string label)
    {
        var scaled = source * numerator / denominator;
        if (!double.IsFinite(scaled) || scaled < 1 || scaled > MaxEmu)
            throw new NativeRuntimeException($"{label} 按原图比例换算后超出允许范围");
        return (long)Math.Round(scaled);
    }

    private static long PixelsToEmu(double pixels, string path)
    {
        try
        {
            var emu = checked((long)Math.Round(pixels * EmuPerPixel));
            if (emu <= 0 || emu > MaxEmu)
                throw new OverflowException();
            return emu;
        }
        catch (OverflowException)
        {
            throw new NativeRuntimeException($"图片像素尺寸换算为 EMU 后超出允许范围: {path}");
        }
    }

    private static bool TryReadImagePixelSize(string path, out double width, out double height)
    {
        width = 0;
        height = 0;
        var buffer = new byte[512 * 1024];
        int read;
        using (var stream = File.OpenRead(path))
            read = stream.Read(buffer, 0, buffer.Length);
        if (read < 26)
            return false;

        // PNG
        if (buffer[0] == 0x89 && buffer[1] == 0x50 && buffer[2] == 0x4E && buffer[3] == 0x47)
        {
            width = ReadUInt32BigEndian(buffer, 16);
            height = ReadUInt32BigEndian(buffer, 20);
            return width > 0 && height > 0;
        }
        // GIF
        if (buffer[0] == 'G' && buffer[1] == 'I' && buffer[2] == 'F')
        {
            width = buffer[6] | (buffer[7] << 8);
            height = buffer[8] | (buffer[9] << 8);
            return width > 0 && height > 0;
        }
        // BMP
        if (buffer[0] == 'B' && buffer[1] == 'M')
        {
            width = Math.Abs(ReadInt32LittleEndian(buffer, 18));
            height = Math.Abs(ReadInt32LittleEndian(buffer, 22));
            return width > 0 && height > 0;
        }
        // JPEG
        if (buffer[0] == 0xFF && buffer[1] == 0xD8)
        {
            var offset = 2;
            while (offset + 9 < read)
            {
                if (buffer[offset] != 0xFF)
                {
                    offset++;
                    continue;
                }
                var marker = buffer[offset + 1];
                if (marker is 0xD8 or 0xD9)
                {
                    offset += 2;
                    continue;
                }
                var length = (buffer[offset + 2] << 8) | buffer[offset + 3];
                if (length < 2)
                    break;
                if (marker is 0xC0 or 0xC1 or 0xC2 or 0xC3
                    or 0xC5 or 0xC6 or 0xC7 or 0xC9 or 0xCA or 0xCB
                    or 0xCD or 0xCE or 0xCF)
                {
                    height = (buffer[offset + 5] << 8) | buffer[offset + 6];
                    width = (buffer[offset + 7] << 8) | buffer[offset + 8];
                    return width > 0 && height > 0;
                }
                offset += 2 + length;
            }
        }
        return false;
    }

    private static uint ReadUInt32BigEndian(byte[] buffer, int offset) =>
        (uint)((buffer[offset] << 24) | (buffer[offset + 1] << 16)
               | (buffer[offset + 2] << 8) | buffer[offset + 3]);

    private static int ReadInt32LittleEndian(byte[] buffer, int offset) =>
        buffer[offset] | (buffer[offset + 1] << 8)
        | (buffer[offset + 2] << 16) | (buffer[offset + 3] << 24);

    private static string ImageContentType(string path) =>
        Path.GetExtension(path).ToLowerInvariant() switch
        {
            ".png" => "image/png",
            ".jpg" or ".jpeg" => "image/jpeg",
            ".gif" => "image/gif",
            ".bmp" => "image/bmp",
            var extension => throw new NativeRuntimeException(
                $"不支持的图片格式: {extension}（支持 png/jpg/jpeg/gif/bmp）"),
        };
}
