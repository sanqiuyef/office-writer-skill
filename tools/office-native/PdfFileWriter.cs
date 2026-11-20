using System.Buffers.Binary;
using System.Globalization;
using System.Text;
using System.Text.Json;
using PdfSharpCore.Drawing;
using PdfSharpCore.Fonts;
using PdfSharpCore.Pdf;

namespace OfficeNative;

/// <summary>
/// create-pdf-document: PyMuPDF/fpdf2 replacement built on PdfSharpCore.
/// All text is laid out with an embedded TrueType CJK face; TTC collections are
/// reduced to their first face before embedding because PdfSharpCore rejects TTC.
/// </summary>
internal static class PdfFileWriter
{
    private const double MmToPt = 72.0 / 25.4;

    public static Dictionary<string, object?> RunCreatePdfDocument(JsonElement p)
    {
        var requestedPath = Path.GetFullPath(NativeCommon.RequireString(p, "output_path"));
        if (!Path.GetExtension(requestedPath).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            throw new NativeUsageException("输出文件必须使用 .pdf 扩展名");
        if (requestedPath.IndexOfAny(['<', '>', '"', '|', '?', '*']) >= 0)
            throw new NativeUsageException($"路径包含非法字符: {requestedPath}");

        var content = PdfContent.Parse(p);
        if (content.Sections.Count == 0)
            throw new NativeRuntimeException("content.sections 不能为空");

        var pageSetup = PdfPageSetup.Parse(p);
        var fontConfig = PdfFontConfig.Parse(p);
        var faceName = PdfFontRegistry.Register(ResolveFontPath(fontConfig));
        var outputPath = ResolveOutputPath(requestedPath, p);

        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
        var tempPath = Path.Combine(
            Path.GetDirectoryName(outputPath)!,
            $".pdf-{Guid.NewGuid():N}.pdf");
        string? backupPath = null;
        try
        {
            Render(tempPath, content, pageSetup, fontConfig, faceName);
            if (!File.Exists(tempPath) || new FileInfo(tempPath).Length == 0)
                throw new NativeRuntimeException("PDF 生成未产生有效文件");

            if (File.Exists(outputPath) && new FileInfo(outputPath).Length > 0)
                backupPath = NativeCommon.Backup(outputPath);
            File.Move(tempPath, outputPath, overwrite: true);
        }
        catch (NativeUsageException)
        {
            throw;
        }
        catch (NativeRuntimeException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new NativeRuntimeException($"创建 PDF 失败: {ex.Message}");
        }
        finally
        {
            if (File.Exists(tempPath))
            {
                try { File.Delete(tempPath); } catch { /* temp file cleanup */ }
            }
        }

        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["file_path"] = outputPath,
            ["file_size"] = new FileInfo(outputPath).Length,
            ["message"] = $"成功创建 PDF 文档，共 {content.Sections.Count} 个内容块",
            ["backup_path"] = backupPath,
        };
    }

    private static void Render(
        string path, PdfContent content, PdfPageSetup pageSetup,
        PdfFontConfig fontConfig, string faceName)
    {
        var document = new PdfDocument();
        document.Info.Creator = "office-native";
        using (var canvas = new PdfCanvas(document, pageSetup, fontConfig, faceName))
        {
            canvas.NewPage();
            foreach (var section in content.Sections)
                canvas.RenderSection(section);
            canvas.EndPage();
        }
        document.Save(path);
        document.Close();
    }

    private static string ResolveFontPath(PdfFontConfig config)
    {
        var fontsDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        if (string.IsNullOrEmpty(fontsDirectory))
        {
            var windows = Environment.GetEnvironmentVariable("WINDIR")
                          ?? Environment.GetEnvironmentVariable("SystemRoot")
                          ?? @"C:\Windows";
            fontsDirectory = Path.Combine(windows, "Fonts");
        }

        var candidates = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["simhei"] = Path.Combine(fontsDirectory, "simhei.ttf"),
            ["simsun"] = Path.Combine(fontsDirectory, "simsun.ttc"),
            ["msyh"] = Path.Combine(fontsDirectory, "msyh.ttc"),
            ["msyhbd"] = Path.Combine(fontsDirectory, "msyhbd.ttc"),
            ["simkai"] = Path.Combine(fontsDirectory, "simkai.ttf"),
            ["simfang"] = Path.Combine(fontsDirectory, "simfang.ttf"),
            ["fangsong"] = Path.Combine(fontsDirectory, "simfang.ttf"),
            ["kaiti"] = Path.Combine(fontsDirectory, "simkai.ttf"),
        };

        if (!string.IsNullOrWhiteSpace(config.CnFontPath))
        {
            if (!File.Exists(config.CnFontPath))
                throw new NativeRuntimeException(
                    $"指定的中文字体文件不存在: {config.CnFontPath}");
            return config.CnFontPath;
        }

        if (!string.IsNullOrWhiteSpace(config.CnFont))
        {
            var key = config.CnFont.Trim();
            if (candidates.TryGetValue(key, out var mapped) && File.Exists(mapped))
                return mapped;
            if (File.Exists(key))
                return Path.GetFullPath(key);
            throw new NativeRuntimeException(
                $"指定的中文字体文件不存在: {key}。可用候选: "
                + string.Join(", ", candidates.Keys));
        }

        foreach (var name in new[] { "simsun", "simhei", "msyh", "simkai", "simfang" })
        {
            if (candidates.TryGetValue(name, out var path) && File.Exists(path))
                return path;
        }

        if (Directory.Exists(fontsDirectory))
        {
            foreach (var path in Directory.EnumerateFiles(fontsDirectory)
                         .Where(file => file.EndsWith(".ttf", StringComparison.OrdinalIgnoreCase)
                                        || file.EndsWith(".ttc", StringComparison.OrdinalIgnoreCase))
                         .OrderBy(file => file, StringComparer.OrdinalIgnoreCase))
                return path;
        }

        throw new NativeRuntimeException(
            "未找到可用的中文字体文件。请通过 font_config.cn_font_path 指定字体路径，"
            + "或在系统字体目录安装中文字体（如 simhei.ttf / simsun.ttc）。");
    }

    private static string ResolveOutputPath(string path, JsonElement p)
    {
        var mode = NativeCommon.OptString(p, "if_exists");
        var overwrite = NativeCommon.OptBool(p, "overwrite", false);
        var rename = NativeCommon.OptBool(p, "rename", false);
        var confirm = NativeCommon.OptBool(p, "confirm_overwrite", false);
        if (string.IsNullOrWhiteSpace(mode))
            mode = rename ? "rename" : overwrite ? "overwrite" : "error";
        if (mode is not ("error" or "overwrite" or "rename"))
            throw new NativeUsageException(
                $"if_exists 必须是 error/overwrite/rename 之一，收到 {mode}");

        var full = Path.GetFullPath(path);
        if (mode == "rename" && File.Exists(full) && new FileInfo(full).Length > 0)
        {
            var directory = Path.GetDirectoryName(full)!;
            var stem = Path.GetFileNameWithoutExtension(full);
            var extension = Path.GetExtension(full);
            for (var counter = 1; ; counter++)
            {
                var candidate = Path.Combine(directory, $"{stem} ({counter}){extension}");
                if (!File.Exists(candidate))
                    return candidate;
            }
        }
        NativeCommon.EnsureOutput(full, mode == "overwrite", confirm);
        return full;
    }
}

/// <summary>Registers every requested face with one process-wide resolver.</summary>
internal static class PdfFontRegistry
{
    private static readonly object Gate = new();
    private static readonly Dictionary<string, byte[]> Faces = new(StringComparer.Ordinal);
    private static PdfFaceResolver? _resolver;
    private static int _counter;

    public static string Register(string fontPath)
    {
        byte[] bytes;
        try
        {
            bytes = File.ReadAllBytes(fontPath);
        }
        catch (Exception ex)
        {
            throw new NativeRuntimeException($"读取中文字体失败: {ex.Message}");
        }

        bytes = TrueTypeCollection.ExtractFirstFace(bytes, fontPath);
        if (bytes.Length < 12)
            throw new NativeRuntimeException($"中文字体文件无效: {fontPath}");
        var magic = BinaryPrimitives.ReadUInt32BigEndian(bytes);
        if (magic == 0x4F54544F) // OTTO (CFF outlines)
            throw new NativeRuntimeException(
                $"PdfSharpCore 不支持 CFF/OpenType 轮廓字体: {fontPath}；"
                + "请改用 simhei.ttf / simsun.ttc / msyh.ttc");
        if (magic is not (0x00010000 or 0x74727565))
            throw new NativeRuntimeException($"中文字体不是有效的 TrueType 文件: {fontPath}");

        lock (Gate)
        {
            if (_resolver is null)
            {
                _resolver = new PdfFaceResolver(Faces);
                GlobalFontSettings.FontResolver = _resolver;
            }
            var name = $"office-native-cn-{_counter++}";
            Faces[name] = bytes;
            return name;
        }
    }

    private sealed class PdfFaceResolver(Dictionary<string, byte[]> faces) : IFontResolver
    {
        public string DefaultFontName =>
            faces.Count > 0 ? faces.Keys.First() : "Helvetica";

        public byte[] GetFont(string faceName) =>
            faces.TryGetValue(faceName, out var bytes)
                ? bytes
                : throw new InvalidOperationException($"未知字体: {faceName}");

        public FontResolverInfo ResolveTypeface(
            string familyName, bool bold, bool italic) =>
            new(familyName, bold, italic);
    }
}

/// <summary>Converts a TrueType collection (.ttc) into a stand-alone first face.</summary>
internal static class TrueTypeCollection
{
    public static byte[] ExtractFirstFace(byte[] data, string source)
    {
        if (data.Length < 16
            || Encoding.ASCII.GetString(data, 0, 4) != "ttcf")
            return data;

        try
        {
            var fontOffset = checked((int)BinaryPrimitives.ReadUInt32BigEndian(
                data.AsSpan(12)));
            if (fontOffset + 12 > data.Length)
                throw new InvalidDataException("TTC 字体表偏移越界");
            var numTables = BinaryPrimitives.ReadUInt16BigEndian(
                data.AsSpan(fontOffset + 4));
            var headerSize = 12 + numTables * 16;
            var records = new (uint Tag, int Offset, int Length)[numTables];
            var total = headerSize;
            for (var i = 0; i < numTables; i++)
            {
                var record = fontOffset + 12 + i * 16;
                if (record + 16 > data.Length)
                    throw new InvalidDataException("TTC 字体表记录越界");
                var offset = checked((int)BinaryPrimitives.ReadUInt32BigEndian(
                    data.AsSpan(record + 8)));
                var length = checked((int)BinaryPrimitives.ReadUInt32BigEndian(
                    data.AsSpan(record + 12)));
                if (offset < 0 || length < 0 || offset + length > data.Length)
                    throw new InvalidDataException("TTC 字体数据越界");
                records[i] = (
                    BinaryPrimitives.ReadUInt32BigEndian(data.AsSpan(record)),
                    offset, length);
                total = checked(total + ((length + 3) & ~3));
            }

            var output = new byte[total];
            data.AsSpan(fontOffset, 12).CopyTo(output);
            BinaryPrimitives.WriteUInt16BigEndian(output.AsSpan(4), numTables);
            var destination = headerSize;
            var headRecord = -1;
            for (var i = 0; i < numTables; i++)
            {
                var record = 12 + i * 16;
                BinaryPrimitives.WriteUInt32BigEndian(output.AsSpan(record), records[i].Tag);
                BinaryPrimitives.WriteUInt32BigEndian(
                    output.AsSpan(record + 8), (uint)destination);
                BinaryPrimitives.WriteUInt32BigEndian(
                    output.AsSpan(record + 12), (uint)records[i].Length);
                data.AsSpan(records[i].Offset, records[i].Length)
                    .CopyTo(output.AsSpan(destination));
                if (records[i].Tag == 0x68656164) // 'head'
                {
                    headRecord = record;
                    BinaryPrimitives.WriteUInt32BigEndian(
                        output.AsSpan(destination + 8), 0);
                }
                var checksum = Checksum(output.AsSpan(destination, records[i].Length));
                BinaryPrimitives.WriteUInt32BigEndian(output.AsSpan(record + 4), checksum);
                destination += (records[i].Length + 3) & ~3;
            }

            if (headRecord >= 0)
            {
                var headOffset = checked((int)BinaryPrimitives.ReadUInt32BigEndian(
                    output.AsSpan(headRecord + 8)));
                BinaryPrimitives.WriteUInt32BigEndian(output.AsSpan(headOffset + 8), 0);
                var whole = Checksum(output);
                BinaryPrimitives.WriteUInt32BigEndian(
                    output.AsSpan(headOffset + 8), 0xB1B0AFBA - whole);
            }
            return output;
        }
        catch (Exception ex) when (ex is InvalidDataException or OverflowException)
        {
            throw new NativeRuntimeException($"解析 TTC 中文字体失败（{source}）: {ex.Message}");
        }
    }

    private static uint Checksum(ReadOnlySpan<byte> data)
    {
        uint sum = 0;
        var index = 0;
        for (; index + 4 <= data.Length; index += 4)
            sum += BinaryPrimitives.ReadUInt32BigEndian(data[index..]);
        if (index < data.Length)
        {
            Span<byte> tail = stackalloc byte[4];
            data[index..].CopyTo(tail);
            sum += BinaryPrimitives.ReadUInt32BigEndian(tail);
        }
        return sum;
    }
}

/// <summary>Top-down page cursor and section renderers.</summary>
internal sealed class PdfCanvas : IDisposable
{
    private readonly PdfDocument _document;
    private readonly PdfPageSetup _setup;
    private readonly PdfFontConfig _fontConfig;
    private readonly string _faceName;
    private readonly Dictionary<(double Size, bool Bold, bool Italic), XFont> _fonts = [];
    private readonly Dictionary<string, XBrush> _brushes = new(StringComparer.Ordinal);
    private XGraphics? _graphics;

    public PdfCanvas(
        PdfDocument document, PdfPageSetup setup, PdfFontConfig fontConfig,
        string faceName)
    {
        _document = document;
        _setup = setup;
        _fontConfig = fontConfig;
        _faceName = faceName;
    }

    private double PageWidthPt => _setup.WidthMm * PdfPageSetup.MmToPt;
    private double PageHeightPt => _setup.HeightMm * PdfPageSetup.MmToPt;
    private double ContentLeft => _setup.MarginLeftMm * PdfPageSetup.MmToPt;
    private double ContentTop => _setup.MarginTopMm * PdfPageSetup.MmToPt;
    private double ContentWidth =>
        Math.Max(1, (_setup.WidthMm - _setup.MarginLeftMm - _setup.MarginRightMm)
                    * PdfPageSetup.MmToPt);
    private double ContentBottom =>
        PageHeightPt - _setup.MarginBottomMm * PdfPageSetup.MmToPt;
    private double CursorY { get; set; }

    public void NewPage()
    {
        _graphics?.Dispose();
        var page = _document.AddPage();
        page.Width = XUnit.FromPoint(PageWidthPt);
        page.Height = XUnit.FromPoint(PageHeightPt);
        _graphics = XGraphics.FromPdfPage(page);
        CursorY = ContentTop;
    }

    public void EndPage()
    {
        _graphics?.Dispose();
        _graphics = null;
    }

    public void Dispose() => EndPage();

    private XGraphics Graphics =>
        _graphics ?? throw new InvalidOperationException("PDF 页面尚未初始化");

    private XFont GetFont(double size, bool bold, bool italic)
    {
        var key = (Math.Round(size, 2), bold, italic);
        if (_fonts.TryGetValue(key, out var font))
            return font;
        var style = bold && italic
            ? XFontStyle.BoldItalic
            : bold
                ? XFontStyle.Bold
                : italic
                    ? XFontStyle.Italic
                    : XFontStyle.Regular;
        font = new XFont(_faceName, size, style, XPdfFontOptions.UnicodeDefault);
        _fonts[key] = font;
        return font;
    }

    private XBrush GetBrush(string? hex)
    {
        var key = string.IsNullOrWhiteSpace(hex) ? "" : hex.Trim();
        if (_brushes.TryGetValue(key, out var brush))
            return brush;
        var color = XBrushes.Black;
        if (key.Length > 0)
        {
            var text = key.StartsWith('#') ? key[1..] : key;
            if (text.Length == 6
                && int.TryParse(text, NumberStyles.HexNumber,
                    CultureInfo.InvariantCulture, out var rgb))
            {
                color = new XSolidBrush(XColor.FromArgb(
                    255, (rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF));
            }
        }
        _brushes[key] = color;
        return color;
    }

    private static double AscentPt(XFont font)
    {
        var ascent = font.Metrics.Ascent * font.Size / font.Metrics.UnitsPerEm;
        return ascent > 0 ? ascent : font.Size * 0.85;
    }

    private void EnsureSpace(double height)
    {
        if (CursorY + height <= ContentBottom + 0.01)
            return;
        if (CursorY <= ContentTop + 0.01)
            return;
        NewPage();
    }

    private void DrawAt(string text, XFont font, XBrush brush,
        double x, double baselineY) =>
        Graphics.DrawString(text, font, brush, x, baselineY);

    private void DrawAlignedLine(
        string text, XFont font, XBrush brush, double left, double width,
        double baselineY, string? alignment)
    {
        var textWidth = Graphics.MeasureString(text, font).Width;
        var x = alignment switch
        {
            "center" => left + Math.Max(0, (width - textWidth) / 2),
            "right" => left + Math.Max(0, width - textWidth),
            _ => left,
        };
        DrawAt(text, font, brush, x, baselineY);
    }

    public void RenderSection(PdfSection section)
    {
        switch (section.Type)
        {
            case "heading":
                RenderHeading(section);
                break;
            case "paragraph":
                RenderParagraph(section);
                break;
            case "table":
                RenderTable(section);
                break;
            case "image":
                RenderImage(section);
                break;
            case "page_break":
                NewPage();
                break;
            case "bullet_list":
                RenderList(section, ordered: false);
                break;
            case "numbered_list":
                RenderList(section, ordered: true);
                break;
        }
    }

    private void RenderHeading(PdfSection section)
    {
        var level = PdfJson.Clamp(section.Level ?? 1, 1, 6);
        var span = Math.Max(0, _fontConfig.HeadingSizePt - _fontConfig.BodySizePt);
        var step = span / 5.0;
        var size = Math.Round(_fontConfig.BodySizePt + step * (6 - level), 1);
        var style = section.Style;
        if (style?.FontSizePt is { } explicitSize)
            size = explicitSize;
        var font = GetFont(size, style?.Bold ?? true, style?.Italic ?? false);
        var brush = GetBrush(style?.Color);
        var lineHeight = size * 0.5 * PdfPageSetup.MmToPt;

        CursorY += 4 * PdfPageSetup.MmToPt;
        var lines = WrapText(section.Text ?? "", font, ContentWidth, 0);
        EnsureSpace(lineHeight + 2 * PdfPageSetup.MmToPt);
        foreach (var line in lines)
        {
            EnsureSpace(lineHeight);
            DrawAlignedLine(line, font, brush, ContentLeft, ContentWidth,
                CursorY + AscentPt(font), style?.Alignment ?? "left");
            CursorY += lineHeight;
        }
        CursorY += 2 * PdfPageSetup.MmToPt;
    }

    private void RenderParagraph(PdfSection section)
    {
        var text = section.Text ?? "";
        if (string.IsNullOrWhiteSpace(text))
        {
            CursorY += 4 * PdfPageSetup.MmToPt;
            return;
        }
        var style = section.Style;
        var size = style?.FontSizePt ?? _fontConfig.BodySizePt;
        var font = GetFont(size, style?.Bold ?? false, style?.Italic ?? false);
        var brush = GetBrush(style?.Color);
        var lineHeight = Math.Max(1, size * 0.6 * PdfPageSetup.MmToPt);
        var indent = Math.Max(0, style?.FirstLineIndentPt ?? 0);
        var lines = WrapText(text, font, ContentWidth, indent);
        for (var index = 0; index < lines.Count; index++)
        {
            EnsureSpace(lineHeight);
            var left = ContentLeft + (index == 0 ? indent : 0);
            var width = Math.Max(1, ContentWidth - (index == 0 ? indent : 0));
            DrawAlignedLine(lines[index], font, brush, left, width,
                CursorY + AscentPt(font), style?.Alignment ?? "justify");
            CursorY += lineHeight;
        }
    }

    private void RenderTable(PdfSection section)
    {
        var table = section.Table;
        if (table is null)
            return;
        var headers = table.Headers;
        var rows = table.Rows;
        if (headers.Count == 0 && rows.Count == 0)
            return;

        CursorY += 2 * PdfPageSetup.MmToPt;

        if (headers.Count == 0)
        {
            var columns = rows[0].Count;
            headers = Enumerable.Range(1, columns).Select(i => $"Col{i}").ToList();
        }

        var columnCount = headers.Count;
        if (columnCount == 0)
            return;
        var widths = new double[columnCount];
        if (table.ColumnWidths is { Count: > 0 } ratios
            && ratios.Count == columnCount)
        {
            for (var i = 0; i < columnCount; i++)
                widths[i] = ContentWidth * ratios[i];
        }
        else
        {
            for (var i = 0; i < columnCount; i++)
                widths[i] = ContentWidth / columnCount;
        }
        var totalWidth = widths.Sum();
        if (totalWidth > ContentWidth && totalWidth > 0)
        {
            var factor = ContentWidth / totalWidth;
            for (var i = 0; i < columnCount; i++)
                widths[i] *= factor;
        }

        var headerSize = Math.Max(6, _fontConfig.BodySizePt - 1);
        var rowSize = Math.Max(6, _fontConfig.BodySizePt - 1);
        var headerFont = GetFont(headerSize, bold: true, italic: false);
        var rowFont = GetFont(rowSize, bold: false, italic: false);
        var headerLineHeight = Math.Max(10, headerSize * 0.7 * PdfPageSetup.MmToPt);
        var rowLineHeight = Math.Max(10, rowSize * 0.6 * PdfPageSetup.MmToPt);
        const double padding = 2.0;

        var headerLines = headers
            .Select((header, index) => WrapText(header, headerFont,
                Math.Max(1, widths[index] - padding * 2), 0))
            .ToList();
        var headerRows = headerLines.Max(lines => lines.Count);
        var headerHeight = Math.Max(headerLineHeight, headerRows * headerLineHeight + padding * 2);

        var needsHeader = headers.Count > 0;
        void DrawHeader()
        {
            if (!needsHeader)
                return;
            var left = ContentLeft;
            for (var column = 0; column < columnCount; column++)
            {
                var rect = new XRect(left, CursorY, widths[column], headerHeight);
                Graphics.DrawRectangle(
                    new XSolidBrush(XColor.FromArgb(230, 230, 230)), rect);
                Graphics.DrawRectangle(XPens.Black, rect);
                var baseline = CursorY + padding + AscentPt(headerFont);
                foreach (var line in headerLines[column])
                {
                    var textWidth = Graphics.MeasureString(line, headerFont).Width;
                    var x = left + (widths[column] - textWidth) / 2;
                    DrawAt(line, headerFont, XBrushes.Black,
                        Math.Max(left + padding, x), baseline);
                    baseline += headerLineHeight;
                }
                left += widths[column];
            }
            CursorY += headerHeight;
        }

        EnsureSpace(headerHeight + rowLineHeight);
        DrawHeader();

        foreach (var row in rows)
        {
            var cellTexts = new string[columnCount];
            for (var column = 0; column < columnCount; column++)
                cellTexts[column] = column < row.Count ? row[column] : "";
            var cellLines = new List<List<string>>(columnCount);
            for (var column = 0; column < columnCount; column++)
                cellLines.Add(WrapText(cellTexts[column], rowFont,
                    Math.Max(1, widths[column] - padding * 2), 0));
            var lineCount = Math.Max(1, cellLines.Max(lines => lines.Count));
            var rowHeight = Math.Max(rowLineHeight, lineCount * rowLineHeight + padding * 2);

            if (CursorY + rowHeight > ContentBottom + 0.01
                && CursorY > ContentTop + 0.01)
            {
                NewPage();
                DrawHeader();
            }

            var left = ContentLeft;
            for (var column = 0; column < columnCount; column++)
            {
                var rect = new XRect(left, CursorY, widths[column], rowHeight);
                Graphics.DrawRectangle(XPens.Black, rect);
                var brush = XBrushes.Black;
                var baseline = CursorY + padding + AscentPt(rowFont);
                foreach (var line in cellLines[column])
                {
                    var textWidth = Graphics.MeasureString(line, rowFont).Width;
                    var x = column == 0
                        ? left + padding
                        : left + (widths[column] - textWidth) / 2;
                    DrawAt(line, rowFont, brush,
                        Math.Max(left + padding, x), baseline);
                    baseline += rowLineHeight;
                }
                left += widths[column];
            }
            CursorY += rowHeight;
        }

        CursorY += 2 * PdfPageSetup.MmToPt;
    }

    private void RenderImage(PdfSection section)
    {
        var imageDef = section.Image;
        if (imageDef is null || string.IsNullOrWhiteSpace(imageDef.Path))
            return;

        CursorY += 2 * PdfPageSetup.MmToPt;
        if (!File.Exists(imageDef.Path))
        {
            RenderPlaceholder($"_ [图片未找到: {imageDef.Path}] _");
            return;
        }

        XImage image;
        try
        {
            image = XImage.FromFile(imageDef.Path);
        }
        catch (Exception ex)
        {
            RenderPlaceholder($"_ [图片加载失败: {imageDef.Path}（{ex.Message}）] _");
            return;
        }

        using (image)
        {
            var naturalWidth = image.PointWidth > 0
                ? image.PointWidth
                : image.PixelWidth * 72.0 / 96.0;
            var naturalHeight = image.PointHeight > 0
                ? image.PointHeight
                : image.PixelHeight * 72.0 / 96.0;
            if (naturalWidth <= 0 || naturalHeight <= 0)
            {
                RenderPlaceholder($"_ [图片尺寸无效: {imageDef.Path}] _");
                return;
            }

            double width;
            double height;
            if (imageDef.WidthInches is { } widthInches
                && imageDef.HeightInches is { } heightInches)
            {
                width = widthInches * 72;
                height = heightInches * 72;
            }
            else if (imageDef.WidthInches is { } onlyWidth)
            {
                width = onlyWidth * 72;
                height = naturalHeight * width / naturalWidth;
            }
            else if (imageDef.HeightInches is { } onlyHeight)
            {
                height = onlyHeight * 72;
                width = naturalWidth * height / naturalHeight;
            }
            else
            {
                width = ContentWidth;
                height = naturalHeight * width / naturalWidth;
            }

            if (width > ContentWidth)
            {
                var factor = ContentWidth / width;
                width *= factor;
                height *= factor;
            }
            if (height > ContentBottom - ContentTop)
            {
                var factor = (ContentBottom - ContentTop) / height;
                width *= factor;
                height *= factor;
            }

            EnsureSpace(height + 2 * PdfPageSetup.MmToPt);
            Graphics.DrawImage(image, ContentLeft, CursorY, width, height);
            CursorY += height + 2 * PdfPageSetup.MmToPt;
        }
    }

    private void RenderPlaceholder(string text)
    {
        var size = _fontConfig.BodySizePt;
        var font = GetFont(size, bold: false, italic: true);
        var lineHeight = Math.Max(1, size * 0.6 * PdfPageSetup.MmToPt);
        foreach (var line in WrapText(text, font, ContentWidth, 0))
        {
            EnsureSpace(lineHeight);
            DrawAt(line, font, XBrushes.Black, ContentLeft, CursorY + AscentPt(font));
            CursorY += lineHeight;
        }
        CursorY += 2 * PdfPageSetup.MmToPt;
    }

    private void RenderList(PdfSection section, bool ordered)
    {
        var items = section.Items;
        if (items is null || items.Count == 0)
            return;

        var style = section.Style;
        var size = style?.FontSizePt ?? _fontConfig.BodySizePt;
        var font = GetFont(size, style?.Bold ?? false, style?.Italic ?? false);
        var brush = GetBrush(style?.Color);
        var lineHeight = Math.Max(1, size * 0.55 * PdfPageSetup.MmToPt);
        var indent = 5 * PdfPageSetup.MmToPt;
        var available = Math.Max(1, ContentWidth - indent);

        CursorY += 2 * PdfPageSetup.MmToPt;
        for (var index = 0; index < items.Count; index++)
        {
            var prefix = ordered ? $"{index + 1}. " : "- ";
            var lines = WrapText(prefix + items[index], font, available, 0);
            foreach (var line in lines)
            {
                EnsureSpace(lineHeight);
                DrawAt(line, font, brush, ContentLeft + indent,
                    CursorY + AscentPt(font));
                CursorY += lineHeight;
            }
        }
        CursorY += 2 * PdfPageSetup.MmToPt;
    }

    // ── text wrapping ──────────────────────────────────────────────

    private List<string> WrapText(
        string text, XFont font, double maxWidth, double firstLineIndent)
    {
        var lines = new List<string>();
        if (maxWidth <= 1)
        {
            lines.Add(text);
            return lines;
        }

        var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        var explicitLines = normalized.Split('\n');
        var firstOutputLine = true;
        foreach (var explicitLine in explicitLines)
        {
            if (explicitLine.Length == 0)
            {
                lines.Add("");
                firstOutputLine = false;
                continue;
            }
            var limit = Math.Max(1, maxWidth - (firstOutputLine ? firstLineIndent : 0));
            var current = new StringBuilder();
            var currentWidth = 0.0;
            var units = Tokenize(explicitLine);
            var unitIndex = 0;
            while (unitIndex < units.Count)
            {
                var unit = units[unitIndex];
                var unitWidth = Measure(unit.Text, font);
                if (unit.IsWhitespace)
                {
                    if (current.Length == 0)
                    {
                        unitIndex++;
                        continue;
                    }
                    if (currentWidth + unitWidth > limit + 0.01)
                    {
                        FlushLine();
                        unitIndex++;
                        continue;
                    }
                    current.Append(unit.Text);
                    currentWidth += unitWidth;
                    unitIndex++;
                    continue;
                }

                if (currentWidth + unitWidth <= limit + 0.01)
                {
                    current.Append(unit.Text);
                    currentWidth += unitWidth;
                    unitIndex++;
                    continue;
                }

                if (current.Length > 0)
                {
                    FlushLine();
                    continue;
                }

                // A single anchor unit (Latin word) wider than the line: split it.
                var emitted = false;
                foreach (var piece in SplitOversized(unit.Text, font, limit))
                {
                    if (piece.Length == 0)
                        continue;
                    current.Append(piece);
                    currentWidth = Measure(piece, font);
                    emitted = true;
                    if (currentWidth >= limit - 0.01)
                    {
                        FlushLine();
                        emitted = false;
                    }
                }
                if (!emitted && current.Length > 0)
                    FlushLine();
                unitIndex++;
            }
            if (current.Length > 0 || lines.Count == 0)
                FlushLine();
            firstOutputLine = false;

            void FlushLine()
            {
                lines.Add(current.ToString().TrimEnd());
                current.Clear();
                currentWidth = 0;
                limit = maxWidth;
            }
        }

        if (lines.Count == 0)
            lines.Add("");
        return lines;
    }

    private double Measure(string text, XFont font) =>
        text.Length == 0 ? 0 : Graphics.MeasureString(text, font).Width;

    private static List<(string Text, bool IsWhitespace)> Tokenize(string text)
    {
        var units = new List<(string Text, bool IsWhitespace)>();
        var index = 0;
        while (index < text.Length)
        {
            var c = text[index];
            if (char.IsWhiteSpace(c))
            {
                var start = index;
                while (index < text.Length && char.IsWhiteSpace(text[index]))
                    index++;
                units.Add((text[start..index], true));
                continue;
            }
            if (IsCjk(c))
            {
                var start = index;
                index++;
                if (char.IsHighSurrogate(c) && index < text.Length)
                    index++;
                units.Add((text[start..index], false));
                continue;
            }
            var wordStart = index;
            while (index < text.Length)
            {
                var current = text[index];
                if (char.IsWhiteSpace(current) || IsCjk(current))
                    break;
                index++;
            }
            units.Add((text[wordStart..index], false));
        }
        return units;
    }

    private IEnumerable<string> SplitOversized(
        string text, XFont font, double limit)
    {
        // Split on text elements; the caller only needs forward progress.
        var start = 0;
        var width = 0.0;
        for (var index = 0; index < text.Length; index++)
        {
            var length = char.IsHighSurrogate(text[index]) && index + 1 < text.Length
                ? 2
                : 1;
            var piece = text.Substring(index, length);
            var pieceWidth = Measure(piece, font);
            if (width + pieceWidth > limit && index > start)
            {
                yield return text[start..index];
                start = index;
                width = 0;
            }
            width += pieceWidth;
            index += length - 1;
        }
        if (start < text.Length)
            yield return text[start..];
    }

    private static bool IsCjk(char c) =>
        (c >= 0x2E80 && c <= 0x9FFF)
        || (c >= 0xAC00 && c <= 0xD7AF)
        || (c >= 0xF900 && c <= 0xFAFF)
        || (c >= 0xFE30 && c <= 0xFE4F)
        || (c >= 0xFF00 && c <= 0xFFEF);
}
