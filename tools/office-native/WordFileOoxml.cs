using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeNative;

// CoreFilePropertiesPart.CoreFileProperties is the only SDK surface for /docProps/core.xml.
#pragma warning disable OOXML0001

/// <summary>Formatting values collected while building one paragraph or run.</summary>
internal sealed class WordRunFormat
{
    public string? FontName { get; set; }
    public double? SizePt { get; set; }
    public bool? Bold { get; set; }
    public bool? Italic { get; set; }
    public string? Color { get; set; }
    public bool Underline { get; set; }
}

internal sealed class WordParagraphFormat
{
    public string? StyleId { get; set; }
    public int? NumberingId { get; set; }
    public string? ShadingFill { get; set; }
    public bool ShadingIsNilPattern { get; set; }
    public int? OutlineLevel { get; set; }
    public string? Alignment { get; set; }
    public double? LineSpacing { get; set; }
    public int? FirstLineIndentPt { get; set; }
    public int? SpaceBeforePt { get; set; }
    public int? SpaceAfterPt { get; set; }
    public bool ContextualSpacing { get; set; }
    public bool BottomBorder { get; set; }
    public string BorderColor { get; set; } = "000000";
}

internal static class WordFileOoxml
{
    public const double TwipsPerInch = 1440.0;
    public const double TwipsPerPoint = 20.0;
    public const long EmuPerInch = 914400L;
    public const int MaxRunTextChars = 16000;

    private static readonly Dictionary<string, double> CnFontSizes = new(StringComparer.Ordinal)
    {
        ["初号"] = 42.0,
        ["小初"] = 36.0,
        ["一号"] = 26.0,
        ["小一"] = 24.0,
        ["二号"] = 22.0,
        ["小二"] = 18.0,
        ["三号"] = 16.0,
        ["小三"] = 15.0,
        ["四号"] = 14.0,
        ["小四"] = 12.0,
        ["五号"] = 10.5,
        ["小五"] = 9.0,
        ["六号"] = 7.5,
        ["小六"] = 6.5,
        ["七号"] = 5.5,
        ["八号"] = 5.0,
    };

    private static readonly Dictionary<string, (uint Width, uint Height)> PaperSizes =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["A4"] = (11909, 16834),
            ["Letter"] = (12240, 15840),
            ["A3"] = (16834, 23818),
            ["B5"] = (10325, 14630),
        };

    private static readonly string[] SectionChildOrder =
    [
        "headerReference", "footerReference", "footnotePr", "endnotePr", "type",
        "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols",
        "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi",
        "rtlGutter", "docGrid", "printerSettings", "sectPrChange",
    ];

    public static double CnSizeToPt(JsonElement value)
    {
        switch (value.ValueKind)
        {
            case JsonValueKind.Number when value.TryGetDouble(out var number):
                return number;
            case JsonValueKind.String:
                return CnSizeToPt(value.GetString() ?? "");
            default:
                throw new NativeUsageException($"字号必须是中文标准字号或 pt 数值: {value}");
        }
    }

    public static double CnSizeToPt(string raw)
    {
        var text = raw.Trim();
        if (CnFontSizes.TryGetValue(text, out var mapped))
            return mapped;
        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture,
                out var parsed))
            return parsed;
        throw new NativeUsageException(
            $"未知的中文字号: {text}，可选值: {string.Join(", ", CnFontSizes.Keys)}");
    }

    public static string Twips(double twips) =>
        ((int)Math.Round(twips, MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture);

    public static string PtToTwips(double points) => Twips(points * TwipsPerPoint);

    public static string InchesToTwips(double inches) => Twips(inches * TwipsPerInch);

    // ── run / paragraph formatting ────────────────────────────────

    public static RunProperties BuildRunProperties(WordRunFormat? format)
    {
        var properties = new RunProperties();
        if (format is null)
            return properties;
        if (!string.IsNullOrEmpty(format.FontName))
        {
            properties.Append(new RunFonts
            {
                Ascii = format.FontName,
                HighAnsi = format.FontName,
                EastAsia = format.FontName,
            });
        }
        if (format.Bold == true)
            properties.Append(new Bold());
        else if (format.Bold == false)
            properties.Append(new Bold { Val = false });
        if (format.Italic == true)
            properties.Append(new Italic());
        else if (format.Italic == false)
            properties.Append(new Italic { Val = false });
        if (!string.IsNullOrEmpty(format.Color))
            properties.Append(new Color { Val = format.Color });
        if (format.SizePt is { } size)
        {
            var halfPoints = ((int)Math.Round(size * 2, MidpointRounding.AwayFromZero))
                .ToString(CultureInfo.InvariantCulture);
            properties.Append(new FontSize { Val = halfPoints });
            properties.Append(new FontSizeComplexScript { Val = halfPoints });
        }
        if (format.Underline)
            properties.Append(new Underline { Val = UnderlineValues.Single });
        return properties;
    }

    public static Run NewTextRun(string text, WordRunFormat? format)
    {
        var run = new Run();
        var properties = BuildRunProperties(format);
        if (properties.HasChildren)
            run.Append(properties);
        var buffer = new StringBuilder();
        void FlushText()
        {
            if (buffer.Length == 0)
                return;
            run.Append(new Text(buffer.ToString())
            {
                Space = SpaceProcessingModeValues.Preserve,
            });
            buffer.Clear();
        }
        for (var index = 0; index < text.Length; index++)
        {
            var character = text[index];
            if (character == '\t')
            {
                FlushText();
                run.Append(new TabChar());
            }
            else if (character is '\r' or '\n')
            {
                FlushText();
                if (character == '\r' && index + 1 < text.Length && text[index + 1] == '\n')
                    index++;
                run.Append(new Break());
            }
            else
            {
                buffer.Append(character);
            }
        }
        FlushText();
        return run;
    }

    public static ParagraphProperties BuildParagraphProperties(WordParagraphFormat? format)
    {
        var properties = new ParagraphProperties();
        if (format is null)
            return properties;
        if (!string.IsNullOrEmpty(format.StyleId))
            properties.Append(new ParagraphStyleId { Val = format.StyleId });
        if (format.NumberingId is { } numberingId)
        {
            properties.Append(new NumberingProperties(
                new NumberingLevelReference { Val = 0 },
                new NumberingId { Val = numberingId }));
        }
        if (format.BottomBorder)
        {
            properties.Append(new ParagraphBorders(
                new BottomBorder
                {
                    Val = BorderValues.Single,
                    Size = 12,
                    Space = 1,
                    Color = format.BorderColor,
                }));
        }
        if (!string.IsNullOrEmpty(format.ShadingFill))
        {
            properties.Append(new Shading
            {
                Val = format.ShadingIsNilPattern
                    ? ShadingPatternValues.Nil
                    : ShadingPatternValues.Clear,
                Fill = format.ShadingFill,
            });
        }
        if (format.ContextualSpacing)
            properties.Append(new ContextualSpacing());
        if (format.LineSpacing is { } lineSpacing
            || format.SpaceBeforePt is not null
            || format.SpaceAfterPt is not null)
        {
            var spacing = new SpacingBetweenLines();
            if (format.SpaceBeforePt is { } before)
                spacing.Before = PtToTwips(before);
            if (format.SpaceAfterPt is { } after)
                spacing.After = PtToTwips(after);
            if (format.LineSpacing is { } lineSpacingValue)
            {
                var line = (int)Math.Round(lineSpacingValue * 240,
                    MidpointRounding.AwayFromZero);
                spacing.Line = line.ToString(CultureInfo.InvariantCulture);
                spacing.LineRule = LineSpacingRuleValues.Auto;
            }
            properties.Append(spacing);
        }
        if (format.FirstLineIndentPt is { } indent)
        {
            var twips = PtToTwips(Math.Abs(indent));
            properties.Append(indent >= 0
                ? new Indentation { FirstLine = twips }
                : new Indentation { Hanging = twips });
        }
        if (!string.IsNullOrEmpty(format.Alignment))
        {
            properties.Append(new Justification
            {
                Val = format.Alignment switch
                {
                    "center" => JustificationValues.Center,
                    "right" => JustificationValues.Right,
                    "justify" => JustificationValues.Both,
                    _ => JustificationValues.Left,
                },
            });
        }
        if (format.OutlineLevel is { } outlineLevel)
            properties.Append(new OutlineLevel { Val = outlineLevel });
        return properties;
    }

    public static IEnumerable<string> SplitLongText(string? text)
    {
        if (string.IsNullOrEmpty(text))
            yield break;
        for (var start = 0; start < text.Length; start += MaxRunTextChars)
        {
            var length = Math.Min(MaxRunTextChars, text.Length - start);
            yield return text.Substring(start, length);
        }
    }

    // ── field code helpers ────────────────────────────────────────

    public static string FieldCodeFor(FieldDef field)
    {
        var fieldType = field.FieldType;
        if (fieldType == "CUSTOM")
        {
            if (string.IsNullOrWhiteSpace(field.Code))
                throw new NativeUsageException("CUSTOM 域必须提供 field.code");
            return field.Code!;
        }
        return fieldType switch
        {
            "PAGE" or "NUMPAGES" or "SECTIONPAGES" =>
                string.IsNullOrEmpty(field.NumberFormat)
                    ? $" {fieldType} "
                    : $" {fieldType} \\* {field.NumberFormat} ",
            "DATE" => " DATE \\@ \"yyyy年M月d日\" ",
            "TIME" => " TIME \\@ \"HH:mm\" ",
            _ => $" {fieldType} ",
        };
    }

    public static Paragraph BuildFieldParagraph(FieldDef field, WordParagraphFormat? format)
    {
        var paragraph = new Paragraph();
        var properties = BuildParagraphProperties(format);
        if (properties.HasChildren)
            paragraph.Append(properties);
        paragraph.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin, Dirty = true }));
        paragraph.Append(new Run(new FieldCode(FieldCodeFor(field))
        {
            Space = SpaceProcessingModeValues.Preserve,
        }));
        paragraph.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
        paragraph.Append(new Run(new Text(field.FallbackText ?? " ")
        {
            Space = SpaceProcessingModeValues.Preserve,
        }));
        paragraph.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        return paragraph;
    }

    public static Paragraph BuildTocFieldParagraph(string levels)
    {
        return new Paragraph(
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin, Dirty = true }),
            new Run(new FieldCode($" TOC \\o \"{levels}\" \\h \\z \\u ")
            {
                Space = SpaceProcessingModeValues.Preserve,
            }),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
    }

    public static Run BuildPageFieldRuns(string instruction)
    {
        return new Run(
            new FieldChar { FieldCharType = FieldCharValues.Begin },
            new FieldCode(instruction) { Space = SpaceProcessingModeValues.Preserve },
            new FieldChar { FieldCharType = FieldCharValues.End });
    }

    // ── body / section helpers ───────────────────────────────────

    public static void AppendBodyElement(Body body, OpenXmlElement element)
    {
        var section = body.Elements<SectionProperties>().LastOrDefault();
        if (section is not null)
            body.InsertBefore(element, section);
        else
            body.Append(element);
    }

    public static SectionProperties EnsureBodySection(Body body)
    {
        var section = body.Elements<SectionProperties>().LastOrDefault();
        if (section is not null)
            return section;
        section = DefaultSection();
        body.Append(section);
        return section;
    }

    /// <summary>
    /// python-docx Document.add_section semantics for the sentinel (body-level)
    /// section: it becomes the new last section, starts on a new page, and drops
    /// explicit header/footer references so it inherits them from the previous
    /// section.  The clone placed in the preceding paragraph keeps the old refs.
    /// </summary>
    public static void PrepareNewLastSection(SectionProperties section)
    {
        foreach (var reference in section.Elements<HeaderReference>().ToList())
            reference.Remove();
        foreach (var reference in section.Elements<FooterReference>().ToList())
            reference.Remove();
        GetOrAddSectionChild(section, () => new SectionType()).Val =
            SectionMarkValues.NextPage;
    }

    public static SectionProperties DefaultSection() => new(
        new PageSize { Width = 12240U, Height = 15840U },
        new PageMargin
        {
            Top = 1440,
            Right = 1800U,
            Bottom = 1440,
            Left = 1800U,
            Header = 720U,
            Footer = 720U,
            Gutter = 0U,
        },
        new Columns { Space = "720" },
        new DocGrid { LinePitch = 360 });

    public static void ApplyPageSetup(Body body, PageSetup setup)
    {
        foreach (var section in body.Descendants<SectionProperties>())
        {
            var pageSize = GetOrAddSectionChild(section, () => new PageSize());
            if (setup.PaperSize is { } paper
                && PaperSizes.TryGetValue(paper, out var dimensions))
            {
                pageSize.Width = dimensions.Width;
                pageSize.Height = dimensions.Height;
            }
            if (setup.Orientation is { } orientation)
            {
                pageSize.Orient = orientation == "landscape"
                    ? PageOrientationValues.Landscape
                    : PageOrientationValues.Portrait;
                if (orientation == "landscape")
                {
                    var width = pageSize.Width?.Value ?? 0U;
                    var height = pageSize.Height?.Value ?? 0U;
                    pageSize.Width = height;
                    pageSize.Height = width;
                }
            }
            if (setup.MarginTop is not null || setup.MarginBottom is not null
                || setup.MarginLeft is not null || setup.MarginRight is not null)
            {
                var margin = GetOrAddSectionChild(section, () => new PageMargin
                {
                    Top = 1440,
                    Right = 1800U,
                    Bottom = 1440,
                    Left = 1800U,
                    Header = 720U,
                    Footer = 720U,
                    Gutter = 0U,
                });
                if (setup.MarginTop is { } top)
                    margin.Top = (int)Math.Round(top * TwipsPerInch);
                if (setup.MarginBottom is { } bottom)
                    margin.Bottom = (int)Math.Round(bottom * TwipsPerInch);
                if (setup.MarginLeft is { } left)
                    margin.Left = (uint)Math.Round(left * TwipsPerInch);
                if (setup.MarginRight is { } right)
                    margin.Right = (uint)Math.Round(right * TwipsPerInch);
            }
            ApplyColumns(section, setup.Columns, setup.ColumnSpacing);
            ApplyPageBorder(section, setup.PageBorder, setup.PageBorderColor);
        }
    }

    // ── columns / page borders / watermark ───────────────────────

    private static readonly Dictionary<string, (BorderValues Value, uint Size)> PageBorderStyles =
        new(StringComparer.Ordinal)
        {
            ["single"] = (BorderValues.Single, 4U),
            ["double"] = (BorderValues.Double, 6U),
            ["dashed"] = (BorderValues.Dashed, 4U),
            ["dotted"] = (BorderValues.Dotted, 4U),
        };

    public static void ApplyColumns(SectionProperties section, int? columns, double? spacing)
    {
        if (columns is null && spacing is null)
            return;
        if (columns is < 1 or > 4)
            throw new NativeUsageException("columns 必须是 1-4 的整数");
        if (spacing is <= 0)
            throw new NativeUsageException("column_spacing 必须大于 0（英寸）");
        var element = GetOrAddSectionChild(section, () => new Columns());
        if (columns is { } count)
            element.ColumnCount = (short)count;
        if (spacing is { } inches)
            element.Space = InchesToTwips(inches);
    }

    public static void ApplyPageBorder(
        SectionProperties section, string? style, string? color)
    {
        if (style is null && color is null)
            return;
        var existing = section.GetFirstChild<PageBorders>();
        if (style is null)
        {
            // 仅给颜色：沿用已有边框线型，避免生成缺少 w:val 的非法边框
            if (existing is null)
                return;
            foreach (var edge in PageBorderEdges(existing))
                edge.Color = color;
            return;
        }
        if (!PageBorderStyles.TryGetValue(style, out var borderStyle))
            throw new NativeUsageException(
                "page_border 必须是 single/double/dashed/dotted 之一");
        var borders = existing is null || PageBorderEdges(existing).Count() != 4
            ? CreatePageBorders(section)
            : existing;
        borders.OffsetFrom = PageBorderOffsetValues.Page;
        var borderColor = color ?? "auto";
        foreach (var edge in PageBorderEdges(borders))
        {
            edge.Val = borderStyle.Value;
            edge.Size = borderStyle.Size;
            edge.Space = 24U;
            edge.Color = borderColor;
        }
    }

    private static PageBorders CreatePageBorders(SectionProperties section)
    {
        // 插入位置沿用既有 pgBorders（替换场景），否则按 schema 顺序定位
        var existing = section.GetFirstChild<PageBorders>();
        var insertAt = -1;
        if (existing is not null)
        {
            insertAt = section.ChildElements.ToList().IndexOf(existing);
            existing.Remove();
        }
        var borders = new PageBorders(
            new TopBorder(),
            new LeftBorder(),
            new BottomBorder(),
            new RightBorder())
        {
            OffsetFrom = PageBorderOffsetValues.Page,
        };
        if (insertAt >= 0)
        {
            section.InsertAt(borders, insertAt);
            return borders;
        }
        var order = Array.IndexOf(SectionChildOrder, "pgBorders");
        var before = section.ChildElements.FirstOrDefault(
            child => Array.IndexOf(SectionChildOrder, child.LocalName) > order);
        if (before is not null)
            section.InsertBefore(borders, before);
        else
            section.Append(borders);
        return borders;
    }

    private static IEnumerable<BorderType> PageBorderEdges(PageBorders borders)
    {
        BorderType?[] slots =
        [
            borders.TopBorder, borders.LeftBorder, borders.BottomBorder, borders.RightBorder,
        ];
        foreach (var slot in slots)
        {
            if (slot is not null)
                yield return slot;
        }
    }

    /// <summary>
    /// Writes a VML text watermark into the default header of every section.
    /// The SDK does not model VML, so the pict/shape payload is injected as raw XML.
    /// </summary>
    public static void ApplyWatermark(MainDocumentPart main, string? text, string? color)
    {
        if (text is null && color is null)
            return;
        if (string.IsNullOrWhiteSpace(text))
            throw new NativeUsageException("watermark_text 不能为空");
        var body = main.Document?.Body;
        if (body is null)
            return;
        var fillColor = color ?? "D9D9D9";
        var seenParts = new HashSet<string>(StringComparer.Ordinal);
        var sections = body.Descendants<SectionProperties>().ToList();
        for (var index = 0; index < sections.Count; index++)
        {
            var section = sections[index];
            var reference = section.Elements<HeaderReference>()
                .FirstOrDefault(item => item.Type?.Value == HeaderFooterValues.Default);
            HeaderPart header;
            if (reference?.Id?.Value is { } relationshipId
                && main.GetPartById(relationshipId) is HeaderPart existing)
            {
                header = existing;
            }
            else if (index == 0)
            {
                // 首个节没有页眉：创建默认页眉；后续节保持“链接到前一节”的继承关系
                header = main.AddNewPart<HeaderPart>();
                header.Header = new Header(new Paragraph(new ParagraphProperties(
                    new ParagraphStyleId { Val = "Header" })));
                AddHeaderReference(section, main.GetIdOfPart(header));
            }
            else
            {
                continue;
            }
            if (!seenParts.Add(main.GetIdOfPart(header)))
                continue;
            header.Header ??= new Header();
            RemoveWatermarkRuns(header.Header);
            var paragraph = header.Header.Elements<Paragraph>().FirstOrDefault();
            if (paragraph is null)
            {
                paragraph = new Paragraph(new ParagraphProperties(
                    new ParagraphStyleId { Val = "Header" }));
                header.Header.Append(paragraph);
            }
            paragraph.Append(ParseWatermarkRun(text!, fillColor));
            header.Header.Save();
        }
    }

    private static void RemoveWatermarkRuns(Header header)
    {
        foreach (var paragraph in header.Elements<Paragraph>().ToList())
        {
            var removed = false;
            foreach (var run in paragraph.Elements<Run>().ToList())
            {
                if (!IsWatermarkRun(run))
                    continue;
                run.Remove();
                removed = true;
            }
            if (removed && !paragraph.ChildElements.Any(child => child is not ParagraphProperties)
                && header.Elements<Paragraph>().Count() > 1)
                paragraph.Remove();
        }
    }

    private static bool IsWatermarkRun(Run run) =>
        run.Descendants().Any(element => element.LocalName == "shape"
            && (element.GetAttributes()
                    .FirstOrDefault(attribute => attribute.LocalName == "type").Value ?? "")
                .Contains("_x0000_t136", StringComparison.Ordinal));

    private static Run ParseWatermarkRun(string text, string fillColor)
    {
        var style =
            "position:absolute;margin-left:0;margin-top:0;width:500pt;height:120pt;"
            + "rotation:-45;z-index:-251658752;mso-position-horizontal:center;"
            + "mso-position-horizontal-relative:margin;mso-position-vertical:center;"
            + "mso-position-vertical-relative:margin";
        var xml =
            "<w:r xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
            + " xmlns:v=\"urn:schemas-microsoft-com:vml\""
            + " xmlns:o=\"urn:schemas-microsoft-com:office:office\">"
            + "<w:rPr><w:noProof/></w:rPr><w:pict>"
            + WatermarkShapeTypeXml
            + "<v:shape id=\"PowerPlusWaterMarkObject\" o:spid=\"_x0000_s2049\""
            + " type=\"#_x0000_t136\""
            + $" style=\"{style}\" o:allowincell=\"f\" fillcolor=\"#{fillColor}\" stroked=\"f\">"
            + "<v:fill opacity=\".5\"/>"
            + "<v:textpath style=\"font-family:&quot;宋体&quot;;font-size:1in\""
            + $" string=\"{XmlEscapeAttribute(text)}\"/>"
            + "</v:shape></w:pict></w:r>";
        return new Run(xml);
    }

    private const string WatermarkShapeTypeXml =
        "<v:shapetype id=\"_x0000_t136\" coordsize=\"21600,21600\" o:spt=\"136\""
        + " adj=\"10800\" path=\"m@7,l@8,m@5,21600l@6,21600e\">"
        + "<v:formulas>"
        + "<v:f eqn=\"sum #0 0 10800\"/><v:f eqn=\"prod #0 2 1\"/>"
        + "<v:f eqn=\"sum 21600 0 @1\"/><v:f eqn=\"sum 0 0 @2\"/>"
        + "<v:f eqn=\"sum 21600 0 @3\"/><v:f eqn=\"if @0 @3 0\"/>"
        + "<v:f eqn=\"if @0 21600 @1\"/><v:f eqn=\"if @0 0 @2\"/>"
        + "<v:f eqn=\"if @0 @4 21600\"/><v:f eqn=\"mid @5 @6\"/>"
        + "<v:f eqn=\"mid @8 @5\"/><v:f eqn=\"mid @7 @8\"/>"
        + "<v:f eqn=\"mid @6 @7\"/><v:f eqn=\"sum @6 0 @5\"/>"
        + "</v:formulas>"
        + "<v:path textpathok=\"t\" o:connecttype=\"custom\""
        + " o:connectlocs=\"@9,0;@10,10800;@11,21600;@12,10800\""
        + " o:connectangles=\"270,180,90,0\"/>"
        + "<v:textpath on=\"t\" fitshape=\"t\"/>"
        + "<v:handles><v:h position=\"#0,bottomRight\" xrange=\"6629,14971\"/></v:handles>"
        + "<o:lock v:ext=\"edit\" text=\"t\" shapetype=\"t\"/>"
        + "</v:shapetype>";

    private static string XmlEscapeAttribute(string value) =>
        value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
            .Replace("\"", "&quot;");

    private static T GetOrAddSectionChild<T>(SectionProperties section, Func<T> factory)
        where T : OpenXmlElement
    {
        var existing = section.GetFirstChild<T>();
        if (existing is not null)
            return existing;
        var element = factory();
        var order = Array.IndexOf(SectionChildOrder, element.LocalName);
        var before = section.ChildElements.FirstOrDefault(
            child => Array.IndexOf(SectionChildOrder, child.LocalName) > order);
        if (before is not null)
            section.InsertBefore(element, before);
        else
            section.Append(element);
        return element;
    }

    public static void ApplyHeaderFooter(MainDocumentPart main, HeaderFooter settings)
    {
        var body = main.Document?.Body;
        if (body is null)
            return;
        foreach (var section in body.Descendants<SectionProperties>().ToList())
        {
            if (!string.IsNullOrEmpty(settings.HeaderText))
                SetHeader(main, section, settings.HeaderText);
            if (!string.IsNullOrEmpty(settings.FooterText))
                SetFooter(main, section, settings.FooterText);
        }
    }

    private static void SetHeader(MainDocumentPart main, SectionProperties section, string text)
    {
        var reference = section.Elements<HeaderReference>()
            .FirstOrDefault(item => item.Type?.Value == HeaderFooterValues.Default);
        HeaderPart? part = null;
        if (reference?.Id?.Value is { } relationshipId)
            part = main.GetPartById(relationshipId) as HeaderPart;
        if (part is null)
        {
            part = main.AddNewPart<HeaderPart>();
            part.Header = new Header();
            AddHeaderReference(section, main.GetIdOfPart(part));
        }
        part.Header ??= new Header();
        var paragraph = part.Header.Elements<Paragraph>().FirstOrDefault();
        if (paragraph is null)
        {
            paragraph = new Paragraph(new ParagraphProperties(
                new ParagraphStyleId { Val = "Header" }));
            part.Header.Append(paragraph);
        }
        ClearParagraphContent(paragraph);
        paragraph.Append(NewTextRun(text, null));
        part.Header.Save();
    }

    private static void SetFooter(MainDocumentPart main, SectionProperties section, string text)
    {
        var reference = section.Elements<FooterReference>()
            .FirstOrDefault(item => item.Type?.Value == HeaderFooterValues.Default);
        FooterPart? part = null;
        if (reference?.Id?.Value is { } relationshipId)
            part = main.GetPartById(relationshipId) as FooterPart;
        if (part is null)
        {
            part = main.AddNewPart<FooterPart>();
            part.Footer = new Footer();
            AddFooterReference(section, main.GetIdOfPart(part));
        }
        part.Footer ??= new Footer();
        var paragraph = part.Footer.Elements<Paragraph>().FirstOrDefault();
        if (paragraph is null)
        {
            paragraph = new Paragraph(new ParagraphProperties(
                new ParagraphStyleId { Val = "Footer" }));
            part.Footer.Append(paragraph);
        }
        ClearParagraphContent(paragraph);
        if (text.Contains("{page}", StringComparison.Ordinal))
        {
            var parts = text.Split("{page}");
            for (var index = 0; index < parts.Length; index++)
            {
                if (parts[index].Length > 0)
                    paragraph.Append(NewTextRun(parts[index], null));
                if (index < parts.Length - 1)
                    paragraph.Append(BuildPageFieldRuns(" PAGE "));
            }
        }
        else
        {
            paragraph.Append(NewTextRun(text, null));
        }
        part.Footer.Save();
    }

    private static void AddHeaderReference(SectionProperties section, string relationshipId)
    {
        var reference = new HeaderReference
        {
            Type = HeaderFooterValues.Default,
            Id = relationshipId,
        };
        var lastHeader = section.Elements<HeaderReference>().LastOrDefault();
        if (lastHeader is not null)
        {
            lastHeader.InsertAfterSelf(reference);
            return;
        }
        var firstFooter = section.Elements<FooterReference>().FirstOrDefault();
        if (firstFooter is not null)
            section.InsertBefore(reference, firstFooter);
        else
            section.PrependChild(reference);
    }

    private static void AddFooterReference(SectionProperties section, string relationshipId)
    {
        var reference = new FooterReference
        {
            Type = HeaderFooterValues.Default,
            Id = relationshipId,
        };
        var lastFooter = section.Elements<FooterReference>().LastOrDefault();
        if (lastFooter is not null)
        {
            lastFooter.InsertAfterSelf(reference);
            return;
        }
        var lastHeader = section.Elements<HeaderReference>().LastOrDefault();
        if (lastHeader is not null)
            lastHeader.InsertAfterSelf(reference);
        else
            section.PrependChild(reference);
    }

    private static void ClearParagraphContent(Paragraph paragraph)
    {
        foreach (var child in paragraph.ChildElements.ToList())
        {
            if (child is ParagraphProperties)
                continue;
            child.Remove();
        }
    }

    // ── core properties ─────────────────────────────────────────

    public static void ApplyCoreProperties(
        WordprocessingDocument document, WordProperties properties)
    {
        var core = GetCoreProperties(document);
        if (properties.Title is not null)
            core.Title = properties.Title;
        if (properties.Subject is not null)
            core.Subject = properties.Subject;
        if (properties.Author is not null)
            core.Creator = properties.Author;
        if (properties.Keywords is not null)
            core.Keywords = properties.Keywords;
        if (properties.Comments is not null)
            core.Description = properties.Comments;
        if (properties.Category is not null)
            core.Category = properties.Category;
        if (properties.LastModifiedBy is not null)
            core.LastModifiedBy = properties.LastModifiedBy;
        if (properties.Created is { } created)
            core.Created = NormaliseCoreDateTime(created);
        if (properties.Modified is { } modified)
            core.Modified = NormaliseCoreDateTime(modified);
    }

    /// <summary>
    /// python-docx serialises datetimes as W3CDTF UTC with second precision;
    /// naive values keep their wall-clock time and are stamped as UTC.
    /// </summary>
    private static DateTime NormaliseCoreDateTime(DateTime value)
    {
        var utc = value.Kind switch
        {
            DateTimeKind.Utc => value,
            DateTimeKind.Local => value.ToUniversalTime(),
            _ => DateTime.SpecifyKind(value, DateTimeKind.Utc),
        };
        return new DateTime(
            utc.Ticks - utc.Ticks % TimeSpan.TicksPerSecond, DateTimeKind.Utc);
    }

    private static IPackageProperties GetCoreProperties(WordprocessingDocument document)
    {
        var part = document.CoreFilePropertiesPart
                   ?? document.AddCoreFilePropertiesPart();
        var empty = true;
        try
        {
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            empty = stream.Length == 0;
        }
        catch (Exception)
        {
            empty = true;
        }
        if (empty)
        {
            using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(stream, new UTF8Encoding(false));
            writer.Write(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<cp:coreProperties"
                + " xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\""
                + " xmlns:dc=\"http://purl.org/dc/elements/1.1/\""
                + " xmlns:dcterms=\"http://purl.org/dc/terms/\""
                + " xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\""
                + " xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"/>");
        }
        return part.CoreFileProperties;
    }

    public static Dictionary<string, object?> ReadCoreProperties(WordprocessingDocument document)
    {
        IPackageProperties core = document.CoreFilePropertiesPart?.CoreFileProperties
                                  ?? document.PackageProperties;
        return new Dictionary<string, object?>
        {
            ["title"] = core.Title,
            ["subject"] = core.Subject,
            ["author"] = core.Creator,
            ["keywords"] = core.Keywords,
            ["comments"] = core.Description,
            ["category"] = core.Category,
            ["last_modified_by"] = core.LastModifiedBy,
            ["created"] = IsoCore(core.Created),
            ["modified"] = IsoCore(core.Modified),
        };
    }

    private static string? IsoCore(DateTime? value)
    {
        if (value is null)
            return null;
        var utc = value.Value.Kind switch
        {
            DateTimeKind.Utc => value.Value,
            DateTimeKind.Local => value.Value.ToUniversalTime(),
            _ => DateTime.SpecifyKind(value.Value, DateTimeKind.Utc),
        };
        return Iso(utc);
    }

    public static string? Iso(DateTime? value)
    {
        if (value is null)
            return null;
        var dateTime = value.Value;
        var hasFraction = dateTime.Ticks % TimeSpan.TicksPerSecond != 0;
        var format = hasFraction ? "yyyy-MM-ddTHH:mm:ss.ffffff" : "yyyy-MM-ddTHH:mm:ss";
        return dateTime.Kind switch
        {
            DateTimeKind.Utc => dateTime.ToString(format, CultureInfo.InvariantCulture) + "+00:00",
            DateTimeKind.Local => new DateTimeOffset(dateTime).ToString(
                format + "zzz", CultureInfo.InvariantCulture),
            _ => dateTime.ToString(format, CultureInfo.InvariantCulture),
        };
    }

    public static string? IsoUtc(DateTime? value)
    {
        if (value is null)
            return null;
        var dateTime = value.Value;
        var utc = dateTime.Kind switch
        {
            DateTimeKind.Utc => dateTime,
            DateTimeKind.Local => dateTime.ToUniversalTime(),
            _ => DateTime.SpecifyKind(dateTime, DateTimeKind.Utc),
        };
        return Iso(utc);
    }

    // ── styles ───────────────────────────────────────────────────

    private static readonly string[] HeadingSizeHalfPoints =
        ["28", "26", "24", "22", "22", "22", "22", "22", "22"];

    private static readonly string[] HeadingSpacingBefore =
        ["480", "200", "200", "200", "200", "200", "200", "200", "200"];

    /// <summary>
    /// python-docx default.docx resolves minorHAnsi to Cambria and majorHAnsi to
    /// Calibri; both theme fonts map their Hans script to 宋体.  Blank documents
    /// use those resolved fonts, while template/appended documents keep the
    /// historical defaults when missing styles have to be synthesised.
    /// </summary>
    public static void EnsureRequiredStyles(
        MainDocumentPart main, bool pythonDocxThemeFonts = false)
    {
        var part = main.StyleDefinitionsPart;
        if (part is null)
        {
            part = main.AddNewPart<StyleDefinitionsPart>();
            part.Styles = new Styles();
        }
        part.Styles ??= new Styles();
        var styles = part.Styles;
        var bodyLatinFont = pythonDocxThemeFonts ? "Cambria" : "Calibri";
        var headingLatinFont = pythonDocxThemeFonts ? "Calibri" : "Calibri Light";
        var headingEastAsiaFont = pythonDocxThemeFonts ? "宋体" : "黑体";
        if (!styles.Elements<DocDefaults>().Any())
        {
            styles.PrependChild(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts
                    {
                        Ascii = bodyLatinFont,
                        HighAnsi = bodyLatinFont,
                        EastAsia = "宋体",
                        ComplexScript = bodyLatinFont,
                    },
                    new FontSize { Val = "22" },
                    new FontSizeComplexScript { Val = "22" },
                    new Languages { Val = "en-US", EastAsia = "zh-CN", Bidi = "ar-SA" })),
                new ParagraphPropertiesDefault(new ParagraphPropertiesBaseStyle(
                    new SpacingBetweenLines
                    {
                        After = "200",
                        Line = "276",
                        LineRule = LineSpacingRuleValues.Auto,
                    }))));
        }
        if (!styles.Elements<Style>().Any(style => style.StyleId?.Value == "Normal"))
        {
            styles.Append(new Style(
                new StyleName { Val = "Normal" },
                new StyleParagraphProperties(),
                new StyleRunProperties())
            {
                Type = StyleValues.Paragraph,
                StyleId = "Normal",
                Default = true,
            });
        }
        for (var level = 1; level <= 9; level++)
            EnsureHeadingStyle(styles, level, headingLatinFont, headingEastAsiaFont);
        EnsureListStyle(styles, "ListBullet", "List Bullet");
        EnsureListStyle(styles, "ListNumber", "List Number");
        EnsureHeaderFooterStyle(styles, "Header", "header");
        EnsureHeaderFooterStyle(styles, "Footer", "footer");
        EnsureTableNormalStyle(styles);
        EnsureTableGridStyle(styles);
        EnsureHyperlinkStyle(styles);
        styles.Save();
    }

    private static void EnsureHeadingStyle(
        Styles styles, int level, string latinFont, string eastAsiaFont)
    {
        var styleId = $"Heading{level}";
        if (styles.Elements<Style>().Any(style => style.StyleId?.Value == styleId))
            return;
        styles.Append(new Style(
            new StyleName { Val = $"heading {level}" },
            new BasedOn { Val = "Normal" },
            new NextParagraphStyle { Val = "Normal" },
            new StyleParagraphProperties(
                new KeepNext(),
                new KeepLines(),
                new SpacingBetweenLines
                {
                    Before = HeadingSpacingBefore[level - 1],
                    After = "0",
                },
                new OutlineLevel { Val = level - 1 }),
            new StyleRunProperties(
                new RunFonts
                {
                    Ascii = latinFont,
                    HighAnsi = latinFont,
                    EastAsia = eastAsiaFont,
                    ComplexScript = latinFont,
                },
                new Bold(),
                new Color { Val = "000000" },
                new FontSize { Val = HeadingSizeHalfPoints[level - 1] },
                new FontSizeComplexScript { Val = HeadingSizeHalfPoints[level - 1] }))
        {
            Type = StyleValues.Paragraph,
            StyleId = styleId,
        });
    }

    private static void EnsureListStyle(Styles styles, string styleId, string name)
    {
        if (styles.Elements<Style>().Any(style => style.StyleId?.Value == styleId))
            return;
        styles.Append(new Style(
            new StyleName { Val = name },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(new ContextualSpacing()))
        {
            Type = StyleValues.Paragraph,
            StyleId = styleId,
        });
    }

    /// <summary>Matches python-docx default.docx Header/Footer paragraph styles.</summary>
    private static void EnsureHeaderFooterStyle(Styles styles, string styleId, string name)
    {
        if (styles.Elements<Style>().Any(style => style.StyleId?.Value == styleId))
            return;
        styles.Append(new Style(
            new StyleName { Val = name },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new Tabs(
                    new TabStop { Val = TabStopValues.Center, Position = 4680 },
                    new TabStop { Val = TabStopValues.Right, Position = 9360 }),
                new SpacingBetweenLines
                {
                    After = "0",
                    Line = "240",
                    LineRule = LineSpacingRuleValues.Auto,
                }))
        {
            Type = StyleValues.Paragraph,
            StyleId = styleId,
        });
    }

    private static void EnsureTableNormalStyle(Styles styles)
    {
        if (styles.Elements<Style>().Any(style => style.StyleId?.Value == "TableNormal"))
            return;
        styles.Append(new Style(
            new StyleName { Val = "Normal Table" },
            new StyleTableProperties(
                new TableCellMarginDefault(
                    new TopMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new TableCellLeftMargin { Width = 108, Type = TableWidthValues.Dxa },
                    new BottomMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new TableCellRightMargin { Width = 108, Type = TableWidthValues.Dxa })))
        {
            Type = StyleValues.Table,
            StyleId = "TableNormal",
            Default = true,
        });
    }

    private static void EnsureTableGridStyle(Styles styles)
    {
        if (styles.Elements<Style>().Any(style => style.StyleId?.Value == "TableGrid"))
            return;
        styles.Append(new Style(
            new StyleName { Val = "Table Grid" },
            new BasedOn { Val = "TableNormal" },
            new StyleParagraphProperties(new SpacingBetweenLines
            {
                After = "0",
                Line = "240",
                LineRule = LineSpacingRuleValues.Auto,
            }),
            new StyleTableProperties(
                new TableIndentation { Width = 0, Type = TableWidthUnitValues.Dxa },
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                    new LeftBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                    new BottomBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                    new RightBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" })))
        {
            Type = StyleValues.Table,
            StyleId = "TableGrid",
        });
    }

    private static void EnsureHyperlinkStyle(Styles styles)
    {
        if (styles.Elements<Style>().Any(style => style.StyleId?.Value == "Hyperlink"))
            return;
        styles.Append(new Style(
            new StyleName { Val = "Hyperlink" },
            new StyleRunProperties(
                new Color { Val = "0563C1" },
                new Underline { Val = UnderlineValues.Single }))
        {
            Type = StyleValues.Character,
            StyleId = "Hyperlink",
        });
    }

    // ── images ───────────────────────────────────────────────────

    public static int MaxDocPrId(MainDocumentPart main) =>
        (main.Document?.Descendants<Wp.DocProperties>() ?? Enumerable.Empty<Wp.DocProperties>())
            .Select(item => (int)(item.Id?.Value ?? 0U))
            .DefaultIfEmpty(0)
            .Max();

    public static (int Width, int Height)? TryReadImageSize(string path)
    {
        byte[] data;
        try
        {
            using var stream = File.OpenRead(path);
            var length = (int)Math.Min(stream.Length, 4 * 1024 * 1024L);
            data = new byte[length];
            var read = 0;
            while (read < length)
            {
                var chunk = stream.Read(data, read, length - read);
                if (chunk <= 0)
                    break;
                read += chunk;
            }
            if (read < length)
                Array.Resize(ref data, read);
        }
        catch (IOException)
        {
            return null;
        }
        if (data.Length < 10)
            return null;
        if (data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47
            && data.Length >= 24)
        {
            var width = ReadInt32BigEndian(data, 16);
            var height = ReadInt32BigEndian(data, 20);
            return width > 0 && height > 0 ? (width, height) : null;
        }
        if (data[0] == (byte)'G' && data[1] == (byte)'I' && data[2] == (byte)'F'
            && data.Length >= 10)
        {
            var width = data[6] | (data[7] << 8);
            var height = data[8] | (data[9] << 8);
            return width > 0 && height > 0 ? (width, height) : null;
        }
        if (data[0] == (byte)'B' && data[1] == (byte)'M' && data.Length >= 26)
        {
            var width = Math.Abs(BitConverter.ToInt32(data, 18));
            var height = Math.Abs(BitConverter.ToInt32(data, 22));
            return width > 0 && height > 0 ? (width, height) : null;
        }
        if (data[0] == 0xFF && data[1] == 0xD8)
            return TryReadJpegSize(data);
        if ((data[0] == (byte)'I' && data[1] == (byte)'I' && data[2] == 0x2A && data[3] == 0x00)
            || (data[0] == (byte)'M' && data[1] == (byte)'M' && data[2] == 0x00 && data[3] == 0x2A))
            return TryReadTiffSize(data);
        return null;
    }

    private static int ReadInt32BigEndian(byte[] data, int offset) =>
        (data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3];

    private static (int Width, int Height)? TryReadJpegSize(byte[] data)
    {
        var index = 2;
        while (index + 9 < data.Length)
        {
            if (data[index] != 0xFF)
            {
                index++;
                continue;
            }
            var marker = data[index + 1];
            if (marker is 0xD8 or 0xD9)
            {
                index += 2;
                continue;
            }
            if (index + 4 > data.Length)
                return null;
            var length = (data[index + 2] << 8) | data[index + 3];
            if (length < 2)
                return null;
            var isStartOfFrame = marker is >= 0xC0 and <= 0xCF
                                 and not (0xC4 or 0xC8 or 0xCC);
            if (isStartOfFrame && index + 9 <= data.Length)
            {
                var height = (data[index + 5] << 8) | data[index + 6];
                var width = (data[index + 7] << 8) | data[index + 8];
                return width > 0 && height > 0 ? (width, height) : null;
            }
            index += 2 + length;
        }
        return null;
    }

    private static (int Width, int Height)? TryReadTiffSize(byte[] data)
    {
        var littleEndian = data[0] == (byte)'I';
        int Read16(int offset) => littleEndian
            ? data[offset] | (data[offset + 1] << 8)
            : (data[offset] << 8) | data[offset + 1];
        int Read32(int offset) => littleEndian
            ? data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) | (data[offset + 3] << 24)
            : ReadInt32BigEndian(data, offset);
        if (data.Length < 8)
            return null;
        var ifdOffset = Read32(4);
        if (ifdOffset + 2 > data.Length)
            return null;
        var count = Read16(ifdOffset);
        var width = 0;
        var height = 0;
        for (var i = 0; i < count; i++)
        {
            var entry = ifdOffset + 2 + i * 12;
            if (entry + 12 > data.Length)
                break;
            var tag = Read16(entry);
            var type = Read16(entry + 2);
            var value = type == 3 ? Read16(entry + 8) : Read32(entry + 8);
            if (tag == 256)
                width = value;
            else if (tag == 257)
                height = value;
        }
        return width > 0 && height > 0 ? (width, height) : null;
    }

    public static Drawing BuildInlineImage(
        string relationshipId, long cx, long cy, int docPrId, string name)
    {
        return new Drawing(new Wp.Inline(
            new Wp.Extent { Cx = cx, Cy = cy },
            new Wp.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
            new Wp.DocProperties { Id = (uint)docPrId, Name = name },
            new Wp.NonVisualGraphicFrameDrawingProperties(
                new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(
                new A.GraphicData(
                    new A.Picture(
                        new A.NonVisualPictureProperties(
                            new A.NonVisualDrawingProperties
                            {
                                Id = (uint)docPrId,
                                Name = $"picture{docPrId}",
                            },
                            new A.NonVisualPictureDrawingProperties()),
                        new A.BlipFill(
                            new A.Blip { Embed = relationshipId },
                            new A.Stretch(new A.FillRectangle())),
                        new A.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 0L, Y = 0L },
                                new A.Extents { Cx = cx, Cy = cy }),
                            new A.PresetGeometry(new A.AdjustValueList())
                            {
                                Preset = A.ShapeTypeValues.Rectangle })))
                {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture",
                }))
        {
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 0U,
            DistanceFromRight = 0U,
        });
    }

    public static PartTypeInfo ImagePartTypeFor(string path)
    {
        return Path.GetExtension(path).ToLowerInvariant() switch
        {
            ".png" => ImagePartType.Png,
            ".jpg" or ".jpeg" => ImagePartType.Jpeg,
            ".gif" => ImagePartType.Gif,
            ".bmp" => ImagePartType.Bmp,
            ".tiff" or ".tif" => ImagePartType.Tiff,
            _ => throw new NativeRuntimeException($"不支持的图片格式: {Path.GetExtension(path)}"),
        };
    }

    // ── hyperlinks ───────────────────────────────────────────────

    public static Paragraph BuildHyperlinkParagraph(
        MainDocumentPart main,
        HyperlinkDef link,
        FontConfig? defaultFont,
        WordParagraphFormat? format)
    {
        var paragraph = new Paragraph();
        var properties = BuildParagraphProperties(format);
        if (properties.HasChildren)
            paragraph.Append(properties);
        var uri = new Uri(link.Url, UriKind.RelativeOrAbsolute);
        var relationshipId = main.AddHyperlinkRelationship(uri, true).Id;
        var hyperlink = new Hyperlink
        {
            Id = relationshipId,
            History = true,
        };
        if (!string.IsNullOrEmpty(link.Tooltip))
            hyperlink.Tooltip = link.Tooltip;
        hyperlink.Append(NewTextRun(link.Text, new WordRunFormat
        {
            FontName = defaultFont?.Name,
            SizePt = defaultFont?.SizePt,
            Color = "0563C1",
            Underline = true,
        }));
        paragraph.Append(hyperlink);
        return paragraph;
    }

    // ── text extraction ─────────────────────────────────────────

    public static string RunText(Run run)
    {
        var builder = new StringBuilder();
        foreach (var child in run.ChildElements)
        {
            switch (child)
            {
                case Text text:
                    builder.Append(text.Text);
                    break;
                case TabChar:
                    builder.Append('\t');
                    break;
                case Break:
                case CarriageReturn:
                    builder.Append('\n');
                    break;
            }
        }
        return builder.ToString();
    }

    /// <summary>CT_P.text semantics: direct runs plus visible hyperlink text.</summary>
    public static string ParagraphText(Paragraph paragraph)
    {
        var builder = new StringBuilder();
        foreach (var child in paragraph.ChildElements)
        {
            switch (child)
            {
                case Run run:
                    builder.Append(RunText(run));
                    break;
                case Hyperlink hyperlink:
                    builder.Append(string.Concat(
                        hyperlink.Elements<Run>().Select(RunText)));
                    break;
            }
        }
        return builder.ToString();
    }

    /// <summary>All runs in document order, including those inside hyperlinks.</summary>
    public static string ParagraphTextAllRuns(Paragraph paragraph) =>
        string.Concat(paragraph.Descendants<Run>().Select(RunText));

    public static IEnumerable<Run> ParagraphRunsAll(Paragraph paragraph) =>
        paragraph.Descendants<Run>();

    public static void SetRunText(Run run, string text)
    {
        var textElements = run.Elements<Text>().ToList();
        Text main;
        if (textElements.Count > 0)
        {
            main = textElements[0];
            foreach (var extra in textElements.Skip(1))
                extra.Remove();
        }
        else
        {
            main = new Text();
            run.Append(main);
        }
        main.Text = text;
        main.Space = SpaceProcessingModeValues.Preserve;
    }

    // ── output helpers ──────────────────────────────────────────

    public static string TempSiblingPath(string targetPath)
    {
        var full = Path.GetFullPath(targetPath);
        var directory = Path.GetDirectoryName(full)
                        ?? throw new NativeRuntimeException($"无法解析输出目录: {targetPath}");
        Directory.CreateDirectory(directory);
        var stem = Path.GetFileNameWithoutExtension(full);
        var extension = Path.GetExtension(full);
        if (string.IsNullOrEmpty(extension))
            extension = ".docx";
        return Path.Combine(directory,
            $".{stem}-{Guid.NewGuid():N}.tmp{extension}");
    }

    public static void ReplaceFile(string tempPath, string targetPath)
    {
        var fullTarget = Path.GetFullPath(targetPath);
        Directory.CreateDirectory(Path.GetDirectoryName(fullTarget)!);
        if (File.Exists(fullTarget))
            File.Replace(tempPath, fullTarget, null, ignoreMetadataErrors: true);
        else
            File.Move(tempPath, fullTarget);
    }

    public static void SaveThroughTemp(string targetPath, Action<WordprocessingDocument> action)
    {
        var tempPath = TempSiblingPath(targetPath);
        File.Copy(Path.GetFullPath(targetPath), tempPath, overwrite: true);
        try
        {
            using (var document = WordprocessingDocument.Open(tempPath, true))
            {
                action(document);
            }
            ReplaceFile(tempPath, targetPath);
        }
        finally
        {
            if (File.Exists(tempPath))
            {
                try { File.Delete(tempPath); } catch { /* best effort */ }
            }
        }
    }

    public static void ValidateOfficeInput(string filePath)
    {
        var info = new FileInfo(filePath);
        if (info.Length > 100L * 1024 * 1024)
            throw new NativeRuntimeException("文件超过 100 MB 安全读取上限");
        try
        {
            using var archive = ZipFile.OpenRead(filePath);
            long total = 0;
            foreach (var entry in archive.Entries)
            {
                try
                {
                    total += entry.Length;
                }
                catch (InvalidDataException)
                {
                    throw new NativeRuntimeException("Office 文件损坏或不是有效的 OOXML 压缩包");
                }
                if (total > 500L * 1024 * 1024)
                    throw new NativeRuntimeException("Office 文件解压后超过 500 MB 安全上限");
            }
        }
        catch (InvalidDataException)
        {
            throw new NativeRuntimeException("Office 文件损坏或不是有效的 OOXML 压缩包");
        }
    }
}

/// <summary>Lazily creates bullet/numbered numbering definitions for direct numPr use.</summary>
internal sealed class WordListNumbering
{
    private readonly MainDocumentPart _main;
    private int? _bulletNumId;
    private int? _numberedNumId;

    public WordListNumbering(MainDocumentPart main) => _main = main;

    public int GetNumberId(bool bullet)
    {
        if (_bulletNumId is null || _numberedNumId is null)
            CreateDefinitions();
        return bullet ? _bulletNumId!.Value : _numberedNumId!.Value;
    }

    private void CreateDefinitions()
    {
        var part = _main.NumberingDefinitionsPart
                   ?? _main.AddNewPart<NumberingDefinitionsPart>();
        part.Numbering ??= new Numbering();
        var numbering = part.Numbering;
        var maxAbstract = numbering.Elements<AbstractNum>()
            .Select(item => item.AbstractNumberId?.Value ?? 0)
            .DefaultIfEmpty(-1)
            .Max();
        var maxNum = numbering.Elements<NumberingInstance>()
            .Select(item => item.NumberID?.Value ?? 0)
            .DefaultIfEmpty(0)
            .Max();
        var bulletAbstractId = maxAbstract + 1;
        var numberedAbstractId = maxAbstract + 2;
        var bulletAbstract = BuildAbstractNum(bulletAbstractId, bullet: true);
        var numberedAbstract = BuildAbstractNum(numberedAbstractId, bullet: false);
        var firstNumber = numbering.Elements<NumberingInstance>().FirstOrDefault();
        if (firstNumber is not null)
        {
            numbering.InsertBefore(bulletAbstract, firstNumber);
            numbering.InsertBefore(numberedAbstract, firstNumber);
        }
        else
        {
            numbering.Append(bulletAbstract);
            numbering.Append(numberedAbstract);
        }
        numbering.Append(new NumberingInstance(
            new AbstractNumId { Val = bulletAbstractId })
        { NumberID = maxNum + 1 });
        numbering.Append(new NumberingInstance(
            new AbstractNumId { Val = numberedAbstractId })
        { NumberID = maxNum + 2 });
        _bulletNumId = maxNum + 1;
        _numberedNumId = maxNum + 2;
        numbering.Save();
    }

    private static AbstractNum BuildAbstractNum(int abstractId, bool bullet)
    {
        var level = new Level(
            new StartNumberingValue { Val = 1 },
            new NumberingFormat
            {
                Val = bullet ? NumberFormatValues.Bullet : NumberFormatValues.Decimal,
            },
            new LevelText { Val = bullet ? "\uF0B7" : "%1." },
            new LevelJustification { Val = LevelJustificationValues.Left },
            new PreviousParagraphProperties(
                new Tabs(new TabStop
                {
                    Val = TabStopValues.Number,
                    Position = 360,
                }),
                new Indentation { Left = "360", Hanging = "360" }))
        { LevelIndex = 0 };
        if (bullet)
        {
            level.Append(new NumberingSymbolRunProperties(
                new RunFonts { Ascii = "Symbol", HighAnsi = "Symbol", Hint = FontTypeHintValues.Default }));
        }
        return new AbstractNum(
            new MultiLevelType { Val = MultiLevelValues.SingleLevel },
            level)
        { AbstractNumberId = abstractId };
    }
}
