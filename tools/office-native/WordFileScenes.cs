using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeNative;

/// <summary>Port of office_engine/scenes.py SceneProfile.</summary>
internal sealed class SceneProfile
{
    public required string Name { get; init; }
    public string DisplayName { get; init; } = "";
    public string? Description { get; init; }
    public string FontName { get; init; } = "宋体";
    public string FontNameWestern { get; init; } = "Times New Roman";
    public string? HeadingFont { get; init; }
    public string BodySize { get; init; } = "小四";
    public double BodySizePt => WordFileOoxml.CnSizeToPt(BodySize);
    public Dictionary<int, string> HeadingSizes { get; init; } = new()
    {
        [1] = "三号",
        [2] = "小三",
        [3] = "四号",
        [4] = "小四",
    };
    public double LineSpacing { get; init; } = 1.3;
    public int? FirstLineIndent { get; init; } = 28;
    public string Alignment { get; init; } = "justify";
    public string PaperSize { get; init; } = "A4";
    public double MarginTop { get; init; } = 1.0;
    public double MarginBottom { get; init; } = 1.0;
    public double MarginLeft { get; init; } = 1.18;
    public double MarginRight { get; init; } = 1.18;
    public string PageNumberStyle { get; init; } = "arabic";
    public bool ShowCover { get; init; }

    public double? HeadingSizePt(int level)
    {
        if (!HeadingSizes.TryGetValue(level, out var size))
            return null;
        return WordFileOoxml.CnSizeToPt(size);
    }
}

internal static class WordFileScenes
{
    public static readonly SceneProfile Academic = new()
    {
        Name = "academic",
        DisplayName = "学术论文",
        Description = "学术论文/学位论文排版，宋体正文+黑体标题，1.3倍行距，两字符缩进",
        FontName = "宋体",
        FontNameWestern = "Times New Roman",
        HeadingFont = "黑体",
        BodySize = "小四",
        HeadingSizes = new() { [1] = "三号", [2] = "四号", [3] = "小四", [4] = "五号" },
        LineSpacing = 1.3,
        FirstLineIndent = 28,
        Alignment = "justify",
        PaperSize = "A4",
        MarginTop = 1.0,
        MarginBottom = 1.0,
        MarginLeft = 1.18,
        MarginRight = 1.18,
        PageNumberStyle = "arabic",
        ShowCover = false,
    };

    public static readonly SceneProfile Report = new()
    {
        Name = "report",
        DisplayName = "报告/提案",
        Description = "正式报告排版，宋体正文+黑体标题，含封面和执行摘要结构",
        FontName = "宋体",
        FontNameWestern = "Times New Roman",
        HeadingFont = "黑体",
        BodySize = "小四",
        HeadingSizes = new() { [1] = "三号", [2] = "四号", [3] = "小四", [4] = "五号" },
        LineSpacing = 1.3,
        FirstLineIndent = 28,
        Alignment = "justify",
        PaperSize = "A4",
        MarginTop = 1.0,
        MarginBottom = 1.0,
        MarginLeft = 1.18,
        MarginRight = 1.18,
        PageNumberStyle = "arabic",
        ShowCover = true,
    };

    public static readonly SceneProfile Official = new()
    {
        Name = "official",
        DisplayName = "公文",
        Description = "GB/T 9704 公文排版，仿宋_GB2312 正文，黑体标题，28pt 固定行距",
        FontName = "仿宋_GB2312",
        FontNameWestern = "Times New Roman",
        HeadingFont = "黑体",
        BodySize = "三号",
        HeadingSizes = new() { [1] = "二号", [2] = "三号", [3] = "小三" },
        LineSpacing = 1.75,
        FirstLineIndent = 40,
        Alignment = "justify",
        PaperSize = "A4",
        MarginTop = 1.46,
        MarginBottom = 1.38,
        MarginLeft = 1.1,
        MarginRight = 1.0,
        PageNumberStyle = "arabic",
        ShowCover = false,
    };

    public static readonly SceneProfile Contract = new()
    {
        Name = "contract",
        DisplayName = "合同/协议",
        Description = "合同排版，宋体正文，1.5倍行距，含签约信息和签名区",
        FontName = "宋体",
        FontNameWestern = "Times New Roman",
        HeadingFont = "黑体",
        BodySize = "小四",
        HeadingSizes = new() { [1] = "三号", [2] = "四号", [3] = "小四" },
        LineSpacing = 1.5,
        FirstLineIndent = 28,
        Alignment = "justify",
        PaperSize = "A4",
        MarginTop = 1.0,
        MarginBottom = 1.0,
        MarginLeft = 1.18,
        MarginRight = 1.18,
        PageNumberStyle = "arabic",
        ShowCover = false,
    };

    public static readonly SceneProfile Resume = new()
    {
        Name = "resume",
        DisplayName = "简历",
        Description = "简历排版，微软雅黑正文，1.15倍行距，紧凑排版",
        FontName = "微软雅黑",
        FontNameWestern = "Calibri",
        HeadingFont = "微软雅黑",
        BodySize = "五号",
        HeadingSizes = new() { [1] = "三号", [2] = "四号", [3] = "五号" },
        LineSpacing = 1.15,
        FirstLineIndent = null,
        Alignment = "left",
        PaperSize = "A4",
        MarginTop = 0.7,
        MarginBottom = 0.7,
        MarginLeft = 0.8,
        MarginRight = 0.8,
        PageNumberStyle = "none",
        ShowCover = false,
    };

    public static readonly SceneProfile Exam = new()
    {
        Name = "exam",
        DisplayName = "试卷/测验",
        Description = "试卷排版，宋体正文，预留答题区域，分值标记",
        FontName = "宋体",
        FontNameWestern = "Times New Roman",
        HeadingFont = "黑体",
        BodySize = "小四",
        HeadingSizes = new() { [1] = "三号", [2] = "四号", [3] = "小四" },
        LineSpacing = 1.5,
        FirstLineIndent = null,
        Alignment = "left",
        PaperSize = "A4",
        MarginTop = 0.8,
        MarginBottom = 0.8,
        MarginLeft = 0.8,
        MarginRight = 0.8,
        PageNumberStyle = "arabic",
        ShowCover = false,
    };

    public static readonly SceneProfile Journal = new()
    {
        Name = "journal",
        DisplayName = "期刊论文",
        Description = "期刊论文排版（测绘学报规范），宋体正文+黑体标题，五号正文，1.3倍行距。",
        FontName = "宋体",
        FontNameWestern = "Times New Roman",
        HeadingFont = "黑体",
        BodySize = "五号",
        HeadingSizes = new() { [1] = "小四", [2] = "五号", [3] = "五号" },
        LineSpacing = 1.3,
        FirstLineIndent = 28,
        Alignment = "justify",
        PaperSize = "A4",
        MarginTop = 1.0,
        MarginBottom = 1.0,
        MarginLeft = 1.18,
        MarginRight = 1.18,
        PageNumberStyle = "arabic",
        ShowCover = false,
    };

    public static readonly Dictionary<string, SceneProfile> Registry =
        new(StringComparer.Ordinal)
        {
            ["academic"] = Academic,
            ["report"] = Report,
            ["official"] = Official,
            ["contract"] = Contract,
            ["resume"] = Resume,
            ["exam"] = Exam,
            ["journal"] = Journal,
        };

    public static SceneProfile? Get(string? name) =>
        name is not null && Registry.TryGetValue(name, out var profile) ? profile : null;
}

/// <summary>Port of office_engine/covers.py CoverRecipe.</summary>
internal sealed class CoverRecipe
{
    public required string RecipeId { get; init; }
    public string DisplayName { get; init; } = "";
    public string Layout { get; init; } = "centered";
    public string TitleFont { get; init; } = "黑体";
    public string SubtitleFont { get; init; } = "宋体";
    public string MetaFont { get; init; } = "宋体";
    public int TitleSizeMin { get; init; } = 22;
    public int TitleSizeMax { get; init; } = 40;
    public int SubtitleSize { get; init; } = 16;
    public int MetaSize { get; init; } = 12;
    public bool UseHorizontalLine { get; init; }
    public string TitleColor { get; init; } = "000000";
    public string SubtitleColor { get; init; } = "333333";
    public string? BgColor { get; init; }
}

internal static class WordFileCovers
{
    public static readonly Dictionary<string, CoverRecipe> Recipes = new(StringComparer.Ordinal)
    {
        ["R1"] = new CoverRecipe
        {
            RecipeId = "R1",
            DisplayName = "经典居中",
            Layout = "centered",
            TitleFont = "黑体",
            SubtitleFont = "宋体",
            MetaFont = "宋体",
            SubtitleSize = 16,
            MetaSize = 12,
            UseHorizontalLine = true,
        },
        ["R2"] = new CoverRecipe
        {
            RecipeId = "R2",
            DisplayName = "现代简约",
            Layout = "modern",
            TitleFont = "微软雅黑",
            SubtitleFont = "微软雅黑",
            MetaFont = "宋体",
            TitleSizeMin = 24,
            TitleSizeMax = 38,
            SubtitleSize = 15,
            MetaSize = 11,
            UseHorizontalLine = false,
        },
        ["R3"] = new CoverRecipe
        {
            RecipeId = "R3",
            DisplayName = "上下分割",
            Layout = "split",
            TitleFont = "黑体",
            SubtitleFont = "宋体",
            MetaFont = "宋体",
            TitleSizeMin = 22,
            TitleSizeMax = 36,
            SubtitleSize = 15,
            MetaSize = 11,
            BgColor = "F5F5F5",
        },
        ["R4"] = new CoverRecipe
        {
            RecipeId = "R4",
            DisplayName = "学术论文",
            Layout = "academic",
            TitleFont = "黑体",
            SubtitleFont = "宋体",
            MetaFont = "宋体",
            TitleSizeMin = 26,
            TitleSizeMax = 42,
            SubtitleSize = 16,
            MetaSize = 12,
            UseHorizontalLine = true,
        },
        ["R5"] = new CoverRecipe
        {
            RecipeId = "R5",
            DisplayName = "公文红头",
            Layout = "official",
            TitleFont = "宋体",
            SubtitleFont = "仿宋_GB2312",
            MetaFont = "仿宋_GB2312",
            TitleSizeMin = 22,
            TitleSizeMax = 30,
            SubtitleSize = 16,
            MetaSize = 14,
            UseHorizontalLine = false,
            TitleColor = "CC0000",
        },
        ["R6"] = new CoverRecipe
        {
            RecipeId = "R6",
            DisplayName = "极简",
            Layout = "minimal",
            TitleFont = "微软雅黑",
            SubtitleFont = "微软雅黑",
            MetaFont = "微软雅黑",
            TitleSizeMin = 20,
            TitleSizeMax = 36,
            SubtitleSize = 14,
            MetaSize = 10,
            UseHorizontalLine = false,
        },
        ["R7"] = new CoverRecipe
        {
            RecipeId = "R7",
            DisplayName = "品牌展示",
            Layout = "brand",
            TitleFont = "黑体",
            SubtitleFont = "微软雅黑",
            MetaFont = "宋体",
            TitleSizeMin = 24,
            TitleSizeMax = 40,
            SubtitleSize = 16,
            MetaSize = 11,
            UseHorizontalLine = true,
            BgColor = "1A1A2E",
            TitleColor = "FFFFFF",
            SubtitleColor = "CCCCCC",
        },
    };

    public static CoverRecipe Select(string? docType = "general", string? scene = null)
    {
        _ = docType;
        return scene switch
        {
            "academic" => Recipes["R4"],
            "official" => Recipes["R5"],
            "report" => Recipes["R1"],
            "contract" => Recipes["R1"],
            "resume" => Recipes["R6"],
            "exam" => Recipes["R6"],
            _ => Recipes["R1"],
        };
    }

    private static int CalcTitleSize(string title) => title.Length switch
    {
        <= 6 => 40,
        <= 10 => 34,
        <= 14 => 28,
        <= 20 => 24,
        _ => 22,
    };

    public static void Generate(
        Body body,
        CoverRecipe recipe,
        string title,
        string? subtitle,
        string? author,
        string? date,
        string? organization,
        IReadOnlyList<string>? extraLines)
    {
        OpenXmlElement host = body;
        if (recipe.BgColor is not null)
        {
            var section = body.Descendants<SectionProperties>().FirstOrDefault();
            var (usableWidth, usableHeightInches) = SectionMetrics(section);
            var rowHeight = (uint)Math.Round(
                Math.Max(1.0, usableHeightInches) * WordFileOoxml.TwipsPerInch);
            var tableCell = new TableCell(
                new TableCellProperties(
                    new TableCellWidth
                    {
                        Width = usableWidth.ToString(CultureInfo.InvariantCulture),
                        Type = TableWidthUnitValues.Dxa,
                    },
                    new Shading
                    {
                        Val = ShadingPatternValues.Clear,
                        Fill = recipe.BgColor,
                    }),
                new Paragraph());
            var table = new Table(
                new TableProperties(
                    new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto },
                    new TableBorders(
                        new TopBorder { Val = BorderValues.None, Size = 0, Space = 0 },
                        new LeftBorder { Val = BorderValues.None, Size = 0, Space = 0 },
                        new BottomBorder { Val = BorderValues.None, Size = 0, Space = 0 },
                        new RightBorder { Val = BorderValues.None, Size = 0, Space = 0 },
                        new InsideHorizontalBorder { Val = BorderValues.None, Size = 0, Space = 0 },
                        new InsideVerticalBorder { Val = BorderValues.None, Size = 0, Space = 0 }),
                    new TableLayout { Type = TableLayoutValues.Fixed }),
                new TableGrid(new GridColumn
                {
                    Width = usableWidth.ToString(CultureInfo.InvariantCulture),
                }),
                new TableRow(
                    new TableRowProperties(new TableRowHeight
                    {
                        Val = rowHeight,
                        HeightType = HeightRuleValues.AtLeast,
                    }),
                    tableCell));
            WordFileOoxml.AppendBodyElement(body, table);
            host = tableCell;
        }

        var titleSize = CalcTitleSize(title);
        AppendCoverParagraph(host, "", recipe.TitleFont, 10, "000000",
            bold: false, spaceBefore: 0, spaceAfter: 0, lineSpacing: true);
        AppendCoverParagraph(host, title, recipe.TitleFont, titleSize, recipe.TitleColor,
            bold: true, spaceBefore: 60, spaceAfter: 12, lineSpacing: true);
        if (recipe.UseHorizontalLine)
            AppendHorizontalLine(host);
        if (!string.IsNullOrEmpty(subtitle))
        {
            AppendCoverParagraph(host, subtitle, recipe.SubtitleFont, recipe.SubtitleSize,
                recipe.SubtitleColor, bold: false, spaceBefore: 20, spaceAfter: 8,
                lineSpacing: true);
        }
        AppendCoverParagraph(host, "", recipe.TitleFont, 10, "000000",
            bold: false, spaceBefore: 0, spaceAfter: 0, lineSpacing: true);

        var metaLines = new List<string>();
        if (!string.IsNullOrEmpty(author))
            metaLines.Add(author!);
        if (!string.IsNullOrEmpty(date))
            metaLines.Add(date!);
        if (!string.IsNullOrEmpty(organization))
            metaLines.Add(organization!);
        if (extraLines is not null)
            metaLines.AddRange(extraLines.Where(line => !string.IsNullOrEmpty(line)));
        var before = 40;
        foreach (var line in metaLines)
        {
            AppendCoverParagraph(host, line, recipe.MetaFont, recipe.MetaSize,
                recipe.SubtitleColor, bold: false, spaceBefore: before, spaceAfter: 2,
                lineSpacing: true);
            before = 0;
        }

        WordFileOoxml.AppendBodyElement(body,
            new Paragraph(new Run(new Break { Type = BreakValues.Page })));
    }

    private static void AppendHorizontalLine(OpenXmlElement host)
    {
        var format = new WordParagraphFormat
        {
            Alignment = "center",
            SpaceBeforePt = 6,
            SpaceAfterPt = 6,
            BottomBorder = true,
            BorderColor = "000000",
        };
        var paragraph = new Paragraph();
        var properties = WordFileOoxml.BuildParagraphProperties(format);
        if (properties.HasChildren)
            paragraph.Append(properties);
        AppendToHost(host, paragraph);
    }

    private static void AppendCoverParagraph(
        OpenXmlElement host,
        string text,
        string fontName,
        double sizePt,
        string color,
        bool bold,
        int spaceBefore,
        int spaceAfter,
        bool lineSpacing)
    {
        var format = new WordParagraphFormat
        {
            Alignment = "center",
            SpaceBeforePt = spaceBefore == 0 ? null : spaceBefore,
            SpaceAfterPt = spaceAfter == 0 ? null : spaceAfter,
        };
        if (lineSpacing)
            format.LineSpacing = 1.5;
        var paragraph = new Paragraph();
        var properties = WordFileOoxml.BuildParagraphProperties(format);
        if (properties.HasChildren)
            paragraph.Append(properties);
        if (text.Length > 0)
        {
            paragraph.Append(WordFileOoxml.NewTextRun(text, new WordRunFormat
            {
                FontName = fontName,
                SizePt = sizePt,
                Bold = bold ? true : null,
                Color = color,
            }));
        }
        AppendToHost(host, paragraph);
    }

    private static void AppendToHost(OpenXmlElement host, OpenXmlElement element)
    {
        if (host is Body body)
            WordFileOoxml.AppendBodyElement(body, element);
        else
            host.AppendChild(element);
    }

    private static (uint UsableWidthTwips, double UsableHeightInches) SectionMetrics(
        SectionProperties? section)
    {
        var pageSize = section?.GetFirstChild<PageSize>();
        var width = pageSize?.Width?.Value ?? 12240U;
        var height = pageSize?.Height?.Value ?? 15840U;
        var margin = section?.GetFirstChild<PageMargin>();
        var left = margin?.Left?.Value ?? 1800U;
        var right = margin?.Right?.Value ?? 1800U;
        var top = margin?.Top?.Value ?? 1440;
        var bottom = margin?.Bottom?.Value ?? 1440;
        var usableWidth = (uint)Math.Max(1L, (long)width - left - right);
        var usableHeight = (height - top - bottom) / WordFileOoxml.TwipsPerInch;
        return (usableWidth, usableHeight);
    }
}
