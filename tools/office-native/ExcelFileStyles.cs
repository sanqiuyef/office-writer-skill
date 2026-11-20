using System.Globalization;
using System.Numerics;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeNative;

/// <summary>
/// 读取时的数字格式判定：对齐 openpyxl 的 cellXfs → numFmtId → numFmts 链，
/// 以及 data_only 读取后 str(value) 的输出形态（日期/时间/时长/Python float）。
/// </summary>
internal sealed class ExcelReadFormats
{
    private static readonly Regex LiteralGroup = new(
        "\".*?\"", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex LocaleGroup = new(
        @"\[(?!hh?\]|mm?\]|ss?\])[^\]]*\]",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex DateChars = new(
        @"(?<![_\\])[dmhysDMHYS]",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex TimeDeltaPattern = new(
        @"\[hh?\](:mm(:ss(\.0*)?)?)?|\[mm?\](:ss(\.0*)?)?|\[ss?\](\.0*)?",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private readonly HashSet<int> _dateStyles = new();
    private readonly HashSet<int> _timedeltaStyles = new();

    public bool Date1904 { get; }

    private ExcelReadFormats(bool date1904) => Date1904 = date1904;

    public static ExcelReadFormats Load(WorkbookPart workbookPart)
    {
        var formats = new ExcelReadFormats(
            workbookPart.Workbook?.WorkbookProperties?.Date1904?.Value == true);
        var stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet;
        var custom = new Dictionary<uint, string>();
        foreach (var format in stylesheet?.NumberingFormats?.Elements<NumberingFormat>()
                                 ?? Enumerable.Empty<NumberingFormat>())
        {
            if (format.NumberFormatId?.Value is { } id && format.FormatCode?.Value is { } code)
                custom[id] = code;
        }
        var cellFormats = stylesheet?.CellFormats?.Elements<CellFormat>().ToList() ?? [];
        for (var index = 0; index < cellFormats.Count; index++)
        {
            var numberFormatId = cellFormats[index].NumberFormatId?.Value ?? 0;
            var code = custom.TryGetValue(numberFormatId, out var customCode)
                ? customCode
                : BuiltinFormatCode(numberFormatId);
            if (code is null)
                continue;
            if (IsDateFormat(code))
                formats._dateStyles.Add(index);
            if (IsTimedeltaFormat(code))
                formats._timedeltaStyles.Add(index);
        }
        return formats;
    }

    public bool IsDateStyle(int styleIndex) => _dateStyles.Contains(styleIndex);

    public bool IsTimedeltaStyle(int styleIndex) => _timedeltaStyles.Contains(styleIndex);

    /// <summary>序列号 → openpyxl from_excel 的字符串（datetime/date/time 的 str()）。</summary>
    public string FormatDateValue(string raw)
    {
        if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
            return raw;
        if (!double.IsFinite(value))
            return "#VALUE!";
        var day = Math.Floor(value);
        var fraction = value - day;
        var diff = TimeSpan.FromMilliseconds(
            Math.Round(fraction * 86400d * 1000d, MidpointRounding.ToEven));
        if (value >= 0 && value < 1 && diff.Days == 0)
            return FormatPythonTime(diff);
        if (value > 0 && value < 60 && !Date1904)
            day += 1;
        var epoch = Date1904 ? new DateTime(1904, 1, 1) : new DateTime(1899, 12, 30);
        try
        {
            var result = epoch.AddDays(day).Add(diff);
            var text = result.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            if (result.Millisecond != 0)
                text += "." + (result.Millisecond * 1000).ToString("D6", CultureInfo.InvariantCulture);
            return text;
        }
        catch (ArgumentOutOfRangeException)
        {
            // openpyxl 对越界日期返回 #VALUE!（data_type='e'）。
            return "#VALUE!";
        }
    }

    /// <summary>序列号 → Python str(timedelta) 输出。</summary>
    public string FormatTimedeltaValue(string raw)
    {
        if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
            return raw;
        if (!double.IsFinite(value) || Math.Abs(value) > 999_999_999d)
            return "#VALUE!";
        var seconds = value * 86400d;
        var micro = Math.Round(seconds * 1_000_000d, MidpointRounding.ToEven);
        if (micro > long.MaxValue || micro < long.MinValue)
            return "#VALUE!";
        var totalMicroseconds = (long)micro;
        // from_excel 对带微秒的 timedelta 先向下取整秒，再把微秒四舍五入到毫秒。
        var microPart = PositiveMod(totalMicroseconds, 1_000_000L);
        if (microPart != 0)
        {
            var wholeSeconds = FloorDiv(totalMicroseconds, 1_000_000L);
            var roundedMicro = (long)Math.Round(
                microPart / 1000d, MidpointRounding.ToEven) * 1000L;
            totalMicroseconds = wholeSeconds * 1_000_000L + roundedMicro;
        }

        var days = FloorDiv(totalMicroseconds, 86_400_000_000L);
        var remainder = totalMicroseconds - days * 86_400_000_000L;
        var hours = remainder / 3_600_000_000L;
        remainder %= 3_600_000_000L;
        var minutes = remainder / 60_000_000L;
        remainder %= 60_000_000L;
        var secs = remainder / 1_000_000L;
        var fraction = remainder % 1_000_000L;
        var text = string.Create(CultureInfo.InvariantCulture,
            $"{hours}:{minutes:D2}:{secs:D2}");
        if (days != 0)
            text = $"{days} day{(days == 1 ? "" : "s")}, " + text;
        if (fraction != 0)
            text += "." + fraction.ToString("D6", CultureInfo.InvariantCulture);
        return text;
    }

    /// <summary>
    /// 普通数值单元格输出：openpyxl 对无 "."/"e"/"E" 的原始文本取 int，
    /// 其余取 float 后 str()（Python repr 规则）。
    /// </summary>
    public static string FormatNumber(string raw)
    {
        if (raw.Length == 0)
            return "";
        if (raw.IndexOfAny(['.', 'e', 'E']) < 0)
        {
            return BigInteger.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture,
                out var integer)
                ? integer.ToString(CultureInfo.InvariantCulture)
                : raw;
        }
        if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
            return raw;
        return FormatPythonDouble(value);
    }

    private static string FormatPythonDouble(double value)
    {
        if (double.IsNaN(value))
            return "nan";
        if (double.IsPositiveInfinity(value))
            return "inf";
        if (double.IsNegativeInfinity(value))
            return "-inf";
        if (value == 0)
            return double.IsNegative(value) ? "-0.0" : "0.0";

        var negative = value < 0;
        var text = value.ToString("R", CultureInfo.InvariantCulture);
        if (text.StartsWith('-'))
            text = text[1..];
        var exponent = 0;
        var exponentIndex = text.IndexOfAny(['E', 'e']);
        if (exponentIndex >= 0)
        {
            int.TryParse(text[(exponentIndex + 1)..], NumberStyles.Integer,
                CultureInfo.InvariantCulture, out exponent);
            text = text[..exponentIndex];
        }
        var dot = text.IndexOf('.');
        var decimalPoint = dot >= 0 ? dot : text.Length;
        var digits = dot >= 0 ? text.Remove(dot, 1) : text;
        var digitPoint = decimalPoint + exponent;

        var leadingZeros = 0;
        while (leadingZeros < digits.Length - 1 && digits[leadingZeros] == '0')
            leadingZeros++;
        if (leadingZeros > 0)
        {
            digits = digits[leadingZeros..];
            digitPoint -= leadingZeros;
        }
        digits = digits.TrimEnd('0');
        if (digits.Length == 0)
            digits = "0";

        var sign = negative ? "-" : "";
        var decimalExponent = digitPoint - 1;
        if (digitPoint <= -4 || digitPoint > 16)
        {
            var mantissa = digits.Length == 1
                ? digits
                : digits[..1] + "." + digits[1..];
            var exponentText = Math.Abs(decimalExponent).ToString("D2", CultureInfo.InvariantCulture);
            return sign + mantissa + "e" + (decimalExponent < 0 ? "-" : "+") + exponentText;
        }
        if (digitPoint <= 0)
            return sign + "0." + new string('0', -digitPoint) + digits;
        if (digitPoint >= digits.Length)
            return sign + digits + new string('0', digitPoint - digits.Length) + ".0";
        return sign + digits[..digitPoint] + "." + digits[digitPoint..];
    }

    private static string FormatPythonTime(TimeSpan value)
    {
        var text = string.Create(CultureInfo.InvariantCulture,
            $"{value.Hours:D2}:{value.Minutes:D2}:{value.Seconds:D2}");
        var micro = (value.Ticks % TimeSpan.TicksPerSecond) / 10;
        if (micro != 0)
            text += "." + micro.ToString("D6", CultureInfo.InvariantCulture);
        return text;
    }

    private static long FloorDiv(long value, long divisor)
    {
        var quotient = value / divisor;
        if (value % divisor != 0 && (value < 0) != (divisor < 0))
            quotient--;
        return quotient;
    }

    private static long PositiveMod(long value, long divisor) =>
        value - FloorDiv(value, divisor) * divisor;

    private static bool IsDateFormat(string format)
    {
        var first = format.Split(';')[0];
        first = LiteralGroup.Replace(first, "");
        first = LocaleGroup.Replace(first, "");
        return DateChars.IsMatch(first);
    }

    private static bool IsTimedeltaFormat(string format) =>
        TimeDeltaPattern.IsMatch(format.Split(';')[0]);

    private static string? BuiltinFormatCode(uint id) => id switch
    {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        5 => "\"$\"#,##0_);(\"$\"#,##0)",
        6 => "\"$\"#,##0_);[Red](\"$\"#,##0)",
        7 => "\"$\"#,##0.00_);(\"$\"#,##0.00)",
        8 => "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "mm-dd-yy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yy h:mm",
        37 => "#,##0_);(#,##0)",
        38 => "#,##0_);[Red](#,##0)",
        39 => "#,##0.00_);(#,##0.00)",
        40 => "#,##0.00_);[Red](#,##0.00)",
        41 => "_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"_);_(@_)",
        42 => "_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"_);_(@_)",
        43 => "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)",
        44 => "_(\"$\"* #,##0.00_)_(\"$\"* \\(#,##0.00\\)_(\"$\"* \"-\"??_)_(@_)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mmss.0",
        48 => "##0.0E+0",
        49 => "@",
        _ => null,
    };
}

internal sealed class ExcelStyleBook
{
    private static readonly string[] SectionOrder =
    [
        "numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs",
        "cellStyles", "dxfs", "tableStyles", "colors", "extLst",
    ];

    private readonly Stylesheet _stylesheet;
    private readonly Fonts _fonts;
    private readonly Fills _fills;
    private readonly Borders _borders;
    private readonly CellFormats _cellFormats;
    private readonly CellStyleFormats _cellStyleFormats;
    private readonly CellStyles _cellStyles;
    private readonly DifferentialFormats _dxfs;
    private readonly TableStyles _tableStyles;
    private NumberingFormats? _numberingFormats;

    private readonly Dictionary<string, int> _fontCache = new(StringComparer.Ordinal);
    private readonly Dictionary<string, int> _fillCache = new(StringComparer.Ordinal);
    private readonly Dictionary<string, int> _borderCache = new(StringComparer.Ordinal);
    private readonly Dictionary<string, int> _numFmtCache = new(StringComparer.Ordinal);
    private readonly Dictionary<string, int> _xfCache = new(StringComparer.Ordinal);
    private readonly Dictionary<string, int> _dxfCache = new(StringComparer.Ordinal);
    private readonly Dictionary<int, int> _hyperlinkCache = new();

    public ExcelStyleBook(Stylesheet stylesheet)
    {
        _stylesheet = stylesheet;
        _fonts = EnsureSection(stylesheet, "fonts", () => new Fonts(
            new Font(
                new FontSize { Val = 11 },
                new FontName { Val = "Calibri" },
                new FontFamilyNumbering { Val = 2 },
                new FontScheme { Val = FontSchemeValues.Minor })) { Count = 1 });
        _fills = EnsureSection(stylesheet, "fills", () => new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 })) { Count = 2 });
        _borders = EnsureSection(stylesheet, "borders", () => new Borders(
            new Border(
                new LeftBorder(), new RightBorder(), new TopBorder(),
                new BottomBorder(), new DiagonalBorder())) { Count = 1 });
        _cellStyleFormats = EnsureSection(stylesheet, "cellStyleXfs", () => new CellStyleFormats(
            new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 }) { Count = 1 });
        _cellFormats = EnsureSection(stylesheet, "cellXfs", () => new CellFormats(
            new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 }) { Count = 1 });
        _cellStyles = EnsureSection(stylesheet, "cellStyles", () => new CellStyles(
            new S.CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 }) { Count = 1 });
        _dxfs = EnsureSection(stylesheet, "dxfs", () => new DifferentialFormats { Count = 0 });
        _tableStyles = EnsureSection(stylesheet, "tableStyles", () => new TableStyles
        {
            Count = 0,
            DefaultTableStyle = "TableStyleMedium2",
            DefaultPivotStyle = "PivotStyleLight16",
        });
    }

    public static Stylesheet CreateDefaultStylesheet() => new(
        new Fonts(
            new Font(
                new FontSize { Val = 11 },
                new FontName { Val = "Calibri" },
                new FontFamilyNumbering { Val = 2 },
                new FontScheme { Val = FontSchemeValues.Minor })) { Count = 1 },
        new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 })) { Count = 2 },
        new Borders(
            new Border(
                new LeftBorder(), new RightBorder(), new TopBorder(),
                new BottomBorder(), new DiagonalBorder())) { Count = 1 },
        new CellStyleFormats(
            new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 }) { Count = 1 },
        new CellFormats(
            new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 }) { Count = 1 },
        new CellStyles(
            new S.CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 }) { Count = 1 },
        new DifferentialFormats { Count = 0 },
        new TableStyles
        {
            Count = 0,
            DefaultTableStyle = "TableStyleMedium2",
            DefaultPivotStyle = "PivotStyleLight16",
        });

    public Stylesheet Stylesheet => _stylesheet;

    public int GetStyle(ExcelCellStyle? style)
    {
        if (style is null)
            return 0;
        var fontId = GetOrCreateFont(style);
        var fillId = GetOrCreateFill(style.BgColor);
        var borderId = GetOrCreateBorder(style.Border);
        var numFmtId = string.IsNullOrEmpty(style.NumberFormat)
            ? 0
            : GetOrCreateNumberFormat(style.NumberFormat);
        var alignment = BuildAlignment(style);
        if (fontId == 0 && fillId == 0 && borderId == 0 && numFmtId == 0 && alignment is null)
            return 0;
        return GetOrCreateCellFormat(fontId, fillId, borderId, numFmtId, alignment);
    }

    /// <summary>表头未给 header_style 时的默认加粗（对齐 openpyxl Font(bold=True)）。</summary>
    public int GetBoldHeaderStyle()
    {
        const string key = "header-default-bold";
        if (_xfCache.TryGetValue(key, out var cached))
            return cached;
        var font = new Font(new Bold());
        var fontId = _fonts.ChildElements.Count;
        _fonts.Append(font);
        var styleId = GetOrCreateCellFormat(fontId, 0, 0, 0, null);
        _xfCache[key] = styleId;
        return styleId;
    }

    public int GetHyperlinkStyle(int sourceStyleIndex)
    {
        if (_hyperlinkCache.TryGetValue(sourceStyleIndex, out var cached))
            return cached;

        var source = _cellFormats.Elements<CellFormat>().ElementAtOrDefault(sourceStyleIndex)
                     ?? new CellFormat
                     {
                         NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0,
                     };
        var sourceFontId = source.FontId?.Value ?? 0;
        var sourceFont = _fonts.Elements<Font>().ElementAtOrDefault((int)sourceFontId);

        var font = new Font();
        if (sourceFont?.Bold is { } bold && (bold.Val is null || bold.Val.Value))
            font.Append(new Bold());
        if (sourceFont?.Italic is { } italic && (italic.Val is null || italic.Val.Value))
            font.Append(new Italic());
        font.Append(new Underline { Val = UnderlineValues.Single });
        if (sourceFont?.FontSize?.Val is { } size)
            font.Append(new FontSize { Val = size });
        font.Append(new Color { Rgb = "FF0563C1" });
        if (sourceFont?.FontName?.Val is { } name)
            font.Append(new FontName { Val = name });
        if (sourceFont?.FontFamilyNumbering?.Val is { } family)
            font.Append(new FontFamilyNumbering { Val = family });
        if (sourceFont?.FontScheme?.Val is { } scheme)
            font.Append(new FontScheme { Val = scheme });

        var fontId = _fonts.ChildElements.Count;
        _fonts.Append(font);

        var xf = (CellFormat)source.CloneNode(true);
        xf.FontId = (uint)fontId;
        xf.ApplyFont = true;
        var newIndex = _cellFormats.ChildElements.Count;
        _cellFormats.Append(xf);
        _hyperlinkCache[sourceStyleIndex] = newIndex;
        return newIndex;
    }

    public int AddDifferentialFormat(string? fillColor, string? fontColor)
    {
        var key = $"{fillColor}|{fontColor}";
        if (_dxfCache.TryGetValue(key, out var cached))
            return cached;

        var elements = new List<OpenXmlElement>();
        if (!string.IsNullOrEmpty(fontColor))
            elements.Add(new Font(new Color { Rgb = fontColor }));
        if (!string.IsNullOrEmpty(fillColor))
            elements.Add(new Fill(new PatternFill(
                new ForegroundColor { Rgb = fillColor },
                new BackgroundColor { Rgb = fillColor })
            {
                PatternType = PatternValues.Solid,
            }));
        var dxf = elements.Count == 0
            ? new DifferentialFormat()
            : new DifferentialFormat(elements);
        var index = _dxfs.ChildElements.Count;
        _dxfs.Append(dxf);
        _dxfCache[key] = index;
        return index;
    }

    public void RefreshCounts()
    {
        if (_stylesheet.NumberingFormats is { } numbering)
            numbering.Count = (uint)numbering.ChildElements.Count;
        _fonts.Count = (uint)_fonts.ChildElements.Count;
        _fills.Count = (uint)_fills.ChildElements.Count;
        _borders.Count = (uint)_borders.ChildElements.Count;
        _cellStyleFormats.Count = (uint)_cellStyleFormats.ChildElements.Count;
        _cellFormats.Count = (uint)_cellFormats.ChildElements.Count;
        _cellStyles.Count = (uint)_cellStyles.ChildElements.Count;
        _dxfs.Count = (uint)_dxfs.ChildElements.Count;
        _tableStyles.Count = (uint)_tableStyles.ChildElements.Count;
    }

    // ── 内部构件 ─────────────────────────────────────────────────

    private int GetOrCreateFont(ExcelCellStyle style)
    {
        var hasFont = style.Bold == true || style.Italic == true
                      || style.FontColor is not null || style.FontSize is not null
                      || style.FontName is not null;
        if (!hasFont)
            return 0;

        var size = style.FontSize ?? 11;
        var key = string.Create(CultureInfo.InvariantCulture,
            $"{style.Bold}|{style.Italic}|{size}|{style.FontColor}|{style.FontName}");
        if (_fontCache.TryGetValue(key, out var cached))
            return cached;

        var font = new Font();
        if (style.Bold == true)
            font.Append(new Bold());
        if (style.Italic == true)
            font.Append(new Italic());
        font.Append(new FontSize { Val = size });
        if (style.FontColor is not null)
            font.Append(new Color { Rgb = style.FontColor });
        if (style.FontName is not null)
            font.Append(new FontName { Val = style.FontName });

        var id = _fonts.ChildElements.Count;
        _fonts.Append(font);
        _fontCache[key] = id;
        return id;
    }

    private int GetOrCreateFill(string? bgColor)
    {
        if (string.IsNullOrEmpty(bgColor))
            return 0;
        if (_fillCache.TryGetValue(bgColor, out var cached))
            return cached;
        var fill = new Fill(new PatternFill(
            new ForegroundColor { Rgb = bgColor },
            new BackgroundColor { Rgb = bgColor })
        {
            PatternType = PatternValues.Solid,
        });
        var id = _fills.ChildElements.Count;
        _fills.Append(fill);
        _fillCache[bgColor] = id;
        return id;
    }

    private int GetOrCreateBorder(ExcelBorderSpec? spec)
    {
        if (spec is null)
            return 0;
        var key = $"{spec.Style}|{spec.Color}";
        if (_borderCache.TryGetValue(key, out var cached))
            return cached;
        var style = spec.Style switch
        {
            "thin" => BorderStyleValues.Thin,
            "medium" => BorderStyleValues.Medium,
            "thick" => BorderStyleValues.Thick,
            "dashed" => BorderStyleValues.Dashed,
            "dotted" => BorderStyleValues.Dotted,
            "double" => BorderStyleValues.Double,
            "hair" => BorderStyleValues.Hair,
            _ => throw new NativeUsageException($"不支持的边框样式: {spec.Style}"),
        };
        var border = new Border(
            MakeSide<LeftBorder>(style, spec.Color),
            MakeSide<RightBorder>(style, spec.Color),
            MakeSide<TopBorder>(style, spec.Color),
            MakeSide<BottomBorder>(style, spec.Color),
            new DiagonalBorder());
        var id = _borders.ChildElements.Count;
        _borders.Append(border);
        _borderCache[key] = id;
        return id;
    }

    private static T MakeSide<T>(BorderStyleValues style, string? color)
        where T : BorderPropertiesType, new()
    {
        var side = new T { Style = style };
        if (!string.IsNullOrEmpty(color))
            side.Append(new Color { Rgb = color });
        return side;
    }

    private int GetOrCreateNumberFormat(string formatCode)
    {
        if (_numFmtCache.TryGetValue(formatCode, out var cached))
            return cached;
        _numberingFormats ??= EnsureSection(_stylesheet, "numFmts", () => new NumberingFormats());
        var start = 164;
        foreach (var existing in _numberingFormats.Elements<NumberingFormat>())
        {
            if (existing.NumberFormatId?.Value is { } id && id >= start)
                start = (int)id + 1;
        }
        var formatId = start;
        _numberingFormats.Append(new NumberingFormat
        {
            NumberFormatId = (uint)formatId,
            FormatCode = formatCode,
        });
        _numFmtCache[formatCode] = formatId;
        return formatId;
    }

    private Alignment? BuildAlignment(ExcelCellStyle style)
    {
        if (style.Alignment is null && style.VerticalAlignment is null && style.WrapText != true)
            return null;
        var alignment = new Alignment();
        if (style.Alignment is not null)
        {
            alignment.Horizontal = style.Alignment switch
            {
                "left" => HorizontalAlignmentValues.Left,
                "center" => HorizontalAlignmentValues.Center,
                "right" => HorizontalAlignmentValues.Right,
                _ => throw new NativeUsageException($"不支持的水平对齐: {style.Alignment}"),
            };
        }
        if (style.VerticalAlignment is not null)
        {
            alignment.Vertical = style.VerticalAlignment switch
            {
                "top" => VerticalAlignmentValues.Top,
                "center" => VerticalAlignmentValues.Center,
                "bottom" => VerticalAlignmentValues.Bottom,
                _ => throw new NativeUsageException($"不支持的垂直对齐: {style.VerticalAlignment}"),
            };
        }
        if (style.WrapText == true)
            alignment.WrapText = true;
        return alignment;
    }

    private int GetOrCreateCellFormat(
        int fontId, int fillId, int borderId, int numFmtId, Alignment? alignment)
    {
        var alignKey = alignment is null
            ? ""
            : $"{alignment.Horizontal}|{alignment.Vertical}|{alignment.WrapText}";
        var key = $"{fontId}|{fillId}|{borderId}|{numFmtId}|{alignKey}";
        if (_xfCache.TryGetValue(key, out var cached))
            return cached;

        var xf = new CellFormat
        {
            NumberFormatId = (uint)numFmtId,
            FontId = (uint)fontId,
            FillId = (uint)fillId,
            BorderId = (uint)borderId,
            FormatId = 0,
            ApplyNumberFormat = numFmtId > 0,
            ApplyFont = fontId > 0,
            ApplyFill = fillId > 0,
            ApplyBorder = borderId > 0,
            ApplyAlignment = alignment is not null,
        };
        if (alignment is not null)
            xf.Append(alignment);
        var id = _cellFormats.ChildElements.Count;
        _cellFormats.Append(xf);
        _xfCache[key] = id;
        return id;
    }

    private static T EnsureSection<T>(Stylesheet stylesheet, string localName, Func<T> factory)
        where T : OpenXmlElement
    {
        var existing = stylesheet.ChildElements.FirstOrDefault(e => e.LocalName == localName);
        if (existing is not null)
        {
            if (existing is T typed)
                return typed;
            throw new NativeRuntimeException($"样式表节点 {localName} 类型异常: {existing.GetType().Name}");
        }

        var created = factory();
        var targetOrder = Array.IndexOf(SectionOrder, localName);
        var anchor = stylesheet.ChildElements.FirstOrDefault(e =>
        {
            var order = Array.IndexOf(SectionOrder, e.LocalName);
            return order > targetOrder;
        });
        if (anchor is null)
            stylesheet.AppendChild(created);
        else
            stylesheet.InsertBefore(created, anchor);
        return created;
    }
}
