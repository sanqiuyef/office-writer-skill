using System.Runtime.Versioning;
using System.Text.RegularExpressions;
using Microsoft.Win32;

namespace OfficeNative;

/// <summary>
/// 本机字体集合；端口 office_convert._font_key / _split_registry_font_name /
/// _font_names_from_registry / _font_names_from_dirs / installed_font_keys。
/// 只用注册表与字体目录，不依赖 System.Drawing。
/// </summary>
internal static class WordValidateFonts
{
    private const string RegistryKeyPath =
        @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts";

    private static readonly object Gate = new();
    private static HashSet<string>? _installedKeys;

    // Windows 字体文件名与注册表显示名不一致时的补充映射。
    private static readonly Dictionary<string, string> FontFileNames =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["consola"] = "Consolas",
            ["consolab"] = "Consolas",
            ["consolai"] = "Consolas",
            ["consolaz"] = "Consolas",
            ["cascadiamono"] = "Cascadia Mono",
            ["cascadiacode"] = "Cascadia Code",
            ["dejavusansmono"] = "DejaVu Sans Mono",
            ["lucon"] = "Lucida Console",
            ["simsun"] = "SimSun",
            ["simsunb"] = "SimSun-ExtB",
            ["simsunextg"] = "SimSun-ExtG",
            ["simhei"] = "SimHei",
            ["simkai"] = "KaiTi",
            ["simfang"] = "FangSong",
            ["simli"] = "LiSu",
            ["simyou"] = "YouYuan",
            ["msyh"] = "Microsoft YaHei",
            ["msyhbd"] = "Microsoft YaHei Bold",
            ["msyhl"] = "Microsoft YaHei Light",
            ["deng"] = "DengXian",
            ["dengb"] = "DengXian Bold",
            ["dengl"] = "DengXian Light",
            ["stsong"] = "STSong",
            ["stfangsong"] = "STFangsong",
            ["stkaiti"] = "STKaiti",
            ["stxihei"] = "STXihei",
            ["stzhongs"] = "STZhongsong",
            ["stxingka"] = "STXingkai",
        };

    // 中文 Word 常用字体名与 Windows 注册表英文名的对应关系。
    private static readonly (string Chinese, string English)[] FontAliases =
    [
        ("宋体", "SimSun"),
        ("新宋体", "NSimSun"),
        ("黑体", "SimHei"),
        ("微软雅黑", "Microsoft YaHei"),
        ("微软雅黑 Light", "Microsoft YaHei Light"),
        ("等线", "DengXian"),
        ("等线 Light", "DengXian Light"),
        ("楷体", "KaiTi"),
        ("楷体_GB2312", "KaiTi_GB2312"),
        ("仿宋", "FangSong"),
        ("仿宋_GB2312", "FangSong_GB2312"),
        ("隶书", "LiSu"),
        ("幼圆", "YouYuan"),
        ("华文宋体", "STSong"),
        ("华文仿宋", "STFangsong"),
        ("华文楷体", "STKaiti"),
        ("华文细黑", "STXihei"),
        ("华文中宋", "STZhongsong"),
        ("华文行楷", "STXingkai"),
    ];

    /// <summary>Normalise a font name for comparison (case/spacing/parentheses insensitive).</summary>
    public static string Key(string name)
    {
        var withoutParens = Regex.Replace(name ?? "", @"\([^)]*\)", "");
        var lowered = withoutParens.ToLowerInvariant();
        return Regex.Replace(lowered, @"[^0-9a-z\u3400-\u9fff\uf900-\ufaff]+", "");
    }

    /// <summary>Normalised names of fonts installed on this machine (cached).</summary>
    public static HashSet<string> InstalledKeys()
    {
        lock (Gate)
            return _installedKeys ??= BuildInstalledKeys();
    }

    private static HashSet<string> BuildInstalledKeys()
    {
        var names = new HashSet<string>(StringComparer.Ordinal);
        if (OperatingSystem.IsWindows())
            names.UnionWith(NamesFromRegistry());
        names.UnionWith(NamesFromDirectories());

        var keys = new HashSet<string>(StringComparer.Ordinal);
        foreach (var name in names)
        {
            var key = Key(name);
            if (key.Length > 0)
                keys.Add(key);
        }
        foreach (var (chinese, english) in FontAliases)
        {
            var chineseKey = Key(chinese);
            var englishKey = Key(english);
            if (keys.Contains(chineseKey) || keys.Contains(englishKey))
            {
                keys.Add(chineseKey);
                keys.Add(englishKey);
            }
        }
        return keys;
    }

    [SupportedOSPlatform("windows")]
    private static HashSet<string> NamesFromRegistry()
    {
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (var hive in new[] { RegistryHive.LocalMachine, RegistryHive.CurrentUser })
        {
            try
            {
                using var baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Default);
                using var key = baseKey.OpenSubKey(RegistryKeyPath);
                if (key is null)
                    continue;
                foreach (var rawName in key.GetValueNames())
                {
                    foreach (var name in SplitRegistryFontName(rawName))
                        names.Add(name);

                    var value = key.GetValue(rawName)?.ToString();
                    if (string.IsNullOrEmpty(value))
                        continue;
                    var stem = Path.GetFileNameWithoutExtension(value);
                    if (string.IsNullOrEmpty(stem))
                        continue;
                    names.Add(FontFileNames.TryGetValue(stem, out var mapped) ? mapped : stem);
                }
            }
            catch (Exception)
            {
                // 与 Python 的 OSError 容错一致：某个 hive 读不到就继续。
            }
        }
        return names;
    }

    /// <summary>Split combined registry names such as ``SimSun &amp; NSimSun (TrueType)``.</summary>
    private static IEnumerable<string> SplitRegistryFontName(string name)
    {
        var cleaned = Regex.Replace(name ?? "", @"\([^)]*\)\s*$", "").Trim();
        var parts = Regex.Split(cleaned, @"\s*(?:&|,|，|、|\band\b)\s*");
        var result = new List<string>();
        if (cleaned.Length > 0)
            result.Add(cleaned);
        foreach (var part in parts)
        {
            if (part.Length > 0 && !result.Contains(part))
                result.Add(part);
        }
        return result;
    }

    private static HashSet<string> NamesFromDirectories()
    {
        var directories = new List<string>();
        var systemRoot = Environment.GetEnvironmentVariable("WINDIR")
                         ?? Environment.GetEnvironmentVariable("SystemRoot");
        if (!string.IsNullOrEmpty(systemRoot))
            directories.Add(Path.Combine(systemRoot, "Fonts"));
        var localAppData = Environment.GetEnvironmentVariable("LOCALAPPDATA");
        if (!string.IsNullOrEmpty(localAppData))
            directories.Add(Path.Combine(localAppData, "Microsoft", "Windows", "Fonts"));

        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (var directory in directories)
        {
            if (!Directory.Exists(directory))
                continue;
            try
            {
                foreach (var entry in Directory.EnumerateFiles(directory))
                {
                    var extension = Path.GetExtension(entry).ToLowerInvariant();
                    if (extension is not (".ttf" or ".ttc" or ".otf" or ".tte"))
                        continue;
                    var stem = Path.GetFileNameWithoutExtension(entry);
                    if (string.IsNullOrEmpty(stem))
                        continue;
                    names.Add(FontFileNames.TryGetValue(stem, out var mapped) ? mapped : stem);
                }
            }
            catch (Exception)
            {
                // 目录不可读时跳过，与 Python 的 OSError 容错一致。
            }
        }
        return names;
    }
}
