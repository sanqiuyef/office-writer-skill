using System.Text;
using System.Text.Json;

namespace OfficeNative;

internal static class GenericTxtCommands
{
    private const string TruncationMarker = "_(内容已截断)_";

    public static Dictionary<string, object?> RunReadTxt(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        var maxChars = NativeCommon.OptInt(p, "max_chars", 50000);
        if (maxChars < 1)
            throw new NativeUsageException("max_chars 必须大于 0");
        NativeCommon.EnsureFile(filePath);

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var bytes = File.ReadAllBytes(filePath);
        var candidates = new List<(string Name, Encoding Encoding)>();
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
            candidates.Add(("utf-8-sig", new UTF8Encoding(false, true)));
        if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
            candidates.Add(("utf-16", new UnicodeEncoding(false, false, true)));
        if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
            candidates.Add(("utf-16-be", new UnicodeEncoding(true, false, true)));
        candidates.Add(("utf-8-sig", new UTF8Encoding(false, true)));
        candidates.Add(("utf-16", new UnicodeEncoding(false, false, true)));
        candidates.Add(("utf-16-be", new UnicodeEncoding(true, false, true)));
        candidates.Add(("gb18030", Encoding.GetEncoding(54936,
            EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback)));
        candidates.Add(("latin-1", Encoding.Latin1));

        foreach (var (name, encoding) in candidates)
        {
            string text;
            try
            {
                text = encoding.GetString(bytes);
            }
            catch (DecoderFallbackException)
            {
                continue;
            }
            if (text.Length > 0 && text[0] == '\uFEFF')
                text = text[1..];
            // Python 以 universal newlines 读取；统一成 LF 才能与 markdown 输出对齐。
            text = text.Replace("\r\n", "\n").Replace('\r', '\n');
            var (content, truncated) = TruncateByLine(text, maxChars);
            var message = $"成功读取文本文件（编码: {name}，{bytes.Length} 字节）";
            if (truncated)
                message += $"（已截断至 {maxChars} 字符）";
            var result = new Dictionary<string, object?>
            {
                ["success"] = true,
                ["file_path"] = filePath,
                ["content"] = content,
                ["encoding"] = name,
                ["file_size"] = bytes.Length,
                ["truncated"] = truncated,
                ["message"] = message,
            };
            return new Dictionary<string, object?>
            {
                ["markdown"] = BuildReadMarkdown(result),
                ["result"] = result,
            };
        }

        throw new NativeRuntimeException("无法识别文本文件编码");
    }

    private static string BuildReadMarkdown(Dictionary<string, object?> result)
    {
        var parts = new List<string> { (string)(result["content"] ?? string.Empty) };
        var stats = result
            .Where(item => item.Key is not ("success" or "file_path" or "content"
                or "message" or "truncated") && item.Value is not null)
            .Select(item => $"{item.Key}: {item.Value}");
        if (stats.Any())
            parts.Add("> " + string.Join("；", stats));
        if (result.TryGetValue("truncated", out var truncated) && truncated is true)
            parts.Add("> 内容过长已截断");
        return string.Join("\n\n", parts.Where(part => !string.IsNullOrEmpty(part)));
    }

    private static (string Content, bool Truncated) TruncateByLine(string text, int maxChars)
    {
        if (text.Length <= maxChars)
            return (text, false);
        var builder = new StringBuilder();
        var start = 0;
        while (start < text.Length)
        {
            var newline = text.IndexOf('\n', start);
            var end = newline < 0 ? text.Length : newline + 1;
            var line = text[start..end];
            if (builder.Length + line.Length > maxChars)
            {
                if (builder.Length + TruncationMarker.Length <= maxChars)
                    builder.Append(TruncationMarker);
                else if (builder.Length == 0)
                    builder.Append(TruncationMarker[..Math.Min(TruncationMarker.Length, maxChars)]);
                break;
            }
            builder.Append(line);
            start = end;
        }
        return (builder.ToString(), true);
    }
}
