using System.Globalization;
using System.Text;
using System.Text.Json;

namespace OfficeNative;

/// <summary>批量编辑的可纠正错误：code/path/suggestion 与 Python OfficeOperationError 对齐。</summary>
internal sealed class BatchEditException : Exception
{
    public BatchEditException(
        string code,
        string message,
        string? path = null,
        string? suggestion = null,
        IEnumerable<string>? validValues = null)
        : base(message)
    {
        Code = code;
        Path = path;
        Suggestion = suggestion;
        ValidValues = validValues?.ToList();
    }

    public string Code { get; }
    public string? Path { get; }
    public string? Suggestion { get; }
    public IReadOnlyList<string>? ValidValues { get; }

    public string ToDisplay()
    {
        var builder = new StringBuilder();
        builder.Append('[').Append(Code).Append("] ").Append(Message);
        if (!string.IsNullOrEmpty(Path))
            builder.Append(" | path: ").Append(Path);
        if (!string.IsNullOrEmpty(Suggestion))
            builder.Append(" | suggestion: ").Append(Suggestion);
        if (ValidValues is { Count: > 0 })
            builder.Append(" | valid_values: ").Append(string.Join(", ", ValidValues));
        return builder.ToString();
    }
}

internal sealed class OfficeOperation
{
    public int Index { get; init; }
    public string Op { get; init; } = "";
    public string Path { get; init; } = "";
    public string? ElementType { get; init; }
    public Dictionary<string, JsonElement> Props { get; init; } =
        new(StringComparer.Ordinal);
    public string? TargetPath { get; init; }
    public string Position { get; init; } = "after";
}

internal static class BatchEditOperations
{
    private static readonly HashSet<string> AllowedOps = new(StringComparer.Ordinal)
    {
        "set", "add", "remove", "move", "replace",
    };

    public static List<OfficeOperation> Parse(JsonElement root)
    {
        if (root.ValueKind != JsonValueKind.Object
            || !root.TryGetProperty("operations", out var operations)
            || operations.ValueKind != JsonValueKind.Array)
            throw new NativeUsageException("缺少参数 operations（必须是数组）");
        if (operations.GetArrayLength() == 0)
            throw new BatchEditException("invalid_operation", "operations 不能为空");

        var parsed = new List<OfficeOperation>(operations.GetArrayLength());
        var index = 0;
        foreach (var item in operations.EnumerateArray())
        {
            var label = $"第 {index + 1} 个操作";
            if (item.ValueKind != JsonValueKind.Object)
                throw new NativeUsageException($"{label}必须是 JSON 对象");

            var op = RequiredString(item, "op", label);
            if (!AllowedOps.Contains(op))
                throw new NativeUsageException(
                    $"{label}的 op 必须是 set/add/remove/move/replace，收到 {op}");
            var path = RequiredString(item, "path", label);
            var elementType = OptionalString(item, "element_type", label);
            var targetPath = OptionalString(item, "target_path", label);
            var position = OptionalString(item, "position", label) ?? "after";
            if (position is not ("before" or "after"))
                throw new NativeUsageException($"{label}的 position 必须是 before/after");

            var props = new Dictionary<string, JsonElement>(StringComparer.Ordinal);
            if (item.TryGetProperty("props", out var propsElement)
                && propsElement.ValueKind != JsonValueKind.Null)
            {
                if (propsElement.ValueKind != JsonValueKind.Object)
                    throw new NativeUsageException($"{label}的 props 必须是 JSON 对象");
                foreach (var prop in propsElement.EnumerateObject())
                    props[prop.Name] = prop.Value;
            }

            parsed.Add(new OfficeOperation
            {
                Index = index,
                Op = op,
                Path = path,
                ElementType = elementType,
                Props = props,
                TargetPath = targetPath,
                Position = position,
            });
            index++;
        }
        return parsed;
    }

    private static string RequiredString(JsonElement item, string name, string label)
    {
        var value = OptionalString(item, name, label);
        if (string.IsNullOrEmpty(value))
            throw new NativeUsageException($"{label}缺少字符串参数 {name}");
        return value;
    }

    private static string? OptionalString(JsonElement item, string name, string label)
    {
        if (!item.TryGetProperty(name, out var value) || value.ValueKind == JsonValueKind.Null)
            return null;
        if (value.ValueKind != JsonValueKind.String)
            throw new NativeUsageException($"{label}的 {name} 必须是字符串");
        return value.GetString();
    }
}

/// <summary>批量编辑共用的 props 读取与类型转换。</summary>
internal static class BatchEditJson
{
    public static bool Has(OfficeOperation op, string name) =>
        op.Props.ContainsKey(name) && op.Props[name].ValueKind != JsonValueKind.Null;

    public static bool TryGet(OfficeOperation op, string name, out JsonElement value)
    {
        if (op.Props.TryGetValue(name, out value) && value.ValueKind != JsonValueKind.Null)
            return true;
        value = default;
        return false;
    }

    public static string? OptString(OfficeOperation op, string name)
    {
        if (!TryGet(op, name, out var value))
            return null;
        if (value.ValueKind != JsonValueKind.String)
            throw new BatchEditException(
                "invalid_value", $"属性 {name} 必须是字符串", op.Path);
        return value.GetString();
    }

    /// <summary>Python str(props.get(name, fallback))：null → "None"。</summary>
    public static string PythonStringProp(OfficeOperation op, string name, string fallback = "")
    {
        if (!op.Props.TryGetValue(name, out var value))
            return fallback;
        return PythonString(value);
    }

    public static string PythonString(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.String => value.GetString() ?? "",
        JsonValueKind.Number => value.GetRawText(),
        JsonValueKind.True => "True",
        JsonValueKind.False => "False",
        JsonValueKind.Null => "None",
        JsonValueKind.Array => "[" + string.Join(", ",
            value.EnumerateArray().Select(PythonString)) + "]",
        JsonValueKind.Object => "{...}",
        _ => value.ToString(),
    };

    public static double RequireDouble(OfficeOperation op, string name)
    {
        if (!TryGet(op, name, out var value))
            throw new BatchEditException(
                "invalid_value", $"缺少数值属性 {name}", op.Path);
        return ToDoubleValue(value, name, op.Path);
    }

    public static double OptDouble(OfficeOperation op, string name, double fallback) =>
        TryGet(op, name, out var value) ? ToDoubleValue(value, name, op.Path) : fallback;

    public static double ToDoubleValue(JsonElement value, string name, string path)
    {
        if (value.ValueKind == JsonValueKind.Number)
        {
            try { return value.GetDouble(); }
            catch (FormatException) { /* fall through */ }
        }
        if (value.ValueKind == JsonValueKind.String
            && double.TryParse(value.GetString(), NumberStyles.Float,
                CultureInfo.InvariantCulture, out var parsed))
            return parsed;
        throw new BatchEditException(
            "invalid_value", $"属性 {name} 必须是数字", path);
    }

    public static int RequireInt(OfficeOperation op, string name)
    {
        if (!TryGet(op, name, out var value))
            throw new BatchEditException(
                "invalid_value", $"缺少整数属性 {name}", op.Path);
        return ToIntValue(value, name, op.Path);
    }

    public static int OptInt(OfficeOperation op, string name, int fallback) =>
        TryGet(op, name, out var value) ? ToIntValue(value, name, op.Path) : fallback;

    public static int ToIntValue(JsonElement value, string name, string path)
    {
        if (value.ValueKind == JsonValueKind.Number)
        {
            if (value.TryGetInt32(out var integer))
                return integer;
            try { return (int)value.GetDouble(); }
            catch (FormatException) { /* fall through */ }
        }
        if (value.ValueKind == JsonValueKind.String
            && int.TryParse(value.GetString(), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out var parsed))
            return parsed;
        throw new BatchEditException(
            "invalid_value", $"属性 {name} 必须是整数", path);
    }

    /// <summary>Python 真值语义，用于 bold/italic 等布尔属性。</summary>
    public static bool Truthy(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.True => true,
        JsonValueKind.False => false,
        JsonValueKind.Null or JsonValueKind.Undefined => false,
        JsonValueKind.String => (value.GetString() ?? "").Length > 0,
        JsonValueKind.Number => value.GetDouble() != 0,
        JsonValueKind.Array => value.GetArrayLength() > 0,
        JsonValueKind.Object => value.EnumerateObject().Any(),
        _ => false,
    };

    public static bool OptTruthy(OfficeOperation op, string name, bool fallback = false) =>
        TryGet(op, name, out var value) ? Truthy(value) : fallback;

    /// <summary>校验并返回不带 # 的 6 位大写 RGB。</summary>
    public static string RequireHexColor(OfficeOperation op, string name, JsonElement value) =>
        RequireHexColorValue(value, name, op.Path);

    public static string RequireHexColorValue(
        JsonElement value, string name, string path)
    {
        if (value.ValueKind != JsonValueKind.String)
            throw new BatchEditException(
                "invalid_value", $"属性 {name} 必须是 6 位十六进制颜色", path,
                validValues: ["RRGGBB"]);
        var text = (value.GetString() ?? "").Trim().TrimStart('#');
        if (text.Length != 6 || text.Any(character => !Uri.IsHexDigit(character)))
            throw new BatchEditException(
                "invalid_value", $"无效的颜色值: {value.GetString()}", path,
                validValues: ["RRGGBB"]);
        return text.ToUpperInvariant();
    }

    public static string HexColorProp(OfficeOperation op, string name)
    {
        if (!TryGet(op, name, out var value))
            throw new BatchEditException("invalid_value", $"缺少颜色属性 {name}", op.Path);
        return RequireHexColor(op, name, value);
    }

    public static void EnsureAllowed(
        OfficeOperation op, IEnumerable<string> allowed, string label)
    {
        var allowedSet = allowed.ToHashSet(StringComparer.Ordinal);
        var unknown = op.Props.Keys
            .Where(key => !allowedSet.Contains(key))
            .OrderBy(key => key, StringComparer.Ordinal)
            .ToList();
        if (unknown.Count == 0)
            return;
        var displayed = "[" + string.Join(", ", unknown.Select(key => $"'{key}'")) + "]";
        throw new BatchEditException(
            "unsupported_property", $"{label}不支持属性: {displayed}", op.Path,
            validValues: allowedSet.OrderBy(key => key, StringComparer.Ordinal));
    }

    public static string RequireElementType(OfficeOperation op, string label,
        IEnumerable<string>? validValues = null)
    {
        if (!string.IsNullOrEmpty(op.ElementType))
            return op.ElementType;
        throw new BatchEditException(
            "invalid_value", $"{label} add 需要 element_type", op.Path,
            validValues: validValues);
    }

    public static JsonElement RequireArrayProp(OfficeOperation op, string name)
    {
        if (!TryGet(op, name, out var value) || value.ValueKind != JsonValueKind.Array)
            throw new BatchEditException(
                "invalid_value", $"属性 {name} 必须是数组", op.Path);
        return value;
    }

    public static List<string> ArrayStrings(JsonElement array, string name, string path)
    {
        var result = new List<string>(array.GetArrayLength());
        foreach (var item in array.EnumerateArray())
        {
            if (item.ValueKind == JsonValueKind.Array
                || item.ValueKind == JsonValueKind.Object)
                throw new BatchEditException(
                    "invalid_value", $"{name} 的元素必须是标量（字符串/数字/布尔/null）", path);
            result.Add(PythonString(item));
        }
        return result;
    }
}
