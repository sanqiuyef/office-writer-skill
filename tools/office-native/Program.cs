using System.Text.Json;
using OfficeNative;

// office-native：office-writer 的可选原生后端。
// 输出契约与 Python 入口一致：stdout 只输出一个 JSON，退出码 0/2/3。
Console.OutputEncoding = System.Text.Encoding.UTF8;
Console.InputEncoding = System.Text.Encoding.UTF8;

if (args.Length == 0)
{
    Console.Error.WriteLine(
        "usage: office-native <command> --params '<json>' | --params-file <path>");
    return 2;
}

var command = args[0].ToLowerInvariant();
JsonDocument? document = null;
try
{
    document = LoadParams(args);
    var (tool, data) = CommandRouter.Dispatch(command, document.RootElement);
    Emit(tool, data);
    // validate-word-document 的门控语义：errors>0 → 退出码 1（数据仍正常输出）。
    if (command is "validate-word-document"
        && data.TryGetValue("errors", out var errors)
        && Convert.ToInt32(errors) > 0)
        return 1;
    return 0;
}
catch (NativeUsageException ex)
{
    EmitError(command, ex.Message, 2);
    return 2;
}
catch (Exception ex)
{
    var detail = Environment.GetEnvironmentVariable("OFFICE_NATIVE_DEBUG") == "1"
        ? ex.ToString()
        : ex is NativeRuntimeException
            ? ex.Message
            : $"{ex.GetType().Name}: {ex.Message}";
    EmitError(command, detail, 3);
    return 3;
}
finally
{
    document?.Dispose();
}

static JsonDocument LoadParams(string[] args)
{
    string? inline = null;
    string? file = null;
    for (var i = 1; i < args.Length; i++)
    {
        switch (args[i])
        {
            case "--params" when i + 1 < args.Length:
                inline = args[++i];
                break;
            case "--params-file" when i + 1 < args.Length:
                file = args[++i];
                break;
            case "--params-file" when i + 1 < args.Length && args[i + 1] == "-":
                file = "-";
                i++;
                break;
            default:
                throw new NativeUsageException($"未知参数: {args[i]}");
        }
    }

    string json;
    if (inline is not null)
        json = inline;
    else if (file == "-")
        json = Console.In.ReadToEnd();
    else if (file is not null)
    {
        if (!File.Exists(file))
            throw new NativeRuntimeException($"参数文件不存在: {file}");
        json = File.ReadAllText(file, System.Text.Encoding.UTF8);
    }
    else
        json = "{}";

    try
    {
        return JsonDocument.Parse(json);
    }
    catch (JsonException ex)
    {
        throw new NativeUsageException($"参数不是合法 JSON: {ex.Message}");
    }
}

static void Emit(string tool, object data)
{
    var payload = new Dictionary<string, object?>
    {
        ["ok"] = true,
        ["tool"] = tool,
        ["data"] = data,
    };
    Console.Out.WriteLine(JsonSerializer.Serialize(payload, NativeCommon.JsonOptions));
}

static void EmitError(string tool, string message, int exitCode)
{
    var payload = new Dictionary<string, object?>
    {
        ["ok"] = false,
        ["tool"] = tool,
        ["error"] = message,
    };
    Console.Out.WriteLine(JsonSerializer.Serialize(payload, NativeCommon.JsonOptions));
}
