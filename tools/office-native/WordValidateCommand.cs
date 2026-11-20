using System.Text.Json;

namespace OfficeNative;

/// <summary>validate-word-document 命令入口（CommandRouter 只转发参数）。</summary>
internal static class WordValidateCommand
{
    public static Dictionary<string, object?> RunValidate(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        return WordValidateChecks.Validate(filePath).ToDictionary();
    }
}
