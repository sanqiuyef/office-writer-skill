using System.Text.Json;

namespace OfficeNative;

internal static class MergeTemplateCommands
{
    public static Dictionary<string, object?> RunMerge(JsonElement p)
    {
        var templatePath = NativeCommon.RequireString(p, "template_path");
        NativeCommon.EnsureFile(templatePath);
        return Path.GetExtension(templatePath).ToLowerInvariant() switch
        {
            ".docx" or ".dotx" => WordFileCommands.RunMergeOfficeTemplate(p),
            ".xlsx" or ".xlsm" or ".xltx" or ".xltm" => ExcelFileCommands.RunMergeTemplate(p),
            ".pptx" or ".potx" => PptFileCommands.RunMergeTemplate(p),
            var extension => throw new NativeRuntimeException(
                $"不支持模板格式: {extension}（支持 .docx/.xlsx/.xlsm/.pptx/.potx）"),
        };
    }
}
