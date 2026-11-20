using System.Text.Json;

namespace OfficeNative;

/// <summary>
/// edit-open-word-document：向已在 Word/WPS 中打开的文档追加内容块并保存。
/// 对应 Python office_com.try_edit_open_word_document。
/// </summary>
internal static class OpenEditWordCommands
{
    private static readonly HashSet<string> SupportedTypes = new(StringComparer.Ordinal)
    {
        "heading", "paragraph", "bullet_list", "numbered_list", "page_break",
    };

    public static Dictionary<string, object?> RunEditOpenWordDocument(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        var elements = OpenEditCommon.RequireArray(p, "append_sections");
        var sections = elements.EnumerateArray().Select(SectionItem.Parse).ToList();
        var defaultFont = FontConfig.ParseOptional(p, "default_font");
        NativeCommon.EnsureFileLight(filePath);

        var app = OpenEditCommon.FindWord(filePath, out var document, out var progId)
            ?? throw new NativeRuntimeException(
                $"目标文件未在活动 Word/WPS 实例中打开: {filePath}；请先打开文件");
        try
        {
            if (document is null)
                throw new NativeRuntimeException($"目标文件未在活动 Word/WPS 实例中打开: {filePath}");
            if ((bool)document.ReadOnly)
                throw new NativeRuntimeException("活动 Word 文档为只读，无法保存");

            var unsupported = sections
                .Where(section => !SupportedTypes.Contains(section.Type))
                .Select(section => section.Type)
                .Distinct(StringComparer.Ordinal)
                .Select(type => $"'{type}'")
                .ToList();
            if (unsupported.Count > 0)
                throw new NativeRuntimeException(
                    $"活动 Word 文档暂不支持 {string.Join(", ", unsupported)}；"
                    + "请关闭文件后使用原 writer");

            dynamic? undo = null;
            try
            {
                undo = app.UndoRecord;
                if (undo is not null)
                    undo.StartCustomRecord("mcp-office-writer");
            }
            catch { undo = null; }

            try
            {
                foreach (var section in sections)
                    AppendWordItem(document, section, defaultFont);
                document.Save();
            }
            catch
            {
                try
                {
                    if (undo is not null)
                    {
                        undo.EndCustomRecord();
                        app.Undo();
                    }
                }
                catch { /* 部分 Office 变体不提供撤销记录，保持原始异常 */ }
                throw;
            }
            try { if (undo is not null) undo.EndCustomRecord(); } catch { }

            return OpenEditCommon.EditResult(filePath, sections.Count,
                $"已通过 {progId} 修改并保存活动 Word 文档");
        }
        finally
        {
            NativeCommon.ReleaseCom(document);
            NativeCommon.ReleaseCom(app);
        }
    }

    private static void AppendWordItem(
        dynamic document, SectionItem item, FontConfig? defaultFont)
    {
        if (item.Type == "page_break")
        {
            dynamic pageBreak = AppendParagraph(document, "");
            pageBreak.InsertBreak(7); // wdPageBreak
            NativeCommon.ReleaseCom(pageBreak);
            return;
        }

        if (item.Type is "bullet_list" or "numbered_list")
        {
            foreach (var value in item.Items)
            {
                dynamic insertion = AppendParagraph(document, value);
                try
                {
                    ApplyWordStyle(insertion, item.Style, defaultFont);
                    if (item.Type == "bullet_list")
                        insertion.ListFormat.ApplyBulletDefault();
                    else
                        insertion.ListFormat.ApplyNumberDefault();
                }
                finally
                {
                    NativeCommon.ReleaseCom(insertion);
                }
            }
            return;
        }

        dynamic paragraph = AppendParagraph(document, item.Text ?? "");
        try
        {
            if (item.Type == "heading")
            {
                try
                {
                    paragraph.Style = $"Heading {item.Level ?? 1}";
                }
                catch
                {
                    paragraph.Font.Bold = true;
                }
            }
            ApplyWordStyle(paragraph, item.Style, defaultFont);
        }
        finally
        {
            NativeCommon.ReleaseCom(paragraph);
        }
    }

    private static dynamic AppendParagraph(dynamic document, string text)
    {
        dynamic content = document.Content;
        var start = Math.Max(0, (int)content.End - 1);
        content.InsertAfter("\r" + text);
        var end = (int)document.Content.End - 1;
        return document.Range(start + 1, Math.Max(start + 1, end));
    }

    private static void ApplyWordStyle(
        dynamic insertion, ParagraphStyle? style, FontConfig? defaultFont)
    {
        if (defaultFont is not null)
        {
            if (!string.IsNullOrEmpty(defaultFont.Name))
                insertion.Font.Name = defaultFont.Name;
            if (defaultFont.SizePt is not null)
                insertion.Font.Size = defaultFont.SizePt.Value;
            if (defaultFont.Bold is not null)
                insertion.Font.Bold = defaultFont.Bold.Value;
            if (defaultFont.Italic is not null)
                insertion.Font.Italic = defaultFont.Italic.Value;
            if (!string.IsNullOrEmpty(defaultFont.Color))
                insertion.Font.Color = OpenEditCommon.HexToBgr(defaultFont.Color);
        }
        if (style is null)
            return;

        if (!string.IsNullOrEmpty(style.FontName))
            insertion.Font.Name = style.FontName;
        if (style.SizePt is not null)
            insertion.Font.Size = style.SizePt.Value;
        if (style.Bold is not null)
            insertion.Font.Bold = style.Bold.Value;
        if (style.Italic is not null)
            insertion.Font.Italic = style.Italic.Value;
        if (!string.IsNullOrEmpty(style.Color))
            insertion.Font.Color = OpenEditCommon.HexToBgr(style.Color);

        dynamic paragraph = insertion.ParagraphFormat;
        if (!string.IsNullOrEmpty(style.Alignment))
        {
            paragraph.Alignment = style.Alignment switch
            {
                "left" => 0,
                "center" => 1,
                "right" => 2,
                "justify" => 3,
                _ => 0,
            };
        }
        if (style.LineSpacing is not null)
        {
            paragraph.LineSpacingRule = 5; // wdLineSpaceMultiple
            paragraph.LineSpacing = 12 * style.LineSpacing.Value;
        }
        if (style.FirstLineIndentPt is not null)
            paragraph.FirstLineIndent = style.FirstLineIndentPt.Value;
        if (style.SpaceBeforePt is not null)
            paragraph.SpaceBefore = style.SpaceBeforePt.Value;
        if (style.SpaceAfterPt is not null)
            paragraph.SpaceAfter = style.SpaceAfterPt.Value;
    }
}
