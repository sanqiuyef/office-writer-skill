using System.Text.Json;

namespace OfficeNative;

/// <summary>所有 C# 命令的唯一注册点；模块只提供静态 Run 方法，不改 Program.cs。</summary>
internal static class CommandRouter
{
    public static (string Tool, Dictionary<string, object?> Data) Dispatch(
        string command, JsonElement p)
    {
        return command switch
        {
            "version" => ("version", VersionInfo()),
            "word-revisions" or "word-manage-revisions"
                => ("word_manage_revisions", WordCommands.RunRevisions(p)),
            "word-footnotes" or "word-manage-footnotes"
                => ("word_manage_footnotes", WordCommands.RunFootnotes(p)),
            "word-content-controls" or "word-manage-content-controls"
                => ("word_manage_content_controls", WordCommands.RunContentControls(p)),
            "ppt-effects" or "ppt-manage-effects"
                => ("ppt_manage_effects", PptCommands.RunEffects(p)),
            "ppt-smartart" or "ppt-manage-smartart"
                => ("ppt_manage_smartart", PptCommands.RunSmartArt(p)),
            "ppt-shapes" or "ppt-manage-shapes"
                => ("ppt_manage_shapes", PptShapeCommands.RunShapes(p)),
            "excel-pivots" or "excel-manage-pivots"
                => ("excel_manage_pivots", ExcelCommands.RunPivots(p)),
            "excel-print" => ("excel_print", ExcelCommands.RunPrint(p)),
            "create-word-document" => ("create_word_document", WordFileCommands.RunCreateWordDocument(p)),
            "edit-word-document" => ("edit_word_document", WordFileCommands.RunEditWordDocument(p)),
            "create-word-from-markdown" => ("create_word_from_markdown", WordFileMarkdown.RunCreateWordFromMarkdown(p)),
            "read-word-document" => ("read_word_document", WordFileReader.RunReadWordDocument(p)),
            "word-set-properties" => ("word_set_properties", WordFileCommands.RunWordSetProperties(p)),
            "word-manage-comments" => ("word_manage_comments", WordFileCommands.RunWordManageComments(p)),
            "validate-word-document" => ("validate_word_document", WordValidateCommand.RunValidate(p)),
            "create-excel-workbook" => ("create_excel_workbook", ExcelFileCommands.RunCreateWorkbook(p)),
            "edit-excel-workbook" => ("edit_excel_workbook", ExcelFileCommands.RunEditWorkbook(p)),
            "read-excel-workbook" => ("read_excel_workbook", ExcelFileCommands.RunReadWorkbook(p)),
            "excel-add-chart" => ("excel_add_chart", ExcelFileCommands.RunAddChart(p)),
            "recalculate-excel-formulas"
                => ("recalculate_excel_formulas", RecalcCommands.RunRecalculate(p)),
            "create-powerpoint" => ("create_powerpoint", PptFileCommands.RunCreate(p)),
            "edit-powerpoint" => ("edit_powerpoint", PptFileCommands.RunEdit(p)),
            "read-powerpoint" => ("read_powerpoint", PptFileCommands.RunRead(p)),
            "ppt-list-color-schemes" => ("ppt_list_color_schemes", PptFileCommands.RunListColorSchemes(p)),
            "ppt-apply-color-scheme" => ("ppt_apply_color_scheme", PptFileCommands.RunApplyColorScheme(p)),
            "ppt-optimize-slide-layout" => ("ppt_optimize_slide_layout", PptFileCommands.RunOptimizeSlideLayout(p)),
            "ppt-clean-unused-placeholders" => ("ppt_clean_unused_placeholders", PptFileCommands.RunCleanUnusedPlaceholders(p)),
            "ppt-analyze-template" => ("ppt_analyze_template", PptFileCommands.RunAnalyzeTemplate(p)),
            "ppt-set-hyperlinks" => ("ppt_set_hyperlinks", PptFileCommands.RunSetHyperlinks(p)),
            "ppt-create-framework-diagram" => ("ppt_create_framework_diagram", PptFileFramework.CreateDiagram(p)),
            "merge-office-template" => ("merge_office_template", MergeTemplateCommands.RunMerge(p)),
            "batch-edit-office-document" => ("batch_edit_office_document", BatchEditCommands.RunBatchEdit(p)),
            "create-pdf-document" => ("create_pdf_document", PdfFileCommands.RunCreatePdfDocument(p)),
            "read-pdf-document" => ("read_pdf_document", PdfFileCommands.RunReadPdfDocument(p)),
            "read-pdf-structured" => ("read_pdf_structured", PdfFileCommands.RunReadPdfStructured(p)),
            "render-document-pages" => ("render_document_pages", PdfFileCommands.RunRenderDocumentPages(p)),
            "read-txt-document" => ("read_txt_document", GenericTxtCommands.RunReadTxt(p)),
            "convert-to-pdf" => ("convert_to_pdf", ConvertCommands.RunConvert(p)),
            "inspect-office-document" => ("inspect_office_document", InspectCommands.RunInspect(p)),
            "get-open-office-documents" => ("get_open_office_documents", GenericOpenCommands.RunGetOpen(p)),
            "replace-open-document-content"
                => ("replace_open_document_content", OpenReplaceCommands.RunReplace(p)),
            "edit-open-word-document"
                => ("edit_open_word_document", OpenEditWordCommands.RunEditOpenWordDocument(p)),
            "edit-open-excel-workbook"
                => ("edit_open_excel_workbook", OpenEditExcelCommands.RunEditOpenExcelWorkbook(p)),
            "edit-open-powerpoint"
                => ("edit_open_powerpoint", OpenEditPptCommands.RunEditOpenPowerpoint(p)),
            "ppt-add-flowchart-slide"
                => ("ppt_add_flowchart_slide", OpenEditPptCommands.RunAddFlowchartSlide(p)),
            _ => throw new NativeUsageException($"未知命令: {command}"),
        };
    }

    private static Dictionary<string, object?> VersionInfo() => new()
    {
        ["runtime"] = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
        ["openxml"] = typeof(DocumentFormat.OpenXml.Packaging.WordprocessingDocument)
            .Assembly.GetName().Version?.ToString(),
        ["word_progid"] = Type.GetTypeFromProgID("Word.Application") is not null,
        ["excel_progid"] = Type.GetTypeFromProgID("Excel.Application") is not null,
        ["ppt_progid"] = Type.GetTypeFromProgID("PowerPoint.Application") is not null,
    };
}
