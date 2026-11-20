using System.Text.Json;

namespace OfficeNative;

/// <summary>PDF command facade registered by CommandRouter.</summary>
internal static class PdfFileCommands
{
    public static Dictionary<string, object?> RunCreatePdfDocument(JsonElement p) =>
        PdfFileWriter.RunCreatePdfDocument(p);

    public static Dictionary<string, object?> RunReadPdfDocument(JsonElement p) =>
        PdfFileReader.RunReadPdfDocument(p);

    public static Dictionary<string, object?> RunReadPdfStructured(JsonElement p) =>
        PdfFileReader.RunReadPdfStructured(p);

    public static Dictionary<string, object?> RunRenderDocumentPages(JsonElement p) =>
        PdfFileRenderCommands.RunRenderDocumentPages(p);
}
