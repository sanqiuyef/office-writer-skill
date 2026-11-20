using System.Runtime.InteropServices;
using System.Text.Json;

namespace OfficeNative;

internal static class GenericOpenCommands
{
    public static Dictionary<string, object?> RunGetOpen(JsonElement p)
    {
        var result = new Dictionary<string, object?>
        {
            ["word"] = new List<Dictionary<string, object?>>(),
            ["excel"] = new List<Dictionary<string, object?>>(),
            ["powerpoint"] = new List<Dictionary<string, object?>>(),
        };
        CollectWord(result, "Word.Application");
        CollectWord(result, "kwps.Application");
        CollectExcel(result, "Excel.Application");
        CollectExcel(result, "KET.Application");
        CollectPowerPoint(result, "PowerPoint.Application");
        CollectPowerPoint(result, "KWPP.Application");
        return result;
    }

    private static void CollectWord(Dictionary<string, object?> result, string progId)
    {
        dynamic? app = GetActive(progId);
        if (app is null)
            return;
        try
        {
            var items = (List<Dictionary<string, object?>>)result["word"]!;
            dynamic documents = app.Documents;
            for (var index = 1; index <= (int)documents.Count; index++)
            {
                dynamic document = documents.Item(index);
                AddUnique(items, (string)document.FullName, progId);
            }
        }
        catch { /* application may be closing */ }
        finally { NativeCommon.ReleaseCom(app); }
    }

    private static void CollectExcel(Dictionary<string, object?> result, string progId)
    {
        dynamic? app = GetActive(progId);
        if (app is null)
            return;
        try
        {
            var items = (List<Dictionary<string, object?>>)result["excel"]!;
            dynamic workbooks = app.Workbooks;
            for (var index = 1; index <= (int)workbooks.Count; index++)
            {
                dynamic workbook = workbooks.Item(index);
                AddUnique(items, (string)workbook.FullName, progId);
            }
        }
        catch { }
        finally { NativeCommon.ReleaseCom(app); }
    }

    private static void CollectPowerPoint(Dictionary<string, object?> result, string progId)
    {
        dynamic? app = GetActive(progId);
        if (app is null)
            return;
        try
        {
            var items = (List<Dictionary<string, object?>>)result["powerpoint"]!;
            dynamic presentations = app.Presentations;
            for (var index = 1; index <= (int)presentations.Count; index++)
            {
                dynamic presentation = presentations.Item(index);
                AddUnique(items, (string)presentation.FullName, progId);
            }
        }
        catch { }
        finally { NativeCommon.ReleaseCom(app); }
    }

    private static void AddUnique(
        List<Dictionary<string, object?>> items, string filePath, string progId)
    {
        if (items.Any(item => string.Equals(
                item["file_path"] as string, filePath, StringComparison.OrdinalIgnoreCase)))
            return;
        items.Add(new Dictionary<string, object?>
        {
            ["file_path"] = filePath,
            ["application"] = progId,
        });
    }

    internal static dynamic? GetActive(string progId)
    {
        try
        {
            if (CLSIDFromProgID(progId, out var clsid) != 0)
                return null;
            GetActiveObject(ref clsid, IntPtr.Zero, out var instance);
            return instance;
        }
        catch (COMException)
        {
            return null;
        }
        catch (InvalidCastException)
        {
            return null;
        }
    }

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string progId, out Guid clsid);

    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(
        ref Guid rclsid, IntPtr reserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object instance);
}
