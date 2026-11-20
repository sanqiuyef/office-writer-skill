using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeNative;

/// <summary>
/// Word 表格布局网格：与 python-docx 的 row.cells 语义对齐。
/// gridSpan 按网格列展开；vMerge=continue 回溯到顶层 restart 单元格。
/// </summary>
internal static class WordTableGrid
{
    public static int GridSpan(W.TableCell cell) =>
        Math.Max(1, cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1);

    /// <summary>按网格列展开一行，合并单元格重复出现。</summary>
    public static List<W.TableCell> ExpandRow(
        IReadOnlyList<W.TableRow> rows, int rowIndex)
    {
        var cells = new List<W.TableCell>();
        foreach (var cell in rows[rowIndex].Elements<W.TableCell>())
        {
            var root = RootCell(rows, rowIndex, cell);
            for (var index = 0; index < GridSpan(root); index++)
                cells.Add(root);
        }
        return cells;
    }

    /// <summary>vMerge=continue 时沿同一网格列向上回溯到 restart 单元格。</summary>
    public static W.TableCell RootCell(
        IReadOnlyList<W.TableRow> rows, int rowIndex, W.TableCell cell)
    {
        while (rowIndex > 0
               && cell.TableCellProperties?.VerticalMerge is { } merge
               && merge.Val?.Value != W.MergedCellValues.Restart)
        {
            var column = GridOffset(rows[rowIndex], cell);
            rowIndex--;
            var above = CellAtGridOffset(rows[rowIndex], column);
            if (above is null)
                break;
            cell = above;
        }
        return cell;
    }

    private static int GridOffset(W.TableRow row, W.TableCell cell)
    {
        var offset = row.TableRowProperties?.GetFirstChild<W.GridBefore>()?.Val?.Value
                     ?? 0;
        foreach (var item in row.Elements<W.TableCell>())
        {
            if (ReferenceEquals(item, cell))
                break;
            offset += GridSpan(item);
        }
        return offset;
    }

    private static W.TableCell? CellAtGridOffset(W.TableRow row, int column)
    {
        var offset = row.TableRowProperties?.GetFirstChild<W.GridBefore>()?.Val?.Value
                     ?? 0;
        foreach (var cell in row.Elements<W.TableCell>())
        {
            var span = GridSpan(cell);
            if (column >= offset && column < offset + span)
                return cell;
            offset += span;
        }
        return null;
    }
}
