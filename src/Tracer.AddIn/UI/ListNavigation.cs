using System;

namespace Tracer.AddIn.UI;

/// <summary>
/// Tree-style keyboard rules for the trace list. Row 0 is the traced cell or range itself (the root);
/// rows 1..n are its precedents or dependents. Opening a trace selects the root, where Excel already is,
/// so ↓ or → always moves Excel to the first argument, even when there is only one.
/// Pure logic (no WinForms/Excel), unit-tested in Tracer.Core.Tests.
/// </summary>
internal static class ListNavigation
{
    public enum Move
    {
        None,
        /// <summary>Select another row (Excel follows the selection).</summary>
        SelectRow,
        /// <summary>Open the selected item: its precedents/dependents, the cells of a range, the parts of a name.</summary>
        Drill,
        /// <summary>Return to the previous trace.</summary>
        Back,
    }

    public static int RowCount(int items) => items + 1;

    /// <summary>Index into the item list for a row, or -1 for the root row (or a row out of range).</summary>
    public static int ItemIndex(int row, int items) => row >= 1 && row <= items ? row - 1 : -1;

    /// <summary>→: from the root, step down to the first item; from an item, drill into it.</summary>
    public static (Move Move, int Row) Right(int row, int items) =>
        row <= 0 ? (items > 0 ? (Move.SelectRow, 1) : (Move.None, 0)) : (Move.Drill, row);

    /// <summary>←: from an item, go up to the root; from the root, go back to the previous trace.</summary>
    public static (Move Move, int Row) Left(int row) => row > 0 ? (Move.SelectRow, 0) : (Move.Back, 0);

    /// <summary>A remembered selection, kept valid when the list is shorter than it was.</summary>
    public static int Clamp(int row, int items) => Math.Max(0, Math.Min(row, items));
}
