using System.Collections.Generic;

namespace Tracer.AddIn.UI;

/// <summary>One formula edit made from the trace window: where, and what the cell held before.</summary>
internal readonly struct FormulaEdit
{
    public FormulaEdit(string book, string sheet, int row, int col, string previous)
    {
        Book = book;
        Sheet = sheet;
        Row = row;
        Col = col;
        Previous = previous;
    }

    public string Book { get; }
    public string Sheet { get; }
    public int Row { get; }
    public int Col { get; }
    /// <summary>The cell's A1 formula (or constant text) before the edit; empty for an empty cell.</summary>
    public string Previous { get; }
}

/// <summary>
/// Undo history for edits made in the trace window. Excel clears its own undo stack whenever an add-in writes a
/// cell, so the window keeps this one. Bounded: the oldest edits drop off. Pure, unit-tested.
/// </summary>
internal sealed class EditHistory
{
    private readonly LinkedList<FormulaEdit> _edits = new();
    private readonly int _capacity;

    public EditHistory(int capacity = 50) => _capacity = capacity;

    public int Count => _edits.Count;

    public void Push(FormulaEdit edit)
    {
        _edits.AddLast(edit);
        while (_edits.Count > _capacity) _edits.RemoveFirst();
    }

    public bool TryPop(out FormulaEdit edit)
    {
        if (_edits.Count == 0)
        {
            edit = default;
            return false;
        }
        edit = _edits.Last!.Value;
        _edits.RemoveLast();
        return true;
    }
}
