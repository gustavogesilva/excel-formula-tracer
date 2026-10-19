using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Tracer.AddIn.Excel;
using Tracer.Core.Model;
using Tracer.Core.Text;
using Tracer.Core.Tracing;

namespace Tracer.AddIn.UI;

internal enum TraceMode
{
    Precedents,
    Dependents,
}

/// <summary>One screen of the trace window: a trace of a cell/range, or a listing (cells of a range, parts of a name).</summary>
internal sealed class Level
{
    public TraceMode Mode;
    public string Book = string.Empty;
    public string Sheet = string.Empty;
    public CellRect Focus;
    public string Title = string.Empty;
    public string? FocusValue;
    /// <summary>A1 formula of the focus cell; precedent item spans point into it.</summary>
    public string? Formula;
    public List<TraceItem> Items = new();
    public string Note = string.Empty;
    public int Selected;
    public double Millis;
    /// <summary>Display values already read in bulk (range listings), keyed by item.</summary>
    public Dictionary<TraceItem, string>? Values;
}

/// <summary>Builds levels from the engine and Excel; owns the window and the jump history.</summary>
internal sealed class TraceController
{
    private const int MaxListedCells = 10_000;
    private const int MaxJumpHistory = 100;

    private readonly ExcelHost _host = ExcelHost.Instance;
    private readonly LinkedList<SavedPosition> _jumps = new();
    private TraceWindow? _window;

    public static TraceController Instance { get; } = new();

    public TraceWindow Window => _window ??= new TraceWindow(this);

    /// <summary>Creates the window handle up front so the first hotkey shows it without construction cost.</summary>
    public void Prewarm() => _ = Window.Handle;

    // ---- entry points (hotkeys / ribbon) -------------------------------------------------------

    public void TraceActive(TraceMode mode)
    {
        var sw = Stopwatch.StartNew();
        // Pressed again while the window is open (after clicking into the sheet): re-trace, keep the original return point.
        var origin = Window.Visible ? null : SavedPosition.Capture();
        Level? level;
        if (mode == TraceMode.Precedents)
        {
            var ac = ExcelHost.ActiveCell();
            if (ac == null) return;
            var (wb, ws, cell) = ac.Value;
            level = Precedents(wb.Name, ws.Name, ExcelHost.RectOf(cell));
        }
        else
        {
            var sel = ExcelHost.ActiveSelection();
            if (sel == null) return;
            var (wb, ws, rect) = sel.Value;
            level = Dependents(wb.Name, ws.Name, rect);
        }
        if (level == null) return;
        level.Millis = sw.Elapsed.TotalMilliseconds;
        Window.Open(level, origin);
    }

    public void JumpBack()
    {
        while (_jumps.Count > 0)
        {
            var p = _jumps.Last!.Value;
            _jumps.RemoveLast();
            if (p.Restore()) return;
        }
    }

    public void PushJump(SavedPosition? p)
    {
        if (p == null) return;
        _jumps.AddLast(p);
        while (_jumps.Count > MaxJumpHistory) _jumps.RemoveFirst();
    }

    public void RebuildIndex()
    {
        var sw = Stopwatch.StartNew();
        _host.Engine.MarkAllDirty();
        foreach (var b in _host.Engine.Workspace.Books) _host.InvalidateStructure(b.Name);
        int sheets = _host.EnsureIndexed();
        var (cells, blocks, rules, templates) = _host.Engine.Stats();
        ExcelHost.App.StatusBar = $"Formula Tracer: indexed {sheets} sheets, {cells:N0} formula cells → {blocks:N0} blocks, {rules:N0} rules, {templates:N0} distinct formulas in {sw.ElapsedMilliseconds:N0} ms";
        ExcelAsyncUtilClearStatusLater();
    }

    private static void ExcelAsyncUtilClearStatusLater()
    {
        var t = new System.Windows.Forms.Timer { Interval = 6000 };
        t.Tick += (_, _) =>
        {
            t.Stop();
            t.Dispose();
            try { ExcelHost.App.StatusBar = false; } catch (COMException) { }
        };
        t.Start();
    }

    // ---- level builders ---------------------------------------------------------------

    public Level? Precedents(string book, string sheet, CellRect cell)
    {
        var range = ExcelHost.GetRange(book, sheet, CellRect.Cell(cell.Row1, cell.Col1));
        if (range == null) return null;
        _host.EnsureModel(ExcelHost.App.Workbooks[book]);

        var focus = CellRect.Cell(cell.Row1, cell.Col1);
        var (value, _) = _host.ReadCell(book, sheet, focus);
        var level = new Level
        {
            Mode = TraceMode.Precedents,
            Book = book,
            Sheet = sheet,
            Focus = focus,
            FocusValue = value,
            Title = "Precedents of " + Address(book, sheet, focus),
        };

        string? formula = _host.ReadFormula(range);
        if (formula == null)
        {
            level.Note = "No formula: this cell holds a constant. Tab shows its dependents.";
            return level;
        }

        var result = _host.Engine.Precedents(book, sheet, focus.Row1, focus.Col1, formula);
        level.Formula = result.Formula.Text;
        level.Items = result.Items;
        int dyn = result.Items.FindAll(i => i.Kind == TraceItemKind.Dynamic).Count;
        if (dyn > 0) level.Note = "→ on an OFFSET/INDIRECT row evaluates it to its current target.";
        return level;
    }

    public Level Dependents(string book, string sheet, CellRect rect)
    {
        int indexed = _host.EnsureIndexed(msg => Window.SetStatus(msg));
        var result = _host.Engine.Dependents(book, sheet, rect);
        var (value, formula) = rect.IsSingleCell ? _host.ReadCell(book, sheet, rect) : ("", null);
        var level = new Level
        {
            Mode = TraceMode.Dependents,
            Book = book,
            Sheet = sheet,
            Focus = rect,
            FocusValue = value,
            Formula = formula,
            Items = result.Items,
            Title = "Dependents of " + Address(book, sheet, rect),
        };
        var notes = new List<string>();
        if (indexed > 0) notes.Add($"indexed {indexed} sheet(s)");
        if (result.DynamicFormulaCells > 0) notes.Add($"{result.DynamicFormulaCells:N0} formula cells use OFFSET/INDIRECT and may depend on this without being listed");
        level.Note = string.Join(" · ", notes);
        return level;
    }

    /// <summary>What → / double-click opens for an item, or null with a reason.</summary>
    public Level? Drill(Level current, TraceItem item, TraceMode mode, out string? reason)
    {
        reason = null;
        if (item.Kind == TraceItemKind.Dynamic && !item.HasTarget && !ResolveDynamic(current, item))
        {
            reason = "Could not evaluate " + item.Label;
            return null;
        }
        if (item.IsSpill && item.Rect.IsSingleCell && ExcelHost.SpillRange(item.Book, item.Sheet, item.Rect) is CellRect spill)
        {
            item.Rect = spill;
            item.Kind = TraceItemKind.Range;
        }

        switch (item.Kind)
        {
            case TraceItemKind.Name when item.Children != null && item.Children.Count != 1:
                return Listing(current, item.Label, item.Children);
            case TraceItemKind.External:
                reason = "Workbook is not open: " + item.Label;
                return null;
            case TraceItemKind.Broken:
            case TraceItemKind.Unresolved:
                reason = "Nothing to trace: " + item.Label;
                return null;
        }
        if (!item.HasTarget)
        {
            reason = "No target for " + item.Label;
            return null;
        }

        if (mode == TraceMode.Dependents)
        {
            // A dependents block of several cells lists them first; a single cell traces further out.
            if (current.Mode == TraceMode.Dependents && !item.Rect.IsSingleCell && current.Items.Contains(item))
                return CellsOf(current, item, TraceMode.Dependents);
            return Dependents(item.Book, item.Sheet, item.Rect);
        }
        return item.Rect.IsSingleCell ? Precedents(item.Book, item.Sheet, item.Rect) : CellsOf(current, item, TraceMode.Precedents);
    }

    /// <summary>Tab: the other direction for the same focus.</summary>
    public Level? Toggle(Level current) =>
        current.Mode == TraceMode.Precedents
            ? Dependents(current.Book, current.Sheet, current.Focus)
            : Precedents(current.Book, current.Sheet, current.Focus);

    private bool ResolveDynamic(Level current, TraceItem item)
    {
        if (item.Expression == null) return false;
        var target = ExcelHost.EvaluateReference(current.Book, current.Sheet, item.Expression);
        if (target == null) return false;
        var (book, sheet, rect) = target.Value;
        item.Book = book;
        item.Sheet = sheet;
        item.Rect = rect;
        item.HasTarget = true;
        item.Kind = rect.IsSingleCell ? TraceItemKind.Cell : TraceItemKind.Range;
        item.Label = A1.Qualified(book, sheet, rect, current.Book, current.Sheet) + "  ⇐ " + item.Expression;
        return true;
    }

    private Level Listing(Level current, string label, List<TraceItem> items) => new()
    {
        Mode = current.Mode,
        Book = current.Book,
        Sheet = current.Sheet,
        Focus = current.Focus,
        FocusValue = current.FocusValue,
        Formula = current.Formula,
        Title = label,
        Items = items,
    };

    /// <summary>Non-empty cells of a range (or of a dependents block), with values from one bulk read.</summary>
    private Level CellsOf(Level current, TraceItem item, TraceMode mode)
    {
        var cells = _host.ReadRangeCells(item.Book, item.Sheet, item.Rect, MaxListedCells, out bool truncated);
        var level = new Level
        {
            Mode = mode,
            Book = item.Book,
            Sheet = item.Sheet,
            Focus = item.Rect,
            Title = "Cells of " + Address(item.Book, item.Sheet, item.Rect),
            Values = new Dictionary<TraceItem, string>(),
        };
        foreach (var (cell, value, f) in cells)
        {
            var ti = new TraceItem
            {
                Kind = TraceItemKind.Cell,
                Book = item.Book,
                Sheet = item.Sheet,
                Rect = cell,
                HasTarget = true,
                Label = A1.Qualified(item.Book, item.Sheet, cell, current.Book, current.Sheet),
                Expression = f == null ? null : FormulaRenderer.ToA1(_host.Engine.Templates.Get(f), cell.Row1, cell.Col1).Text,
            };
            level.Items.Add(ti);
            level.Values[ti] = value;
        }
        level.Note = truncated ? $"showing the first {MaxListedCells:N0} non-empty cells" : $"{cells.Count:N0} non-empty cells";
        return level;
    }

    /// <summary>Value and formula for a list row (read on demand while the list paints).</summary>
    public (string Value, string Formula) Display(Level level, TraceItem item)
    {
        if (level.Values != null && level.Values.TryGetValue(item, out var cached)) return (cached, item.Expression ?? "");
        if (!item.HasTarget) return ("", item.Expression ?? "");
        if (item.Rect.IsSingleCell)
        {
            var (v, f) = _host.ReadCell(item.Book, item.Sheet, item.Rect);
            return (v, item.Expression ?? f ?? "");
        }
        return ($"[{item.Rect.Rows:N0}×{item.Rect.Cols:N0}]", item.Expression ?? "");
    }

    private static string Address(string book, string sheet, CellRect r)
    {
        string active = "";
        try { active = ExcelHost.App.ActiveWorkbook?.Name ?? ""; } catch (COMException) { }
        return A1.Qualified(book, sheet, r, active, null);
    }
}
