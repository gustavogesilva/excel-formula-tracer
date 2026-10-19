using System;
using System.Collections.Generic;
using Tracer.Core.Model;

namespace Tracer.Core.Binding;

/// <summary>A defined name, as read from <c>Workbook.Names</c>.</summary>
public sealed class NameDef
{
    public NameDef(string name, string? scopeSheet, string refersToR1C1)
    {
        Name = name;
        ScopeSheet = scopeSheet;
        RefersToR1C1 = refersToR1C1;
    }

    public string Name { get; }
    /// <summary>Worksheet the name is local to, or null for workbook scope.</summary>
    public string? ScopeSheet { get; }
    /// <summary>The definition, e.g. <c>=Inputs!R2C3</c> or <c>=LAMBDA(x,x*2)</c>.</summary>
    public string RefersToR1C1 { get; }
}

/// <summary>An Excel table (ListObject) layout.</summary>
public sealed class TableDef
{
    public TableDef(string name, string sheet, int headerRow, int firstDataRow, int lastDataRow, int totalsRow, int firstCol, IReadOnlyList<string> columns)
    {
        Name = name;
        Sheet = sheet;
        HeaderRow = headerRow;
        FirstDataRow = firstDataRow;
        LastDataRow = lastDataRow;
        TotalsRow = totalsRow;
        FirstCol = firstCol;
        Columns = columns;
    }

    public string Name { get; }
    public string Sheet { get; }
    /// <summary>0 when the header row is hidden.</summary>
    public int HeaderRow { get; }
    public int FirstDataRow { get; }
    /// <summary>Less than <see cref="FirstDataRow"/> when the table has no data rows.</summary>
    public int LastDataRow { get; }
    /// <summary>0 when there is no totals row.</summary>
    public int TotalsRow { get; }
    public int FirstCol { get; }
    public IReadOnlyList<string> Columns { get; }

    public int LastCol => FirstCol + Columns.Count - 1;
    public int TopRow => HeaderRow > 0 ? HeaderRow : FirstDataRow;
    public int BottomRow => TotalsRow > 0 ? TotalsRow : LastDataRow;

    public int ColumnIndex(string name)
    {
        for (int i = 0; i < Columns.Count; i++)
            if (string.Equals(Columns[i], name, StringComparison.OrdinalIgnoreCase)) return FirstCol + i;
        return -1;
    }

    public bool Contains(int row, int col) => row >= TopRow && row <= BottomRow && col >= FirstCol && col <= LastCol;
}

/// <summary>
/// Snapshot of one open workbook's structure: sheet order, defined names and tables.
/// Built by the add-in from COM (or by tests), then used for binding.
/// </summary>
public sealed class WorkbookModel
{
    private static readonly StringComparer Cmp = StringComparer.OrdinalIgnoreCase;
    private readonly Dictionary<string, int> _sheetIndex = new(Cmp);
    private readonly Dictionary<string, NameDef> _globalNames = new(Cmp);
    private readonly Dictionary<string, NameDef> _localNames = new(Cmp);
    private readonly Dictionary<string, TableDef> _tables = new(Cmp);
    private readonly Dictionary<string, List<TableDef>> _tablesBySheet = new(Cmp);

    public WorkbookModel(string name, IReadOnlyList<string> sheets, IEnumerable<NameDef>? names = null, IEnumerable<TableDef>? tables = null)
    {
        Name = name;
        Sheets = sheets;
        for (int i = 0; i < sheets.Count; i++) _sheetIndex[sheets[i]] = i;
        if (names != null)
            foreach (var n in names)
            {
                if (n.ScopeSheet == null) _globalNames[n.Name] = n;
                else _localNames[n.ScopeSheet + "!" + n.Name] = n;
            }
        if (tables != null)
            foreach (var t in tables)
            {
                _tables[t.Name] = t;
                if (!_tablesBySheet.TryGetValue(t.Sheet, out var list)) _tablesBySheet[t.Sheet] = list = new List<TableDef>();
                list.Add(t);
            }
    }

    public string Name { get; }
    public IReadOnlyList<string> Sheets { get; }
    public int NameCount => _globalNames.Count + _localNames.Count;
    public IEnumerable<NameDef> AllNames => System.Linq.Enumerable.Concat(_globalNames.Values, _localNames.Values);

    public bool HasSheet(string sheet) => _sheetIndex.ContainsKey(sheet);

    /// <summary>Canonical sheet name (Excel's casing), or null.</summary>
    public string? Sheet(string sheet) => _sheetIndex.TryGetValue(sheet, out int i) ? Sheets[i] : null;

    /// <summary>Sheets from <paramref name="first"/> to <paramref name="last"/> in tab order (a 3D reference).</summary>
    public IEnumerable<string> SheetSpan(string first, string last)
    {
        if (!_sheetIndex.TryGetValue(first, out int a) || !_sheetIndex.TryGetValue(last, out int b)) yield break;
        if (a > b) (a, b) = (b, a);
        for (int i = a; i <= b; i++) yield return Sheets[i];
    }

    /// <summary>Sheet-local name on <paramref name="sheet"/> first (when given), then workbook scope.</summary>
    public NameDef? FindName(string name, string? sheet)
    {
        if (sheet != null && _localNames.TryGetValue(sheet + "!" + name, out var local)) return local;
        return _globalNames.TryGetValue(name, out var global) ? global : null;
    }

    public NameDef? FindLocalName(string name, string sheet) =>
        _localNames.TryGetValue(sheet + "!" + name, out var local) ? local : null;

    public TableDef? FindTable(string name) => _tables.TryGetValue(name, out var t) ? t : null;

    public TableDef? FindTableAt(string sheet, int row, int col)
    {
        if (!_tablesBySheet.TryGetValue(sheet, out var list)) return null;
        foreach (var t in list)
            if (t.Contains(row, col)) return t;
        return null;
    }

    /// <summary>Cheap fingerprint of the structure; a change means bindings (names/tables/sheet order) may be stale.</summary>
    public string Signature()
    {
        var sb = new System.Text.StringBuilder();
        foreach (var s in Sheets) sb.Append(s).Append('|');
        sb.Append('#');
        foreach (var n in AllNames) sb.Append(n.ScopeSheet).Append('!').Append(n.Name).Append('=').Append(n.RefersToR1C1).Append('|');
        sb.Append('#');
        foreach (var t in _tables.Values)
            sb.Append(t.Name).Append('@').Append(t.Sheet).Append(t.HeaderRow).Append(',').Append(t.FirstDataRow).Append(',')
              .Append(t.LastDataRow).Append(',').Append(t.TotalsRow).Append(',').Append(t.FirstCol).Append(',')
              .Append(string.Join(";", t.Columns)).Append('|');
        return sb.ToString();
    }
}

/// <summary>The set of open workbooks.</summary>
public sealed class WorkspaceModel
{
    private readonly Dictionary<string, WorkbookModel> _books = new(StringComparer.OrdinalIgnoreCase);

    public IEnumerable<WorkbookModel> Books => _books.Values;
    public WorkbookModel? Book(string name) => _books.TryGetValue(name, out var b) ? b : null;
    public void Set(WorkbookModel book) => _books[book.Name] = book;
    public bool Remove(string name) => _books.Remove(name);
}
