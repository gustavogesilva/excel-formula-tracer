using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Tracer.AddIn.UI;
using Tracer.Core.Binding;
using Tracer.Core.Model;
using Tracer.Core.Text;
using Tracer.Core.Tracing;
using XlRange = Microsoft.Office.Interop.Excel.Range;

namespace Tracer.AddIn.Excel;

/// <summary>
/// All Excel object-model access. Every member must run on Excel's main thread
/// (macro commands, ribbon callbacks, WinForms events of windows created on that thread).
/// Reads are bulk (<c>object[,]</c> per call) because per-call COM overhead dominates.
/// </summary>
internal sealed class ExcelHost : ICellReader
{
    private static readonly TimeSpan SnapshotMaxAge = TimeSpan.FromSeconds(30);

    private sealed class Snapshot
    {
        public WorkbookModel Model = null!;
        public string SheetsKey = string.Empty;
        public int NameCount;
        public DateTime TakenUtc;
    }

    private readonly Dictionary<string, Snapshot> _snapshots = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _structureDirty = new(StringComparer.OrdinalIgnoreCase);
    private bool? _hasFormula2;
    /// <summary>Every bulk read goes through this, so none asks Excel for too much memory at once.</summary>
    private readonly ChunkedReads _reads = new();

    public static ExcelHost Instance { get; } = new();

    public TraceEngine Engine { get; } = new();

    public static Application App => (Application)ExcelDnaUtil.Application;

    /// <summary>
    /// Excel is in the middle of a recalculation. Data tables recalculate asynchronously, possibly for minutes;
    /// background work (idle indexing, the highlight) waits for it rather than interrupting it. Manual mode's
    /// "pending" doesn't count: nothing is running.
    /// </summary>
    public static bool IsCalculating() => App.CalculationState == XlCalculationState.xlCalculating;

    // ---- workbook structure ---------------------------------------------------------------

    /// <summary>Forces the next <see cref="EnsureModel"/> for this workbook to re-read names, tables and sheets.</summary>
    public void InvalidateStructure(string book) => _structureDirty.Add(book);

    /// <summary>Keeps the engine's workbook model current, re-reading names/tables only when something looks changed.</summary>
    public WorkbookModel EnsureModel(Workbook wb)
    {
        string name = wb.Name;
        var sheets = new List<string>();
        foreach (Worksheet ws in wb.Worksheets) sheets.Add(ws.Name);
        string sheetsKey = string.Join("\u0001", sheets);
        int nameCount = wb.Names.Count;

        if (_snapshots.TryGetValue(name, out var snap) && !_structureDirty.Contains(name)
            && snap.SheetsKey == sheetsKey && snap.NameCount == nameCount && DateTime.UtcNow - snap.TakenUtc < SnapshotMaxAge)
            return snap.Model;

        var model = new WorkbookModel(name, sheets, ReadNames(wb), ReadTables(wb));
        _snapshots[name] = new Snapshot { Model = model, SheetsKey = sheetsKey, NameCount = nameCount, TakenUtc = DateTime.UtcNow };
        _structureDirty.Remove(name);
        Engine.UpdateBook(model);
        return model;
    }

    /// <summary>Syncs the engine with the set of open workbooks (closed/renamed books are dropped).</summary>
    public List<Workbook> ReconcileBooks()
    {
        var open = new List<Workbook>();
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (Workbook wb in App.Workbooks)
        {
            open.Add(wb);
            names.Add(wb.Name);
        }
        var gone = new List<string>();
        foreach (var b in Engine.Workspace.Books)
            if (!names.Contains(b.Name)) gone.Add(b.Name);
        foreach (var b in gone)
        {
            Engine.RemoveBook(b);
            _snapshots.Remove(b);
        }
        foreach (var wb in open) EnsureModel(wb);
        return open;
    }

    private static List<NameDef> ReadNames(Workbook wb)
    {
        var list = new List<NameDef>();
        foreach (Name n in wb.Names)
        {
            try
            {
                string full = n.Name;
                int bang = full.LastIndexOf('!');
                string bare = bang >= 0 ? full.Substring(bang + 1) : full;
                string? scope = (n.Parent as Worksheet)?.Name;
                if (n.RefersToR1C1 is string refers) list.Add(new NameDef(bare, scope, refers));
            }
            catch (COMException)
            {
                // Some names (e.g. broken external links) throw on access; skip them.
            }
        }
        return list;
    }

    private static List<TableDef> ReadTables(Workbook wb)
    {
        var list = new List<TableDef>();
        foreach (Worksheet ws in wb.Worksheets)
        {
            var los = ws.ListObjects;
            if (los.Count == 0) continue;
            foreach (ListObject lo in los)
            {
                try
                {
                    XlRange all = lo.Range;
                    int top = all.Row, firstCol = all.Column;
                    int header = lo.ShowHeaders ? top : 0;
                    XlRange? body = lo.DataBodyRange;
                    int firstData = body?.Row ?? (header > 0 ? header + 1 : top);
                    int lastData = body != null ? body.Row + body.Rows.Count - 1 : firstData - 1;
                    int totals = lo.ShowTotals ? lo.TotalsRowRange.Row : 0;
                    var cols = new List<string>();
                    foreach (ListColumn lc in lo.ListColumns) cols.Add(lc.Name);
                    list.Add(new TableDef(lo.Name, ws.Name, header, firstData, lastData, totals, firstCol, cols));
                }
                catch (COMException)
                {
                }
            }
        }
        return list;
    }

    // ---- formulas -----------------------------------------------------------------------

    /// <summary>
    /// <c>Formula2R1C1</c> (dynamic-array aware, Excel 365) with a fallback to <c>FormulaR1C1</c> on older builds.
    /// Always en-US syntax regardless of the user's locale. One COM call for any range size.
    /// </summary>
    public object? ReadFormulasR1C1(XlRange r)
    {
        if (_hasFormula2 != false)
        {
            try
            {
                var v = r.GetType().InvokeMember("Formula2R1C1", BindingFlags.GetProperty, null, r, null);
                _hasFormula2 = true;
                return v;
            }
            catch (Exception ex) when (_hasFormula2 == null && IsUnknownMember(ex))
            {
                _hasFormula2 = false;
            }
        }
        return r.FormulaR1C1;
    }

    /// <summary>The member doesn't exist in this Excel (e.g. Formula2 before dynamic arrays). Anything else, such as
    /// out-of-memory or Excel being busy, is a real failure of the call.</summary>
    private static bool IsUnknownMember(Exception ex)
    {
        const int DispUnknownName = unchecked((int)0x80020006);
        var com = (ex as TargetInvocationException)?.InnerException as COMException ?? ex as COMException;
        return ex is MissingMethodException || com?.HResult == DispUnknownName;
    }

    /// <summary>A1 formula for display, <c>Formula2</c> when available so no implicit-intersection <c>@</c> is added.</summary>
    public string? ReadFormulaA1(XlRange cell)
    {
        if (_hasFormula2 == true)
        {
            try
            {
                return cell.GetType().InvokeMember("Formula2", BindingFlags.GetProperty, null, cell, null) as string;
            }
            catch (Exception ex) when (ex is TargetInvocationException || ex is COMException)
            {
            }
        }
        return cell.Formula as string;
    }

    /// <summary>What the cell holds as A1 text (its formula, or its constant), for editing in the trace window.</summary>
    public string ReadEditText(string book, string sheet, CellRect cell)
    {
        var r = GetRange(book, sheet, cell) ?? throw new InvalidOperationException("The cell is no longer available.");
        return ReadFormulaA1(r) ?? string.Empty;
    }

    /// <summary>
    /// Writes a formula (en-US A1 syntax, as the window shows it) or a constant, exactly as if typed into the cell:
    /// <c>Formula2</c> when available (no implicit-intersection rewriting), else <c>Formula</c>.
    /// Throws <see cref="COMException"/> when Excel rejects the text.
    /// </summary>
    public void WriteFormula(string book, string sheet, CellRect cell, string text)
    {
        var r = GetRange(book, sheet, cell) ?? throw new InvalidOperationException("The cell is no longer available.");
        if (_hasFormula2 != false)
        {
            try
            {
                r.GetType().InvokeMember("Formula2", BindingFlags.SetProperty, null, r, new object[] { text });
                _hasFormula2 = true;
                return;
            }
            catch (Exception ex) when (ex is TargetInvocationException || ex is COMException || ex is MissingMethodException)
            {
                if (IsUnknownMember(ex)) _hasFormula2 = false; // Excel without Formula2
                else throw (ex as TargetInvocationException)?.InnerException as COMException ?? ex;
            }
        }
        r.Formula = text;
    }

    /// <summary>The cell's R1C1 formula, or null when it holds a constant or nothing.</summary>
    public string? ReadFormula(XlRange cell) =>
        ReadFormulasR1C1(cell) is string s && s.Length > 1 && s[0] == '=' ? s : null;

    /// <summary>Re-indexes one sheet: bulk-reads its used range in row chunks and feeds formula cells to the engine.</summary>
    public int IndexSheet(Workbook wb, Worksheet ws)
    {
        var writer = Engine.BeginSheet(wb.Name, ws.Name);
        int count = 0;
        XlRange used = ws.UsedRange;
        int r0 = used.Row, c0 = used.Column;
        int nr = used.Rows.Count, nc = used.Columns.Count;

        _reads.Walk(nr, nc, (off, rows) =>
        {
            XlRange part = rows == nr ? used : ws.Range[ws.Cells[r0 + off, c0], ws.Cells[r0 + off + rows - 1, c0 + nc - 1]];
            var v = ReadFormulasR1C1(part);
            if (v is object[,] arr)
            {
                int lr = arr.GetLowerBound(0), lc = arr.GetLowerBound(1);
                int ur = arr.GetUpperBound(0), uc = arr.GetUpperBound(1);
                for (int i = lr; i <= ur; i++)
                    for (int j = lc; j <= uc; j++)
                        if (arr[i, j] is string s && s.Length > 1 && s[0] == '=')
                        {
                            writer.Add(r0 + off + (i - lr), c0 + (j - lc), s);
                            count++;
                        }
            }
            else if (v is string s1 && s1.Length > 1 && s1[0] == '=')
            {
                writer.Add(r0 + off, c0, s1);
                count++;
            }
            return true;
        });
        writer.Commit();
        return count;
    }

    /// <summary>Indexes every dirty or new sheet of every open workbook. Returns the number of sheets indexed.</summary>
    public int EnsureIndexed(Action<string>? progress = null)
    {
        var books = ReconcileBooks();
        int done = 0;
        foreach (var (book, sheet) in Engine.SheetsNeedingIndex())
        {
            var wb = FindBook(books, book);
            if (wb == null) continue;
            progress?.Invoke($"Indexing {book} › {sheet}…");
            if (wb.Worksheets[sheet] is Worksheet ws)
            {
                IndexSheet(wb, ws);
                done++;
            }
        }
        return done;
    }

    /// <summary>Indexes at most one pending sheet (for idle-time pre-building). Returns false when nothing is left.</summary>
    public bool IndexNextPendingSheet()
    {
        var books = ReconcileBooks();
        foreach (var (book, sheet) in Engine.SheetsNeedingIndex())
        {
            var wb = FindBook(books, book);
            if (wb?.Worksheets[sheet] is Worksheet ws)
            {
                IndexSheet(wb, ws);
                return true;
            }
        }
        return false;
    }

    private static Workbook? FindBook(List<Workbook> books, string name)
    {
        foreach (var wb in books)
            if (string.Equals(wb.Name, name, StringComparison.OrdinalIgnoreCase)) return wb;
        return null;
    }

    // ---- cells, values, navigation ------------------------------------------------------------

    public static (Workbook Book, Worksheet Sheet, XlRange Cell)? ActiveCell()
    {
        var app = App;
        if (app.ActiveSheet is not Worksheet ws) return null;
        XlRange? cell = app.ActiveCell;
        if (cell == null) return null;
        return ((Workbook)ws.Parent, ws, cell);
    }

    /// <summary>First area of the current selection (or the active cell if the selection is not a range).</summary>
    public static (Workbook Book, Worksheet Sheet, CellRect Rect)? ActiveSelection()
    {
        var app = App;
        if (app.ActiveSheet is not Worksheet ws) return null;
        if (app.Selection is not XlRange sel)
        {
            var ac = app.ActiveCell;
            if (ac == null) return null;
            sel = ac;
        }
        return ((Workbook)ws.Parent, ws, RectOf(sel.Areas[1]));
    }

    public static CellRect RectOf(XlRange r) =>
        new(r.Row, r.Column, r.Row + r.Rows.Count - 1, r.Column + r.Columns.Count - 1);

    public static XlRange? GetRange(string book, string sheet, CellRect rect)
    {
        try
        {
            var wb = App.Workbooks[book];
            var ws = (Worksheet)wb.Worksheets[sheet];
            return ws.Range[A1.Address(rect, absolute: false)];
        }
        catch (COMException)
        {
            return null;
        }
    }

    public static bool Goto(string book, string sheet, CellRect rect)
    {
        var r = GetRange(book, sheet, rect);
        if (r == null) return false;
        try
        {
            App.Goto(r, false);
            return true;
        }
        catch (COMException)
        {
            return false;
        }
    }

    /// <summary>
    /// Where a range is drawn on screen, for the row/column highlight. Null when it is not on the active sheet of the
    /// active window, or that window is minimized. Each pane converts with its own <c>PointsToScreenPixelsX/Y</c>,
    /// which follow that pane's scrolling and the zoom (the Window-level methods ignore panes).
    /// About 15 in-process calls for an unsplit window.
    /// </summary>
    public static CrosshairView? ScreenGeometry(string book, string sheet, CellRect rect)
    {
        var wnd = App.ActiveWindow;
        if (wnd?.ActiveSheet is not Worksheet ws || !string.Equals(ws.Name, sheet, StringComparison.OrdinalIgnoreCase)
            || !string.Equals(((Workbook)ws.Parent).Name, book, StringComparison.OrdinalIgnoreCase)) return null;
        IntPtr hwnd = WindowHandle(wnd);
        if (NativeWindows.IsIconic(hwnd)) return null;

        XlRange target = ws.Range[A1.Address(rect, absolute: false)];
        var (left, top, right, bottom) = Points(target);
        var panes = new List<(PixelRect, PixelRect)>();
        var all = wnd.Panes;
        for (int i = 1; i <= all.Count; i++)
        {
            Pane pane = all[i];
            var (vl, vt, vr, vb) = Points(pane.VisibleRange);
            panes.Add((Pixels(pane, vl, vt, vr, vb), Pixels(pane, left, top, right, bottom)));
        }
        var viewport = NativeWindows.GridViewport(hwnd, wnd.DisplayVerticalScrollBar, wnd.DisplayHorizontalScrollBar || wnd.DisplayWorkbookTabs);
        return new CrosshairView(hwnd, viewport, panes);
    }

    private static (double Left, double Top, double Right, double Bottom) Points(XlRange r)
    {
        double left = Convert.ToDouble(r.Left, CultureInfo.InvariantCulture), top = Convert.ToDouble(r.Top, CultureInfo.InvariantCulture);
        return (left, top, left + Convert.ToDouble(r.Width, CultureInfo.InvariantCulture), top + Convert.ToDouble(r.Height, CultureInfo.InvariantCulture));
    }

    private static PixelRect Pixels(Pane pane, double left, double top, double right, double bottom) => new(
        pane.PointsToScreenPixelsX(Round(left)), pane.PointsToScreenPixelsY(Round(top)),
        pane.PointsToScreenPixelsX(Round(right)), pane.PointsToScreenPixelsY(Round(bottom)));

    private static int Round(double points) => (int)Math.Round(Math.Min(points, int.MaxValue));

    /// <summary>The window's own top-level handle (Excel 2013+ gives each workbook window one), else Excel's.</summary>
    private static IntPtr WindowHandle(Window wnd)
    {
        try
        {
            // Late-bound: Window.Hwnd does not exist before Excel 2013.
            if (wnd.GetType().InvokeMember("Hwnd", BindingFlags.GetProperty, null, wnd, null) is int h && h != 0) return new IntPtr(h);
        }
        catch (Exception ex) when (ex is TargetInvocationException || ex is COMException || ex is MissingMethodException)
        {
        }
        return new IntPtr(App.Hwnd);
    }

    /// <summary>Win32 lookups of Excel's window layout: XLMAIN › XLDESK › EXCEL7 (the workbook's grid window).</summary>
    private static class NativeWindows
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left, Top, Right, Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X, Y;
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindowEx(IntPtr parent, IntPtr childAfter, string className, string? windowName);

        [DllImport("user32.dll")]
        private static extern bool GetClientRect(IntPtr hWnd, out RECT rect);

        [DllImport("user32.dll")]
        private static extern bool ClientToScreen(IntPtr hWnd, ref POINT point);

        [DllImport("user32.dll")]
        public static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern int GetSystemMetricsForDpi(int index, uint dpi);

        private const int SM_CXVSCROLL = 2, SM_CYHSCROLL = 3;

        /// <summary>A scroll bar's thickness at the grid window's own DPI (it differs per monitor); system DPI before Windows 10 1607.</summary>
        private static int ScrollBarSize(IntPtr grid, int metric, int systemDpiValue)
        {
            try
            {
                uint dpi = GetDpiForWindow(grid);
                if (dpi > 0) return GetSystemMetricsForDpi(metric, dpi);
            }
            catch (EntryPointNotFoundException)
            {
            }
            return systemDpiValue;
        }

        /// <summary>
        /// The grid window's client area on screen, less the scroll bars and sheet-tab strip at its right and bottom,
        /// so the highlight never covers them or spills outside the window. Falls back to the main window's client area.
        /// </summary>
        public static PixelRect GridViewport(IntPtr main, bool verticalBar, bool bottomStrip)
        {
            var desk = FindWindowEx(main, IntPtr.Zero, "XLDESK", null);
            var grid = desk == IntPtr.Zero ? IntPtr.Zero : FindWindowEx(desk, IntPtr.Zero, "EXCEL7", null);
            var hwnd = grid != IntPtr.Zero ? grid : main;
            if (!GetClientRect(hwnd, out var rc)) return default;
            var origin = new POINT();
            if (!ClientToScreen(hwnd, ref origin)) return default;
            int right = origin.X + rc.Right - (grid != IntPtr.Zero && verticalBar
                ? ScrollBarSize(grid, SM_CXVSCROLL, System.Windows.Forms.SystemInformation.VerticalScrollBarWidth) : 0);
            int bottom = origin.Y + rc.Bottom - (grid != IntPtr.Zero && bottomStrip
                ? ScrollBarSize(grid, SM_CYHSCROLL, System.Windows.Forms.SystemInformation.HorizontalScrollBarHeight) : 0);
            return new PixelRect(origin.X, origin.Y, right, bottom);
        }
    }

    /// <summary>Evaluates an A1 reference expression (OFFSET/INDIRECT) on a sheet. Returns the referenced range, if any.</summary>
    public static (string Book, string Sheet, CellRect Rect)? EvaluateReference(string book, string sheet, string expressionA1)
    {
        var t = EvaluateTarget(book, sheet, expressionA1);
        return t.IsRange ? (t.Book, t.Sheet, t.Rect) : null;
    }

    /// <summary>The range an A1 reference expression points at, or the error it evaluates to (e.g. OFFSET off the sheet gives #REF!).</summary>
    public static DynamicTarget EvaluateTarget(string book, string sheet, string expressionA1)
    {
        try
        {
            var ws = (Worksheet)App.Workbooks[book].Worksheets[sheet];
            switch (ws.Evaluate(expressionA1))
            {
                case XlRange r:
                    var area = r.Areas[1];
                    var tws = area.Worksheet;
                    return new DynamicTarget(((Workbook)tws.Parent).Name, tws.Name, RectOf(area));
                case int code:
                    return new DynamicTarget(ErrorText(code));
            }
        }
        catch (COMException)
        {
        }
        return new DynamicTarget("#VALUE!");
    }

    // ---- ICellReader (error-root tracing) ------------------------------------------------------

    CellRead ICellReader.Read(string book, string sheet, int row, int col)
    {
        var r = GetRange(book, sheet, CellRect.Cell(row, col));
        if (r == null) return default;
        try
        {
            string? error = r.Value2 is int code ? ErrorText(code) : null;
            return new CellRead(ReadFormula(r), error);
        }
        catch (COMException)
        {
            return default;
        }
    }

    IEnumerable<(int Row, int Col)> ICellReader.ErrorCells(string book, string sheet, CellRect rect, string error, int max)
    {
        var found = new List<(int, int)>();
        var r = GetRange(book, sheet, rect);
        if (r == null) return found;
        try
        {
            var ws = r.Worksheet;
            var ur = RectOf(ws.UsedRange);
            if (!ur.Intersects(rect)) return found;
            var clip = ur.Intersect(rect);
            _reads.Walk(clip.Rows, clip.Cols, (off, rows) =>
            {
                var part = new CellRect(clip.Row1 + off, clip.Col1, clip.Row1 + off + rows - 1, clip.Col2);
                switch (ws.Range[A1.Address(part, absolute: false)].Value2)
                {
                    case object[,] values:
                        int lr = values.GetLowerBound(0), lc = values.GetLowerBound(1);
                        for (int i = lr; i <= values.GetUpperBound(0) && found.Count < max; i++)
                            for (int j = lc; j <= values.GetUpperBound(1) && found.Count < max; j++)
                                if (values[i, j] is int code && ErrorText(code) == error)
                                    found.Add((part.Row1 + i - lr, part.Col1 + j - lc));
                        break;
                    case int single when ErrorText(single) == error:
                        found.Add((part.Row1, part.Col1));
                        break;
                }
                return found.Count < max;
            });
        }
        catch (COMException)
        {
        }
        return found;
    }

    DynamicTarget ICellReader.Evaluate(string book, string sheet, string expressionA1) => EvaluateTarget(book, sheet, expressionA1);

    /// <summary>The dynamic-array spill range of an anchor cell (Excel 365), else null.</summary>
    public static CellRect? SpillRange(string book, string sheet, CellRect anchor)
    {
        var cell = GetRange(book, sheet, anchor);
        if (cell == null) return null;
        try
        {
            if (cell.GetType().InvokeMember("SpillingToRange", BindingFlags.GetProperty, null, cell, null) is XlRange spill)
                return RectOf(spill);
        }
        catch (Exception ex) when (ex is TargetInvocationException || ex is COMException || ex is MissingMethodException)
        {
        }
        return null;
    }

    /// <summary>Value and A1 formula of a single cell for display.</summary>
    public (string Value, string? Formula) ReadCell(string book, string sheet, CellRect cell)
    {
        var r = GetRange(book, sheet, cell);
        if (r == null) return ("", null);
        try
        {
            object? v = r.Value2;
            string? f = r.HasFormula is true ? ReadFormulaA1(r) : null;
            return (FormatValue(v), f);
        }
        catch (COMException)
        {
            return ("", null);
        }
    }

    /// <summary>Non-empty cells of a range with value and formula, from bulk reads in chunks (capped at <paramref name="max"/>).</summary>
    public List<(CellRect Cell, string Value, string? FormulaR1C1)> ReadRangeCells(string book, string sheet, CellRect rect, int max, out bool truncated)
    {
        truncated = false;
        bool more = false;
        var list = new List<(CellRect, string, string?)>();
        // Whole columns/rows: clip to the used range first.
        var r = GetRange(book, sheet, rect);
        if (r == null) return list;
        try
        {
            var ws = r.Worksheet;
            var ur = RectOf(ws.UsedRange);
            if (!ur.Intersects(rect)) return list;
            var clip = ur.Intersect(rect);

            _reads.Walk(clip.Rows, clip.Cols, (off, rows) =>
            {
                var part = new CellRect(clip.Row1 + off, clip.Col1, clip.Row1 + off + rows - 1, clip.Col2);
                XlRange pr = ws.Range[A1.Address(part, absolute: false)];
                object? values = pr.Value2;
                object? formulas = ReadFormulasR1C1(pr);
                if (values is object[,] va && formulas is object[,] fa)
                {
                    int lr = va.GetLowerBound(0), lc = va.GetLowerBound(1);
                    for (int i = lr; i <= va.GetUpperBound(0); i++)
                        for (int j = lc; j <= va.GetUpperBound(1); j++)
                            if (!Add(CellRect.Cell(part.Row1 + i - lr, part.Col1 + j - lc), va[i, j], fa[i, j] as string)) return false;
                    return true;
                }
                return Add(part, values, formulas as string);
            });
        }
        catch (COMException)
        {
        }
        truncated = more;
        return list;

        // False once the cap is reached and another non-empty cell remains.
        bool Add(CellRect cell, object? value, string? f)
        {
            bool hasFormula = f != null && f.Length > 1 && f[0] == '=';
            if (value == null && !hasFormula) return true;
            if (list.Count >= max)
            {
                more = true;
                return false;
            }
            list.Add((cell, FormatValue(value), hasFormula ? f : null));
            return true;
        }
    }

    public static string FormatValue(object? v) => v switch
    {
        null => "",
        double d when d != 0 && (Math.Abs(d) >= 1e15 || Math.Abs(d) < 1e-6) => d.ToString("G6", CultureInfo.CurrentCulture),
        double d => d.ToString("#,##0.######", CultureInfo.CurrentCulture),
        string s => s,
        bool b => b ? "TRUE" : "FALSE",
        int code => ErrorText(code),
        object[,] a => $"{{{a.GetLength(0)}×{a.GetLength(1)}}}",
        _ => Convert.ToString(v, CultureInfo.CurrentCulture) ?? "",
    };

    /// <summary>Value2 returns cell errors as Int32 (0x800A07xx); the low word is Excel's xlErr code.</summary>
    internal static string ErrorText(int code) => (code & 0xFFFF) switch
    {
        2000 => "#NULL!",
        2007 => "#DIV/0!",
        2015 => "#VALUE!",
        2023 => "#REF!",
        2029 => "#NAME?",
        2036 => "#NUM!",
        2042 => "#N/A",
        2043 => "#GETTING_DATA",
        2045 => "#SPILL!",
        2046 => "#CONNECT!",
        2047 => "#BLOCKED!",
        2048 => "#UNKNOWN!",
        2049 => "#FIELD!",
        2050 => "#CALC!",
        _ => code.ToString(CultureInfo.InvariantCulture),
    };
}
