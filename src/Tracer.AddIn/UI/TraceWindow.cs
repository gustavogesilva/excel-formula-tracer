using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using Tracer.AddIn.Excel;
using Tracer.Core.Tracing;
using NavMove = Tracer.AddIn.UI.ListNavigation.Move;

namespace Tracer.AddIn.UI;

/// <summary>
/// Keyboard-first trace window. Runs on Excel's main thread (no COM marshalling); hidden, not disposed, between uses.
/// Excel's message loop does not run WinForms pre-processing, so all keys are handled in <see cref="OnKeyDown"/>
/// via KeyPreview, with <see cref="ProcessCmdKey"/> covering the case where a WinForms loop is running.
/// Row 0 is the traced cell itself (see <see cref="ListNavigation"/>).
/// Every event handler is guarded: an exception escaping into WinForms would show its crash dialog inside Excel,
/// so failures are reported on the status line instead.
/// </summary>
internal sealed class TraceWindow : Form
{
    private static readonly Color[] Palette =
    {
        Color.FromArgb(0, 92, 197), Color.FromArgb(0, 128, 56), Color.FromArgb(128, 57, 191), Color.FromArgb(196, 30, 58),
        Color.FromArgb(0, 122, 128), Color.FromArgb(191, 96, 0), Color.FromArgb(176, 0, 128), Color.FromArgb(110, 90, 0),
    };

    private readonly TraceController _controller;
    private readonly Label _title;
    private readonly RichTextBox _formula;
    private readonly ListView _list;
    private readonly Label _status;
    private readonly Font _rootFont;
    private readonly Stack<Level> _history = new();
    private Level _level = new();
    private ListViewItem?[] _rows = { null };
    private Dictionary<int, Color> _spanColors = new();
    private SavedPosition? _origin;
    private bool _suppressNavigation;
    /// <summary>The formula box is open for editing (F2): keys go to it, except Enter / Esc / Alt+Enter.</summary>
    private bool _editing;
    private static readonly Color EditBackColor = Color.FromArgb(255, 251, 230);

    public TraceWindow(TraceController controller)
    {
        _controller = controller;
        // The version in the title tells at a glance which build Excel actually loaded.
        Text = "Formula Tracer " + typeof(TraceWindow).Assembly.GetName().Version.ToString(3);
        Font = new Font("Segoe UI", 9f);
        AutoScaleMode = AutoScaleMode.Dpi;
        FormBorderStyle = FormBorderStyle.SizableToolWindow;
        ShowInTaskbar = false;
        TopMost = true;
        KeyPreview = true;
        StartPosition = FormStartPosition.Manual;
        MinimumSize = new Size(480, 240);
        Size = new Size(820, 420);
        _rootFont = new Font(Font, FontStyle.Bold);

        _list = new SafeListView(ReportInternalError)
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            HideSelection = false,
            MultiSelect = false,
            VirtualMode = true,
            VirtualListSize = 1,
            HeaderStyle = ColumnHeaderStyle.Nonclickable,
            BorderStyle = BorderStyle.None,
        };
        _list.Columns.Add("Location", 260);
        _list.Columns.Add("Value", 130);
        _list.Columns.Add("Formula", 400);
        _list.RetrieveVirtualItem += (_, e) => e.Item = SafeRow(e.ItemIndex);
        _list.SelectedIndexChanged += (_, _) => Guard(OnSelectionChanged);
        _list.DoubleClick += (_, _) => Guard(() => StepInto(_level.Mode));

        _formula = new RichTextBox
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            BorderStyle = BorderStyle.None,
            DetectUrls = false,
            ScrollBars = RichTextBoxScrollBars.Vertical,
            Font = new Font("Consolas", 10f),
            BackColor = SystemColors.Window,
            TabStop = false,
        };
        // MouseDown with Clicks == 2 is raised reliably for rich edit controls, unlike DoubleClick.
        _formula.MouseDown += (_, e) =>
        {
            if (e.Clicks == 2 && !_editing) Guard(BeginEdit);
        };
        _title = new Label { Dock = DockStyle.Top, Height = 22, Font = new Font(Font, FontStyle.Bold), AutoEllipsis = true, Padding = new Padding(4, 4, 4, 0) };
        var header = new Panel { Dock = DockStyle.Top, Height = 74, Padding = new Padding(4, 0, 4, 4) };
        header.Controls.Add(_formula);
        header.Controls.Add(_title);

        _status = new Label { Dock = DockStyle.Bottom, Height = 22, AutoEllipsis = true, Padding = new Padding(4, 3, 4, 0), ForeColor = SystemColors.GrayText };

        // Dock order: the Fill control first, then the edges.
        Controls.Add(_list);
        Controls.Add(header);
        Controls.Add(_status);

        try
        {
            // Own the window to Excel so it minimizes with it.
            NativeOwner.Assign(this, ExcelDnaUtil.WindowHandle);
        }
        catch (Exception)
        {
        }
    }

    // ---- showing ----------------------------------------------------------------------

    /// <summary>Shows a fresh trace. A null origin while already visible keeps the existing return point.</summary>
    public void Open(Level level, SavedPosition? origin)
    {
        if (origin != null || !Visible) _origin = origin;
        _history.Clear();
        PlaceOnScreen();
        Display(level);
        if (!Visible) Show();
        Activate();
        _list.Focus();
    }

    public void SetStatus(string text)
    {
        _status.Text = text;
        _status.Refresh();
    }

    private void Display(Level level)
    {
        _suppressNavigation = true;
        try
        {
            // Shrink the list before swapping in the new level: while the old size is still set, the ListView
            // repaints and asks for rows that the new (possibly shorter or empty) level does not have.
            _list.SelectedIndices.Clear();
            _list.VirtualListSize = 0;

            _level = level;
            int rows = ListNavigation.RowCount(level.Items.Count);
            _rows = new ListViewItem?[rows];
            _spanColors = new Dictionary<int, Color>();
            foreach (var it in level.Items)
                if (it.SpanStart >= 0 && !_spanColors.ContainsKey(it.SpanStart))
                    _spanColors[it.SpanStart] = Palette[_spanColors.Count % Palette.Length];

            string value = string.IsNullOrEmpty(level.FocusValue) ? "" : "  =  " + level.FocusValue;
            _title.Text = level.Title + value;
            RenderFormula(-1);

            _list.VirtualListSize = rows;
            int row = ListNavigation.Clamp(level.Selected, level.Items.Count);
            _list.SelectedIndices.Add(row);
            _list.FocusedItem = _list.Items[row];
            _list.EnsureVisible(row);
            _list.Invalidate();
        }
        finally
        {
            _suppressNavigation = false;
        }
        UpdateStatus();
    }

    private void UpdateStatus()
    {
        var parts = new List<string> { $"{_level.Mode} · {_level.Items.Count} item{(_level.Items.Count == 1 ? "" : "s")}" };
        if (Settings.Current.ShowTimings && _level.Millis > 0) parts.Add($"{_level.Millis:F1} ms");
        if (!string.IsNullOrEmpty(_level.Note)) parts.Add(_level.Note);
        parts.Add("↑↓ go · → into · ← up/back · Tab in/out · E error root · F2 edit · Enter jump · Esc return");
        _status.Text = string.Join("   ", parts);
    }

    private void RenderFormula(int highlightSpan)
    {
        _formula.SuspendLayout();
        _formula.Text = _level.Formula ?? (_level.Items.Count == 0 && _level.Mode == TraceMode.Precedents ? "(no formula)" : "");
        _formula.SelectAll();
        _formula.SelectionColor = SystemColors.WindowText;
        _formula.SelectionBackColor = SystemColors.Window;
        if (_level.Formula != null)
        {
            foreach (var it in _level.Items)
            {
                if (it.SpanStart < 0 || it.SpanStart + it.SpanLength > _formula.TextLength) continue;
                _formula.Select(it.SpanStart, it.SpanLength);
                if (_spanColors.TryGetValue(it.SpanStart, out var color)) _formula.SelectionColor = color;
                if (it.SpanStart == highlightSpan) _formula.SelectionBackColor = Color.FromArgb(255, 241, 168);
            }
        }
        _formula.Select(0, 0);
        _formula.ResumeLayout();
    }

    // ---- rows -----------------------------------------------------------------------------

    /// <summary>The item shown on a row: the root for row 0, null for a row that no longer exists.</summary>
    private TraceItem? ItemAt(int row)
    {
        if (row == 0) return _level.Root;
        int i = ListNavigation.ItemIndex(row, _level.Items.Count);
        return i >= 0 ? _level.Items[i] : null;
    }

    private int SelectedRow => _list.SelectedIndices.Count > 0 ? _list.SelectedIndices[0] : -1;

    private TraceItem? Selected => SelectedRow >= 0 ? ItemAt(SelectedRow) : null;

    /// <summary>RetrieveVirtualItem must always get an item back, or WinForms throws inside Excel.</summary>
    private ListViewItem SafeRow(int index)
    {
        try
        {
            return Row(index);
        }
        catch (Exception ex)
        {
            return new ListViewItem(new[] { "(unavailable: " + ex.Message + ")", "", "" }) { ForeColor = SystemColors.GrayText };
        }
    }

    private ListViewItem Row(int index)
    {
        if (index >= 0 && index < _rows.Length && _rows[index] is { } cached) return cached;
        var item = ItemAt(index);
        if (item == null) return new ListViewItem(new[] { "", "", "" });

        string value = "", formula = "";
        try
        {
            (value, formula) = index == 0 ? _controller.DisplayRoot(_level) : _controller.Display(_level, item);
        }
        catch (Exception ex) when (ex is COMException || ex is InvalidCastException || ex is ArgumentException)
        {
            // Excel busy, or the sheet/workbook went away: show the row without a value.
        }

        ListViewItem row;
        if (index == 0)
        {
            row = new ListViewItem(new[] { "▾ " + item.Label, value, formula }) { Font = _rootFont };
            if (!item.HasTarget) row.ForeColor = SystemColors.GrayText;
        }
        else
        {
            string label = item.Kind switch
            {
                TraceItemKind.Name => item.Label + "  →  " + (item.Children?.Count == 1 ? item.Children[0].Label : item.Expression),
                TraceItemKind.Table => item.Label + "  →  " + item.Expression,
                TraceItemKind.External => item.Label + "  (closed)",
                TraceItemKind.Dynamic when !item.HasTarget => item.Label + "  (→ to evaluate)",
                _ => item.IsAddressOnly ? item.Label + "  (address only)" : item.Label,
            };
            if (item.Detail != null) label += "  — " + item.Detail;
            row = new ListViewItem(new[] { "    " + label, value, formula });
            if (item.SpanStart >= 0 && _spanColors.TryGetValue(item.SpanStart, out var color)) row.ForeColor = color;
            else if (!item.HasTarget) row.ForeColor = SystemColors.GrayText;
        }
        if (index >= 0 && index < _rows.Length) _rows[index] = row;
        return row;
    }

    private void SelectRow(int row)
    {
        if (row < 0 || row >= _list.VirtualListSize) return;
        _list.SelectedIndices.Clear();
        _list.SelectedIndices.Add(row);
        _list.FocusedItem = _list.Items[row];
        _list.EnsureVisible(row);
    }

    private void OnSelectionChanged()
    {
        // Clicking the list while editing abandons the edit (the formula box is about to be re-rendered).
        if (_editing) CancelEdit();
        int row = SelectedRow;
        var item = row >= 0 ? ItemAt(row) : null;
        if (item == null) return;
        _level.Selected = row;
        RenderFormula(row == 0 ? -1 : item.SpanStart);
        if (_suppressNavigation || !item.HasTarget) return;
        ExcelHost.Goto(item.Book, item.Sheet, item.Rect);
        // Goto can activate another workbook window; take focus back so arrows keep working.
        Activate();
        _list.Focus();
    }

    // ---- keys ---------------------------------------------------------------------------

    protected override void OnKeyDown(KeyEventArgs e)
    {
        if (HandleKey(e.KeyData))
        {
            e.Handled = true;
            e.SuppressKeyPress = true;
            return;
        }
        base.OnKeyDown(e);
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData) =>
        HandleKey(keyData) || base.ProcessCmdKey(ref msg, keyData);

    private bool HandleKey(Keys key)
    {
        if (_editing)
        {
            // Everything else (arrows, typing, Ctrl+Z/C/V inside the box) belongs to the text box.
            Action? edit = key switch
            {
                Keys.Enter => ApplyEdit,
                Keys.Escape => CancelEdit,
                Keys.Alt | Keys.Enter => () => _formula.SelectedText = "\n",
                _ => null,
            };
            if (edit == null) return false;
            Guard(edit);
            return true;
        }

        // The user's tracing shortcuts (Tracer › Shortcuts…) work inside the window too and take priority.
        Action? action = Settings.Current.Shortcuts.Find((int)key) switch
        {
            TraceAction.TracePrecedents => () => StepInto(TraceMode.Precedents),
            TraceAction.TraceDependents => () => StepInto(TraceMode.Dependents),
            TraceAction.TraceError => TraceError,
            TraceAction.JumpBack => () =>
            {
                CloseWindow(restore: false);
                _controller.JumpBack();
            },
            _ => key switch
            {
                Keys.Escape => () => CloseWindow(restore: true),
                Keys.Enter => JumpAndClose,
                Keys.Right => () => StepInto(_level.Mode),
                Keys.Left => StepOut,
                Keys.Back => GoBack,
                Keys.Tab or (Keys.Shift | Keys.Tab) => () => Navigate(() => _controller.Toggle(_level), null),
                Keys.E => TraceError,
                Keys.F2 => BeginEdit,
                Keys.Control | Keys.Z => UndoEdit,
                Keys.Control | Keys.C => () =>
                {
                    if (Selected is { } s) Clipboard.SetText(s.Label);
                },
                _ => null,
            },
        };
        if (action == null) return false;
        Guard(action);
        return true;
    }

    /// <summary>
    /// → (or Ctrl+Q / Ctrl+Shift+Q): on the root row in the same direction, step down to the first item;
    /// on an item (or the root in the other direction), trace into it.
    /// </summary>
    private void StepInto(TraceMode mode)
    {
        int row = SelectedRow;
        var (move, target) = ListNavigation.Right(row, _level.Items.Count);
        if (row == 0 && mode != _level.Mode) move = NavMove.Drill;

        switch (move)
        {
            case NavMove.SelectRow:
                SelectRow(target);
                break;
            case NavMove.Drill when Selected is { } item:
                Navigate(() =>
                {
                    var next = _controller.Drill(_level, item, mode, out var reason);
                    if (next == null && reason != null) SetStatus(reason);
                    return next;
                }, item);
                break;
            default:
                SetStatus(_level.Mode == TraceMode.Precedents && _level.Formula == null
                    ? "Nothing to trace into: the cell holds no formula. Tab shows its dependents."
                    : "Nothing to trace into.");
                break;
        }
    }

    /// <summary>←: from an item, back up to the root row; from the root, back to the previous trace.</summary>
    private void StepOut()
    {
        var (move, target) = ListNavigation.Left(SelectedRow);
        if (move == NavMove.SelectRow) SelectRow(target);
        else GoBack();
    }

    /// <summary>E: root of the error in the selected row if it shows one, else in the traced cell.</summary>
    private void TraceError()
    {
        string book = _level.Book, sheet = _level.Sheet;
        var cell = _level.Focus;
        int row = SelectedRow;
        if (row > 0 && Selected is { HasTarget: true } item && TraceController.IsErrorText(SafeRow(row).SubItems[1].Text))
        {
            book = item.Book;
            sheet = item.Sheet;
            cell = item.Rect;
        }
        Navigate(() => _controller.ErrorRoots(book, sheet, cell), null);
    }

    // ---- editing ----------------------------------------------------------------------------

    /// <summary>F2 / double-click on the formula: edit the traced cell's formula in place.</summary>
    private void BeginEdit()
    {
        if (_editing) return;
        if (!_level.Editable || !_level.Focus.IsSingleCell)
        {
            SetStatus("Only the traced cell's formula can be edited here. Press → into a cell first, then F2.");
            return;
        }
        string text = _controller.ReadEditText(_level);
        _editing = true;
        _formula.ReadOnly = false;
        _formula.TabStop = true;
        _formula.BackColor = EditBackColor;
        _formula.Text = text;
        _formula.SelectAll();
        _formula.SelectionColor = SystemColors.WindowText;
        _formula.SelectionBackColor = EditBackColor;
        _formula.Select(text.Length, 0);
        _formula.Focus();
        SetStatus($"Editing {_level.Root.Label}: Enter applies · Esc cancels · Alt+Enter new line. "
                  + "Note: like any add-in edit, this clears Excel's own undo history; Ctrl+Z in this window undoes it.");
    }

    private void EndEdit()
    {
        _editing = false;
        _formula.ReadOnly = true;
        _formula.TabStop = false;
        _formula.BackColor = SystemColors.Window;
        _list.Focus();
    }

    private void CancelEdit()
    {
        EndEdit();
        RenderFormula(SelectedRow > 0 && Selected is { } item ? item.SpanStart : -1);
        UpdateStatus();
    }

    /// <summary>Enter: write the formula (Excel recalculates as if typed), then re-trace the same cell in place.</summary>
    private void ApplyEdit()
    {
        var sw = Stopwatch.StartNew();
        string label = _level.Root.Label;
        Level? next;
        try
        {
            next = _controller.ApplyEdit(_level, _formula.Text);
        }
        catch (COMException)
        {
            // Stay in the editor so the text can be fixed.
            SetStatus("Excel rejected that formula. Check the syntax: this box uses English function names and commas, as shown.");
            return;
        }
        EndEdit();
        if (next == null)
        {
            CancelEdit();
            return;
        }
        next.Millis = sw.Elapsed.TotalMilliseconds;
        Display(next);
        SetStatus($"Updated {label} in {next.Millis:F1} ms (including Excel's recalculation) · Ctrl+Z undoes");
    }

    /// <summary>Ctrl+Z (not editing): restore the cell changed by the last edit made in this window.</summary>
    private void UndoEdit()
    {
        var edit = _controller.UndoLastEdit();
        if (edit == null)
        {
            SetStatus("Nothing to undo: no formula has been edited in this window.");
            return;
        }
        AfterExternalUndo(edit.Value);
    }

    /// <summary>After an undo (here or via Excel's Ctrl+Z): refresh if this trace shows the restored cell.</summary>
    public void AfterExternalUndo(FormulaEdit edit)
    {
        if (_editing) EndEdit();
        bool showsCell = string.Equals(_level.Book, edit.Book, StringComparison.OrdinalIgnoreCase)
            && string.Equals(_level.Sheet, edit.Sheet, StringComparison.OrdinalIgnoreCase)
            && _level.Focus.IsSingleCell && _level.Focus.Row1 == edit.Row && _level.Focus.Col1 == edit.Col;
        if (showsCell && _controller.Rebuild(_level) is { } next) Display(next);
        SetStatus("Restored the previous content of " + Tracer.Core.Text.A1.Qualified(edit.Book, edit.Sheet,
            Tracer.Core.Model.CellRect.Cell(edit.Row, edit.Col), _level.Book, null));
    }

    private void Navigate(Func<Level?> build, TraceItem? from)
    {
        var sw = Stopwatch.StartNew();
        var next = build();
        if (next == null)
        {
            // The item may have been resolved in place (OFFSET evaluated): refresh its row.
            if (from != null) RefreshRow(from);
            return;
        }
        next.Millis = sw.Elapsed.TotalMilliseconds;
        _history.Push(_level);
        Display(next);
    }

    private void RefreshRow(TraceItem item)
    {
        int row = ReferenceEquals(item, _level.Root) ? 0 : _level.Items.IndexOf(item) + 1;
        if (row < 0 || row >= _rows.Length || row >= _list.VirtualListSize) return;
        _rows[row] = null;
        _list.RedrawItems(row, row, false);
    }

    private void GoBack()
    {
        if (_history.Count == 0)
        {
            SetStatus("This is the first trace; Esc returns to where you started.");
            return;
        }
        Display(_history.Pop());
        if (Selected is { HasTarget: true } item) ExcelHost.Goto(item.Book, item.Sheet, item.Rect);
        Activate();
    }

    private void JumpAndClose()
    {
        var item = Selected;
        if (item != null && item.HasTarget) ExcelHost.Goto(item.Book, item.Sheet, item.Rect);
        _controller.PushJump(_origin);
        CloseWindow(restore: false);
    }

    private void CloseWindow(bool restore)
    {
        if (_editing) EndEdit();
        Settings.Current.WindowBounds = Bounds;
        Settings.Current.Save();
        Hide();
        if (restore) _origin?.Restore();
        _origin = null;
        try
        {
            ExcelHost.App.ActiveWindow?.Activate();
        }
        catch (COMException)
        {
        }
    }

    /// <summary>
    /// Last-resort net: WinForms can throw inside its own message handling (e.g. a stale native item index in
    /// <c>WmReflectNotify</c>) before any of our handlers run. Uncaught, that shows its crash dialog inside Excel.
    /// </summary>
    protected override void WndProc(ref Message m)
    {
        try
        {
            base.WndProc(ref m);
        }
        catch (Exception ex)
        {
            ReportInternalError(ex);
        }
    }

    private void ReportInternalError(Exception ex)
    {
        try
        {
            _status.Text = "Ignored an internal error: " + ex.Message;
        }
        catch (Exception)
        {
        }
    }

    /// <summary>A ListView whose window procedure never lets an exception escape (see <see cref="WndProc"/>).</summary>
    private sealed class SafeListView : ListView
    {
        private readonly Action<Exception> _report;

        public SafeListView(Action<Exception> report) => _report = report;

        protected override void WndProc(ref Message m)
        {
            try
            {
                base.WndProc(ref m);
            }
            catch (Exception ex)
            {
                _report(ex);
            }
        }
    }

    /// <summary>Runs a UI action; any failure lands on the status line instead of WinForms' crash dialog.</summary>
    private void Guard(Action action)
    {
        try
        {
            action();
        }
        catch (COMException ex)
        {
            SetStatus("Excel is busy (finish editing the cell, then try again): " + ex.Message);
        }
        catch (Exception ex)
        {
            SetStatus("Could not do that: " + ex.Message);
        }
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        if (e.CloseReason == CloseReason.UserClosing)
        {
            e.Cancel = true;
            Guard(() => CloseWindow(restore: false));
            return;
        }
        base.OnFormClosing(e);
    }

    private void PlaceOnScreen()
    {
        if (Visible) return;
        if (Settings.Current.WindowBounds is Rectangle saved && IsOnAnyScreen(saved))
        {
            Bounds = saved;
            return;
        }
        var screen = Screen.FromHandle(ExcelDnaUtil.WindowHandle).WorkingArea;
        Location = new Point(screen.Right - Width - 24, screen.Bottom - Height - 48);
    }

    private static bool IsOnAnyScreen(Rectangle r)
    {
        foreach (var s in Screen.AllScreens)
            if (s.WorkingArea.IntersectsWith(r)) return true;
        return false;
    }

    /// <summary>Sets a WinForms form's owner to a native (Excel) window handle.</summary>
    private static class NativeOwner
    {
        private const int GWL_HWNDPARENT = -8;

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll", EntryPoint = "SetWindowLong")]
        private static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int dwNewLong);

        public static void Assign(Form form, IntPtr owner)
        {
            if (IntPtr.Size == 8) SetWindowLongPtr64(form.Handle, GWL_HWNDPARENT, owner);
            else SetWindowLong32(form.Handle, GWL_HWNDPARENT, owner.ToInt32());
        }
    }
}
