using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using Tracer.AddIn.Excel;
using Tracer.Core.Tracing;

namespace Tracer.AddIn.UI;

/// <summary>
/// Keyboard-first trace window. Runs on Excel's main thread (no COM marshalling); hidden, not disposed, between uses.
/// Excel's message loop does not run WinForms pre-processing, so all keys are handled in <see cref="OnKeyDown"/>
/// via KeyPreview, with <see cref="ProcessCmdKey"/> covering the case where a WinForms loop is running.
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
    private readonly Stack<Level> _history = new();
    private Level _level = new();
    private ListViewItem?[] _rows = Array.Empty<ListViewItem?>();
    private Dictionary<int, Color> _spanColors = new();
    private SavedPosition? _origin;
    private bool _suppressNavigation;

    public TraceWindow(TraceController controller)
    {
        _controller = controller;
        Text = "Formula Tracer";
        Font = new Font("Segoe UI", 9f);
        AutoScaleMode = AutoScaleMode.Dpi;
        FormBorderStyle = FormBorderStyle.SizableToolWindow;
        ShowInTaskbar = false;
        TopMost = true;
        KeyPreview = true;
        StartPosition = FormStartPosition.Manual;
        MinimumSize = new Size(480, 240);
        Size = new Size(820, 420);

        _list = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            HideSelection = false,
            MultiSelect = false,
            VirtualMode = true,
            HeaderStyle = ColumnHeaderStyle.Nonclickable,
            BorderStyle = BorderStyle.None,
        };
        _list.Columns.Add("Location", 260);
        _list.Columns.Add("Value", 130);
        _list.Columns.Add("Formula", 400);
        _list.RetrieveVirtualItem += (_, e) => e.Item = Row(e.ItemIndex);
        _list.SelectedIndexChanged += (_, _) => OnSelectionChanged();
        _list.DoubleClick += (_, _) => DrillSelected(_level.Mode);

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
        _level = level;
        _rows = new ListViewItem?[level.Items.Count];
        _spanColors = new Dictionary<int, Color>();
        foreach (var it in level.Items)
            if (it.SpanStart >= 0 && !_spanColors.ContainsKey(it.SpanStart))
                _spanColors[it.SpanStart] = Palette[_spanColors.Count % Palette.Length];

        string value = string.IsNullOrEmpty(level.FocusValue) ? "" : "  =  " + level.FocusValue;
        _title.Text = level.Title + value;
        RenderFormula(-1);

        _suppressNavigation = true;
        try
        {
            _list.SelectedIndices.Clear();
            _list.VirtualListSize = level.Items.Count;
            _list.Invalidate();
            if (level.Items.Count > 0)
            {
                int i = Math.Min(level.Selected, level.Items.Count - 1);
                _list.SelectedIndices.Add(i);
                _list.EnsureVisible(i);
                _list.FocusedItem = _list.Items[i];
            }
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
        parts.Add("↑↓ go · → drill · ← back · Tab in/out · Enter jump · Esc return");
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
                _formula.SelectionColor = _spanColors[it.SpanStart];
                if (it.SpanStart == highlightSpan) _formula.SelectionBackColor = Color.FromArgb(255, 241, 168);
            }
        }
        _formula.Select(0, 0);
        _formula.ResumeLayout();
    }

    private ListViewItem Row(int index)
    {
        var cached = _rows[index];
        if (cached != null) return cached;
        var item = _level.Items[index];
        string value = "", formula = "";
        try
        {
            (value, formula) = _controller.Display(_level, item);
        }
        catch (COMException)
        {
        }
        string label = item.Kind switch
        {
            TraceItemKind.Name => item.Label + "  →  " + (item.Children?.Count == 1 ? item.Children[0].Label : item.Expression),
            TraceItemKind.Table => item.Label + "  →  " + item.Expression,
            TraceItemKind.External => item.Label + "  (closed)",
            TraceItemKind.Dynamic when !item.HasTarget => item.Label + "  (→ to evaluate)",
            _ => item.IsAddressOnly ? item.Label + "  (address only)" : item.Label,
        };
        var row = new ListViewItem(new[] { label, value, formula });
        if (item.SpanStart >= 0 && _spanColors.TryGetValue(item.SpanStart, out var color)) row.ForeColor = color;
        else if (!item.HasTarget) row.ForeColor = SystemColors.GrayText;
        _rows[index] = row;
        return row;
    }

    private TraceItem? Selected =>
        _list.SelectedIndices.Count > 0 && _list.SelectedIndices[0] < _level.Items.Count ? _level.Items[_list.SelectedIndices[0]] : null;

    private void OnSelectionChanged()
    {
        var item = Selected;
        if (item == null) return;
        _level.Selected = _list.SelectedIndices[0];
        RenderFormula(item.SpanStart);
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
        try
        {
            switch (key)
            {
                case Keys.Escape:
                    CloseWindow(restore: true);
                    return true;
                case Keys.Enter:
                    JumpAndClose();
                    return true;
                case Keys.Right:
                    DrillSelected(_level.Mode);
                    return true;
                case Keys.Left:
                case Keys.Back:
                    GoBack();
                    return true;
                case Keys.Tab:
                case Keys.Shift | Keys.Tab:
                    Navigate(() => _controller.Toggle(_level), null);
                    return true;
                case Keys.Control | Keys.Q:
                case Keys.Control | Keys.OemOpenBrackets:
                    DrillSelected(TraceMode.Precedents);
                    return true;
                case Keys.Control | Keys.Shift | Keys.Q:
                case Keys.Control | Keys.OemCloseBrackets:
                    DrillSelected(TraceMode.Dependents);
                    return true;
                case Keys.Control | Keys.C:
                    if (Selected is { } s) Clipboard.SetText(s.Label);
                    return true;
                case Keys.Control | Keys.Back:
                    CloseWindow(restore: false);
                    _controller.JumpBack();
                    return true;
            }
        }
        catch (COMException ex)
        {
            SetStatus("Excel is busy: " + ex.Message);
            return true;
        }
        return false;
    }

    private void DrillSelected(TraceMode mode)
    {
        var item = Selected;
        if (item == null) return;
        Navigate(() =>
        {
            var next = _controller.Drill(_level, item, mode, out var reason);
            if (next == null && reason != null) SetStatus(reason);
            return next;
        }, item);
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
        int i = _level.Items.IndexOf(item);
        if (i < 0) return;
        _rows[i] = null;
        _list.RedrawItems(i, i, false);
    }

    private void GoBack()
    {
        if (_history.Count == 0) return;
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

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        if (e.CloseReason == CloseReason.UserClosing)
        {
            e.Cancel = true;
            CloseWindow(restore: false);
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
