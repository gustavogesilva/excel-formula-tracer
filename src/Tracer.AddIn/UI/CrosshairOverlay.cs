using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Tracer.AddIn.Excel;
using Tracer.Core.Model;

namespace Tracer.AddIn.UI;

/// <summary>
/// Translucent row and column bands over Excel's grid, crossing at the cell the trace window last moved to; the
/// cell itself stays clear. A layered, click-through window that never takes focus, owned by the Excel window so it
/// stays with it. Nothing is written to the workbook, so Excel's undo history and saved state are untouched.
/// While shown, a timer follows scrolling, zoom, resizing and sheet switches; it stops when the overlay is hidden.
/// Any failure just hides the bands, so tracing never breaks because of them.
/// </summary>
internal sealed class CrosshairOverlay : Form
{
    private const int WS_EX_LAYERED = 0x80000, WS_EX_TRANSPARENT = 0x20, WS_EX_TOOLWINDOW = 0x80, WS_EX_NOACTIVATE = 0x08000000;
    private static readonly Color BandColor = Color.FromArgb(255, 200, 0);

    private readonly Timer _timer = new() { Interval = 150 };
    private readonly SolidBrush _band = new(BandColor);
    private readonly SolidBrush _clear;
    private (string Book, string Sheet, CellRect Rect)? _target;
    private List<(PixelRect Row, PixelRect Column, PixelRect Cell)> _bands = new();
    private PixelRect _viewport;
    private IntPtr _owner;

    public CrosshairOverlay()
    {
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.Manual;
        AutoScaleMode = AutoScaleMode.None;
        DoubleBuffered = true;
        // Colour-keyed background: only the bands are drawn, at this opacity.
        BackColor = TransparencyKey = Color.Magenta;
        Opacity = 0.3;
        _clear = new SolidBrush(TransparencyKey);
        _timer.Tick += (_, _) => Follow();
    }

    protected override bool ShowWithoutActivation => true;

    protected override CreateParams CreateParams
    {
        get
        {
            var cp = base.CreateParams;
            cp.ExStyle |= WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE;
            return cp;
        }
    }

    /// <summary>Highlights the row(s) and column(s) of a range, and keeps following it until <see cref="HideOverlay"/>.</summary>
    public void ShowAt(string book, string sheet, CellRect rect)
    {
        _target = (book, sheet, rect);
        Follow();
        _timer.Start();
    }

    public void HideOverlay()
    {
        _target = null;
        _timer.Stop();
        SetBands(new List<(PixelRect, PixelRect, PixelRect)>());
    }

    /// <summary>Recomputes where the target is drawn and repaints only when that changed.</summary>
    private void Follow()
    {
        if (_target is not { } t) return;
        CrosshairView? view;
        try
        {
            view = ExcelHost.ScreenGeometry(t.Book, t.Sheet, t.Rect);
        }
        catch (Exception ex) when (ex is COMException || ex is InvalidCastException || ex is TargetInvocationException)
        {
            return; // Excel is busy (editing a cell, a dialog is open): keep what is shown, try again next tick.
        }
        catch (Exception)
        {
            HideOverlay();
            return;
        }

        try
        {
            if (view == null)
            {
                SetBands(new List<(PixelRect, PixelRect, PixelRect)>());
                return;
            }
            if (view.Window != _owner)
            {
                TraceWindow.NativeOwner.Assign(this, view.Window);
                _owner = view.Window;
            }
            if (!view.Viewport.Equals(_viewport))
            {
                _viewport = view.Viewport;
                Bounds = new Rectangle(_viewport.Left, _viewport.Top, _viewport.Width, _viewport.Height);
            }
            SetBands(CrosshairGeometry.Layout(view.Viewport, view.Panes));
        }
        catch (Exception)
        {
            HideOverlay();
        }
    }

    private void SetBands(List<(PixelRect Row, PixelRect Column, PixelRect Cell)> bands)
    {
        if (bands.Count == 0)
        {
            _bands = bands;
            if (Visible) Hide();
            return;
        }
        if (!bands.SequenceEqual(_bands))
        {
            _bands = bands;
            Invalidate();
        }
        if (!Visible) Show();
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        foreach (var (row, column, _) in _bands)
        {
            Fill(g, _band, row);
            Fill(g, _band, column);
        }
        // The cell itself stays clear, so its value reads normally inside the crosshair.
        foreach (var (_, _, cell) in _bands) Fill(g, _clear, cell);
    }

    private void Fill(Graphics g, Brush brush, PixelRect screen)
    {
        var r = screen.RelativeTo(_viewport);
        if (!r.IsEmpty) g.FillRectangle(brush, r.Left, r.Top, r.Width, r.Height);
    }

    /// <summary>A window procedure exception would show WinForms' crash dialog inside Excel.</summary>
    protected override void WndProc(ref Message m)
    {
        try
        {
            base.WndProc(ref m);
        }
        catch (Exception)
        {
        }
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _timer.Dispose();
            _band.Dispose();
            _clear.Dispose();
        }
        base.Dispose(disposing);
    }
}
