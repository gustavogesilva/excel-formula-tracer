using System;
using System.Collections.Generic;

namespace Tracer.AddIn.UI;

/// <summary>A screen rectangle in pixels; <see cref="Right"/> and <see cref="Bottom"/> are exclusive.</summary>
internal readonly struct PixelRect : IEquatable<PixelRect>
{
    public PixelRect(int left, int top, int right, int bottom)
    {
        Left = left;
        Top = top;
        Right = right;
        Bottom = bottom;
    }

    public int Left { get; }
    public int Top { get; }
    public int Right { get; }
    public int Bottom { get; }
    public int Width => Right - Left;
    public int Height => Bottom - Top;
    public bool IsEmpty => Width <= 0 || Height <= 0;

    public PixelRect Intersect(PixelRect o)
    {
        var r = new PixelRect(Math.Max(Left, o.Left), Math.Max(Top, o.Top), Math.Min(Right, o.Right), Math.Min(Bottom, o.Bottom));
        return r.IsEmpty ? default : r;
    }

    /// <summary>The same rectangle relative to <paramref name="origin"/>'s top-left corner.</summary>
    public PixelRect RelativeTo(PixelRect origin) =>
        IsEmpty ? default : new PixelRect(Left - origin.Left, Top - origin.Top, Right - origin.Left, Bottom - origin.Top);

    public bool Equals(PixelRect o) => Left == o.Left && Top == o.Top && Right == o.Right && Bottom == o.Bottom;
    public override bool Equals(object? obj) => obj is PixelRect r && Equals(r);
    public override int GetHashCode() => unchecked(((Left * 397 ^ Top) * 397 ^ Right) * 397 ^ Bottom);
    public override string ToString() => $"({Left},{Top})-({Right},{Bottom})";
}

/// <summary>
/// The crosshair for a traced cell: a row band across the visible grid, a column band down it, and the cell
/// itself, all clipped to the grid. A cell scrolled off screen can still show one band (its column, say).
/// Pure, unit-tested.
/// </summary>
internal static class CrosshairGeometry
{
    public static (PixelRect Row, PixelRect Column, PixelRect Cell) Bands(PixelRect grid, PixelRect target)
    {
        if (grid.IsEmpty || target.Width <= 0 || target.Height <= 0) return default;
        var row = grid.Intersect(new PixelRect(grid.Left, target.Top, grid.Right, target.Bottom));
        var column = grid.Intersect(new PixelRect(target.Left, grid.Top, target.Right, grid.Bottom));
        var cell = grid.Intersect(target);
        return (row, column, cell);
    }

    public static bool IsVisible((PixelRect Row, PixelRect Column, PixelRect Cell) bands) =>
        !bands.Row.IsEmpty || !bands.Column.IsEmpty;

    /// <summary>
    /// Bands for every pane of a window (frozen or split panes give up to four), each pane's grid first clipped to
    /// the viewport: a pane's last row and column can be partly visible and run under the scroll bars.
    /// Panes that show none of the target are left out.
    /// </summary>
    public static List<(PixelRect Row, PixelRect Column, PixelRect Cell)> Layout(PixelRect viewport, IEnumerable<(PixelRect Grid, PixelRect Target)> panes)
    {
        var list = new List<(PixelRect, PixelRect, PixelRect)>();
        foreach (var (grid, target) in panes)
        {
            var bands = Bands(viewport.Intersect(grid), target);
            if (IsVisible(bands)) list.Add(bands);
        }
        return list;
    }
}

/// <summary>
/// When the highlight re-reads Excel's geometry. Each timer tick checks only cheap Win32 state, the time of the last
/// keyboard/mouse input and the Excel window's rectangle, and re-reads (through Excel's object model) only after one
/// of them changed: a couple of times, so scrolling settles, then not again until the next change. An idle user
/// means no calls into Excel at all, so a long recalculation (data tables recalculate asynchronously) is never
/// interrupted by the highlight. Pure, unit-tested.
/// </summary>
internal sealed class FollowGate
{
    private const int SettleTicks = 2;
    private uint _lastInput;
    private PixelRect _window;
    private int _pending;

    /// <summary>Starts over from the state right after a re-read.</summary>
    public void Reset(uint lastInput, PixelRect window)
    {
        _lastInput = lastInput;
        _window = window;
        _pending = 0;
    }

    /// <summary>True when this tick should re-read the geometry.</summary>
    public bool Tick(uint lastInput, PixelRect window)
    {
        if (lastInput != _lastInput || !window.Equals(_window))
        {
            _lastInput = lastInput;
            _window = window;
            _pending = SettleTicks;
        }
        if (_pending == 0) return false;
        _pending--;
        return true;
    }

    /// <summary>The re-read couldn't happen (Excel busy or calculating): try again on the next tick.</summary>
    public void Retry() => _pending = Math.Max(_pending, 1);
}

/// <summary>
/// Where a traced range is drawn: the Excel window showing it, that window's grid area (cells only, no headings
/// or scroll bars), and per pane its visible cells and the range, all in screen pixels.
/// </summary>
internal sealed class CrosshairView
{
    public CrosshairView(IntPtr window, PixelRect viewport, IReadOnlyList<(PixelRect Grid, PixelRect Target)> panes)
    {
        Window = window;
        Viewport = viewport;
        Panes = panes;
    }

    public IntPtr Window { get; }
    public PixelRect Viewport { get; }
    public IReadOnlyList<(PixelRect Grid, PixelRect Target)> Panes { get; }
}
