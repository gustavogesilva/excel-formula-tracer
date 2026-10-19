using Tracer.AddIn.UI;
using Xunit;

namespace Tracer.Core.Tests;

public class CrosshairGeometryTests
{
    // Visible grid on screen: x 100..900, y 200..700.
    private static readonly PixelRect Grid = new(100, 200, 900, 700);

    [Fact]
    public void CellInsideTheGrid()
    {
        var (row, col, cell) = CrosshairGeometry.Bands(Grid, new PixelRect(300, 400, 364, 420));
        Assert.Equal(new PixelRect(100, 400, 900, 420), row);
        Assert.Equal(new PixelRect(300, 200, 364, 700), col);
        Assert.Equal(new PixelRect(300, 400, 364, 420), cell);
    }

    [Fact]
    public void CellScrolledOffBelowKeepsItsColumnBand()
    {
        var bands = CrosshairGeometry.Bands(Grid, new PixelRect(300, 900, 364, 920));
        Assert.True(bands.Row.IsEmpty);
        Assert.True(bands.Cell.IsEmpty);
        Assert.Equal(new PixelRect(300, 200, 364, 700), bands.Column);
        Assert.True(CrosshairGeometry.IsVisible(bands));
    }

    [Fact]
    public void CellEntirelyOffScreenShowsNothing()
    {
        var bands = CrosshairGeometry.Bands(Grid, new PixelRect(950, 900, 1000, 920));
        Assert.False(CrosshairGeometry.IsVisible(bands));
    }

    [Fact]
    public void RangeTargetGivesWideBandsClippedToTheGrid()
    {
        var (row, col, cell) = CrosshairGeometry.Bands(Grid, new PixelRect(50, 650, 400, 760));
        Assert.Equal(new PixelRect(100, 650, 900, 700), row);
        Assert.Equal(new PixelRect(100, 200, 400, 700), col);
        Assert.Equal(new PixelRect(100, 650, 400, 700), cell);
    }

    [Fact]
    public void RelativeToTheOverlayWindow()
    {
        Assert.Equal(new PixelRect(200, 200, 264, 220), new PixelRect(300, 400, 364, 420).RelativeTo(Grid));
        Assert.True(default(PixelRect).RelativeTo(Grid).IsEmpty);
    }

    [Fact]
    public void FrozenPanesEachGetTheirPartOfTheBands()
    {
        // Frozen at row 3 / column B: top-left (100..200 × 200..260), top-right, bottom-left, bottom-right.
        // The target scrolled into the bottom-right pane; through the frozen panes it maps outside their grids.
        var viewport = new PixelRect(100, 200, 900, 700);
        var panes = new[]
        {
            (new PixelRect(100, 200, 200, 260), new PixelRect(1300, 2400, 1364, 2420)),
            (new PixelRect(200, 200, 900, 260), new PixelRect(500, 2400, 564, 2420)),
            (new PixelRect(100, 260, 200, 700), new PixelRect(1300, 400, 1364, 420)),
            (new PixelRect(200, 260, 900, 700), new PixelRect(500, 400, 564, 420)),
        };
        var layout = CrosshairGeometry.Layout(viewport, panes);

        Assert.Equal(3, layout.Count); // the top-left pane shows none of it
        Assert.Equal(new PixelRect(500, 200, 564, 260), layout[0].Column);
        Assert.True(layout[0].Row.IsEmpty);
        Assert.Equal(new PixelRect(100, 400, 200, 420), layout[1].Row);
        Assert.True(layout[1].Column.IsEmpty);
        Assert.Equal(new PixelRect(500, 400, 564, 420), layout[2].Cell);
    }

    [Fact]
    public void PartlyVisibleLastRowAndColumnAreClippedToTheViewport()
    {
        // VisibleRange includes the half-shown last column and row, which run under the scroll bars.
        var layout = CrosshairGeometry.Layout(Grid, new[] { (new PixelRect(100, 200, 960, 730), new PixelRect(880, 690, 960, 730)) });
        var (row, col, cell) = Assert.Single(layout);
        Assert.Equal(new PixelRect(100, 690, 900, 700), row);
        Assert.Equal(new PixelRect(880, 200, 900, 700), col);
        Assert.Equal(new PixelRect(880, 690, 900, 700), cell);
    }

    [Fact]
    public void DegenerateInputsShowNothing()
    {
        Assert.False(CrosshairGeometry.IsVisible(CrosshairGeometry.Bands(default, new PixelRect(1, 1, 5, 5))));
        Assert.False(CrosshairGeometry.IsVisible(CrosshairGeometry.Bands(Grid, new PixelRect(300, 400, 300, 420))));
    }
}
