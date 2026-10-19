using Tracer.AddIn.UI;
using Xunit;
using Move = Tracer.AddIn.UI.ListNavigation.Move;

namespace Tracer.Core.Tests;

/// <summary>Trace-window keyboard rules: row 0 is the traced cell, rows 1..n its items.</summary>
public class ListNavigationTests
{
    [Fact]
    public void RootRowAlwaysExists()
    {
        Assert.Equal(1, ListNavigation.RowCount(0));
        Assert.Equal(2, ListNavigation.RowCount(1));
    }

    [Theory]
    [InlineData(0, 3, -1)]
    [InlineData(1, 3, 0)]
    [InlineData(3, 3, 2)]
    [InlineData(4, 3, -1)] // stale row after the list shrank
    [InlineData(1, 0, -1)] // empty trace: only the root exists
    [InlineData(-1, 3, -1)]
    public void RowsMapToItems(int row, int items, int index) => Assert.Equal(index, ListNavigation.ItemIndex(row, items));

    [Fact]
    public void RightFromRootStepsToTheOnlyArgument()
    {
        // =Sheet2!B5: one precedent. → (or ↓) must move to it so Excel follows, without needing Enter.
        Assert.Equal((Move.SelectRow, 1), ListNavigation.Right(0, 1));
    }

    [Fact]
    public void RightOnRootWithNothingToShowDoesNothing() => Assert.Equal((Move.None, 0), ListNavigation.Right(0, 0));

    [Fact]
    public void RightOnAnItemDrillsIntoIt() => Assert.Equal((Move.Drill, 2), ListNavigation.Right(2, 5));

    [Fact]
    public void LeftGoesUpToRootThenBack()
    {
        Assert.Equal((Move.SelectRow, 0), ListNavigation.Left(3));
        Assert.Equal((Move.Back, 0), ListNavigation.Left(0));
    }

    [Theory]
    [InlineData(5, 2, 2)]
    [InlineData(-3, 2, 0)]
    [InlineData(1, 0, 0)]
    public void RememberedSelectionIsClamped(int row, int items, int expected) => Assert.Equal(expected, ListNavigation.Clamp(row, items));
}
