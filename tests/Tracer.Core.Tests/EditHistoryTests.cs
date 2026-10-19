using Tracer.AddIn.UI;
using Xunit;

namespace Tracer.Core.Tests;

public class EditHistoryTests
{
    private static FormulaEdit E(int row, string previous) => new("Book.xlsx", "Calc", row, 1, previous);

    [Fact]
    public void UndoesInReverseOrder()
    {
        var h = new EditHistory();
        h.Push(E(1, "=A1"));
        h.Push(E(2, ""));
        Assert.True(h.TryPop(out var last));
        Assert.Equal(2, last.Row);
        Assert.Equal("", last.Previous);
        Assert.True(h.TryPop(out var first));
        Assert.Equal("=A1", first.Previous);
        Assert.False(h.TryPop(out _));
    }

    [Fact]
    public void DropsTheOldestBeyondCapacity()
    {
        var h = new EditHistory(capacity: 3);
        for (int i = 1; i <= 5; i++) h.Push(E(i, "=" + i));
        Assert.Equal(3, h.Count);
        Assert.True(h.TryPop(out var e5));
        Assert.True(h.TryPop(out _));
        Assert.True(h.TryPop(out var e3));
        Assert.Equal(5, e5.Row);
        Assert.Equal(3, e3.Row);
        Assert.False(h.TryPop(out _));
    }
}
