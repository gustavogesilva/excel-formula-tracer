using System;
using System.Collections.Generic;
using System.Linq;
using Tracer.Core.Binding;
using Tracer.Core.Model;
using Tracer.Core.Tracing;
using Xunit;

namespace Tracer.Core.Tests;

public class ErrorTracerTests
{
    private const string Book = "Model.xlsx";

    /// <summary>In-memory workbook: (sheet, row, col) → (R1C1 formula, error value).</summary>
    private sealed class Grid : ICellReader
    {
        public readonly Dictionary<(string, int, int), (string? Formula, string? Error)> Cells = new();
        public readonly Dictionary<string, DynamicTarget> Evaluations = new();
        public int Reads;

        public Grid Set(string sheet, int row, int col, string? formula, string? error = null)
        {
            Cells[(sheet, row, col)] = (formula, error);
            return this;
        }

        public CellRead Read(string book, string sheet, int row, int col)
        {
            Reads++;
            return Cells.TryGetValue((sheet, row, col), out var c) ? new CellRead(c.Formula, c.Error) : default;
        }

        public IEnumerable<(int Row, int Col)> ErrorCells(string book, string sheet, CellRect rect, string error, int max) =>
            Cells.Where(kv => kv.Key.Item1 == sheet && rect.Contains(kv.Key.Item2, kv.Key.Item3) && kv.Value.Error == error)
                 .Select(kv => (kv.Key.Item2, kv.Key.Item3)).OrderBy(x => x.Item1).ThenBy(x => x.Item2).Take(max);

        public DynamicTarget Evaluate(string book, string sheet, string expressionA1) =>
            Evaluations.TryGetValue(expressionA1, out var t) ? t : new DynamicTarget("#VALUE!");
    }

    private static TraceEngine Engine()
    {
        var e = new TraceEngine();
        e.UpdateBook(new WorkbookModel(Book, new[] { "Inputs", "Data", "Calc" },
            new[] { new NameDef("Rate", null, "=Inputs!R2C3") }));
        return e;
    }

    private static string[] Roots(ErrorTraceResult r) => r.Roots.Select(i => i.Label).ToArray();

    [Fact]
    public void FollowsBrokenReferenceChainAcrossSheets()
    {
        var g = new Grid()
            .Set("Data", 2, 2, "=#REF!+1", "#REF!")
            .Set("Calc", 3, 3, "=Data!R2C2*2", "#REF!")
            .Set("Calc", 3, 4, "=RC[-1]+1", "#REF!");
        var r = Engine().TraceErrorRoots(Book, "Calc", 3, 4, g);
        Assert.Equal("#REF!", r.Error);
        var root = Assert.Single(r.Roots);
        Assert.Equal("Data!$B$2", root.Label);
        Assert.Contains("contains #REF!", root.Detail);
        Assert.Contains("2 steps back", root.Detail);
        Assert.Equal("=#REF!+1", root.Expression);
        Assert.Equal(3, r.Visited);
    }

    [Fact]
    public void DeletedSheetReferenceIsARoot()
    {
        var g = new Grid().Set("Calc", 1, 1, "=#REF!R5C5*2", "#REF!");
        var root = Assert.Single(Engine().TraceErrorRoots(Book, "Calc", 1, 1, g).Roots);
        Assert.Contains("contains #REF!$E$5", root.Detail);
        Assert.Equal("=#REF!$E$5*2", root.Expression);
        Assert.Contains("the cell itself", root.Detail);
    }

    [Fact]
    public void FindsErrorInsideRange()
    {
        var g = new Grid().Set("Calc", 1, 2, "=SUM(R1C1:R9C1)", "#REF!");
        for (int r = 1; r <= 9; r++) g.Set("Calc", r, 1, r == 5 ? "=#REF!" : null, r == 5 ? "#REF!" : null);
        Assert.Equal(new[] { "$A$5" }, Roots(Engine().TraceErrorRoots(Book, "Calc", 1, 2, g)));
    }

    [Fact]
    public void ErrorProducedByTheFunctionItself()
    {
        var g = new Grid()
            .Set("Calc", 1, 1, "=1/0", "#DIV/0!")
            .Set("Calc", 1, 2, "=RC[-1]+1", "#DIV/0!");
        var root = Assert.Single(Engine().TraceErrorRoots(Book, "Calc", 1, 2, g).Roots);
        Assert.Equal("$A$1", root.Label);
        Assert.Contains("produced here", root.Detail);
    }

    [Fact]
    public void PrecedentWithADifferentErrorIsNotFollowed()
    {
        var g = new Grid()
            .Set("Calc", 1, 1, "=NA()", "#N/A")
            .Set("Calc", 2, 1, "=#REF!", "#REF!")
            .Set("Calc", 1, 2, "=R1C1+R2C1", "#N/A");
        Assert.Equal(new[] { "$A$1" }, Roots(Engine().TraceErrorRoots(Book, "Calc", 1, 2, g)));
    }

    [Fact]
    public void SeveralRoots()
    {
        var g = new Grid()
            .Set("Calc", 1, 1, "=#REF!", "#REF!")
            .Set("Calc", 2, 1, "=Inputs!#REF!", "#REF!")
            .Set("Calc", 1, 2, "=R1C1+R2C1", "#REF!");
        Assert.Equal(new[] { "$A$1", "$A$2" }, Roots(Engine().TraceErrorRoots(Book, "Calc", 1, 2, g)));
    }

    [Fact]
    public void ConstantErrorValue()
    {
        var g = new Grid()
            .Set("Inputs", 2, 3, null, "#N/A")
            .Set("Calc", 1, 1, "=Rate*2", "#N/A");
        var root = Assert.Single(Engine().TraceErrorRoots(Book, "Calc", 1, 1, g).Roots);
        Assert.Equal("Inputs!$C$2", root.Label);
        Assert.Contains("typed into the cell", root.Detail);
    }

    [Fact]
    public void OffsetEvaluatingToAnErrorIsARoot()
    {
        var g = new Grid()
            .Set("Calc", 1, 1, null, "#REF!") // OFFSET's base: address only, never followed
            .Set("Calc", 1, 2, "=SUM(OFFSET(R1C1,-5,0))", "#REF!");
        g.Evaluations["OFFSET($A$1,-5,0)"] = new DynamicTarget("#REF!");
        var root = Assert.Single(Engine().TraceErrorRoots(Book, "Calc", 1, 2, g).Roots);
        Assert.Equal("$B$1", root.Label);
        Assert.Contains("OFFSET($A$1,-5,0) evaluates to #REF!", root.Detail);
    }

    [Fact]
    public void OffsetEvaluatingToARangeIsFollowed()
    {
        var g = new Grid()
            .Set("Data", 7, 1, "=#REF!", "#REF!")
            .Set("Calc", 1, 2, "=SUM(OFFSET(Data!R1C1,6,0,3,1))", "#REF!");
        g.Evaluations["OFFSET(Data!$A$1,6,0,3,1)"] = new DynamicTarget(Book, "Data", new CellRect(7, 1, 9, 1));
        Assert.Equal(new[] { "Data!$A$7" }, Roots(Engine().TraceErrorRoots(Book, "Calc", 1, 2, g)));
    }

    [Fact]
    public void CycleTerminatesWithoutRoots()
    {
        var g = new Grid()
            .Set("Calc", 1, 1, "=RC[1]", "#REF!")
            .Set("Calc", 1, 2, "=RC[-1]", "#REF!");
        var r = Engine().TraceErrorRoots(Book, "Calc", 1, 1, g);
        Assert.Empty(r.Roots);
        Assert.Equal(2, r.Visited);
    }

    [Fact]
    public void StopsAtTheCellBudget()
    {
        var g = new Grid().Set("Calc", 1, 1, "=#REF!", "#REF!");
        for (int r = 2; r <= 50; r++) g.Set("Calc", r, 1, "=R[-1]C", "#REF!");
        var res = Engine().TraceErrorRoots(Book, "Calc", 50, 1, g, maxCells: 10);
        Assert.True(res.Truncated);
        Assert.Equal(10, res.Visited);
        Assert.Empty(res.Roots);
    }

    [Fact]
    public void NoErrorMeansNothingToTrace()
    {
        var g = new Grid().Set("Calc", 1, 1, "=1+1");
        var r = Engine().TraceErrorRoots(Book, "Calc", 1, 1, g);
        Assert.Null(r.Error);
        Assert.Empty(r.Roots);
    }

    [Fact]
    public void UnknownNameIsTheRootOfNameError()
    {
        var g = new Grid().Set("Calc", 1, 1, "=Ratee*2", "#NAME?");
        var root = Assert.Single(Engine().TraceErrorRoots(Book, "Calc", 1, 1, g).Roots);
        Assert.Contains("Ratee is not a defined name", root.Detail);
    }
}
