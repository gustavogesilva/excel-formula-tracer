using System;
using System.Collections.Generic;
using System.Linq;
using Tracer.Core.Binding;
using Tracer.Core.Graph;
using Tracer.Core.Model;
using Tracer.Core.Parsing;
using Tracer.Core.Text;
using Tracer.Core.Tracing;
using Xunit;

namespace Tracer.Core.Tests;

public class DependencyIndexTests
{
    [Fact]
    public void SolveMatchesBruteForce()
    {
        var rnd = new Random(7);
        for (int iter = 0; iter < 20000; iter++)
        {
            Bound B() => rnd.Next(2) == 0 ? Bound.Abs(rnd.Next(1, 40)) : Bound.Rel(rnd.Next(-6, 7));
            var b1 = B();
            var b2 = B();
            int lo = rnd.Next(10, 25), hi = lo + rnd.Next(0, 10);
            int t1 = rnd.Next(1, 40), t2 = t1 + rnd.Next(0, 4);

            var expected = new List<int>();
            for (int a = lo; a <= hi; a++)
            {
                int f1 = b1.IsRelative ? a + b1.Value : b1.Value;
                int f2 = b2.IsRelative ? a + b2.Value : b2.Value;
                if (Math.Min(f1, f2) <= t2 && Math.Max(f1, f2) >= t1) expected.Add(a);
            }

            bool ok = DependencyIndex.Solve(b1, b2, t1, t2, lo, hi, out int x, out int y);
            var actual = ok ? Enumerable.Range(x, y - x + 1).ToList() : new List<int>();
            Assert.Equal(expected, actual);
        }
    }

    [Fact]
    public void BlocksMergeAcrossAndDown()
    {
        var bb = new BlockBuilder();
        // Rows 2..4, cols 3..6 hold template 1; row 5 cols 3..5 hold template 1 (narrower: new block); col 8 rows 2..4 template 2.
        for (int r = 2; r <= 5; r++)
        {
            for (int c = 3; c <= (r == 5 ? 5 : 6); c++) bb.Add(r, c, 1);
            if (r <= 4) bb.Add(r, 8, 2);
        }
        var blocks = bb.Build().Select(b => (b.Cells.ToString(), b.TemplateId)).ToList();
        Assert.Equal(new[] { ("C2:F4", 1), ("H2:H4", 2), ("C5:E5", 1) }, blocks);
    }

    /// <summary>The compressed rule index must agree with a per-cell brute force on random copied-formula grids.</summary>
    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    public void IndexMatchesBruteForceOracle(int seed)
    {
        var rnd = new Random(seed);
        string[] sheets = { "S1", "S2" };
        const int Rows = 45, Cols = 30;

        string RandBound(char axis, int max)
        {
            int k = rnd.Next(10);
            if (k < 4) return $"{axis}[{rnd.Next(-6, 7)}]";
            if (k < 5) return axis.ToString();
            return $"{axis}{rnd.Next(1, max + 1)}";
        }
        string RandRef()
        {
            string prefix = rnd.Next(4) switch { 0 => "S2!", 1 => "S1!", _ => "" };
            int k = rnd.Next(10);
            string body = k switch
            {
                0 => $"{RandBound('C', Cols)}:{RandBound('C', Cols)}",
                1 => $"{RandBound('R', Rows)}:{RandBound('R', Rows)}",
                < 6 => RandBound('R', Rows) + RandBound('C', Cols),
                _ => $"{RandBound('R', Rows)}{RandBound('C', Cols)}:{RandBound('R', Rows)}{RandBound('C', Cols)}",
            };
            return prefix + body;
        }

        var templates = Enumerable.Range(0, 30).Select(_ =>
            "=" + string.Join("+", Enumerable.Range(0, rnd.Next(1, 4)).Select(_ => "SUM(" + RandRef() + ")"))).ToArray();

        var grids = sheets.ToDictionary(s => s, _ => new string?[Rows + 1, Cols + 1]);
        foreach (var g in grids.Values)
        {
            for (int k = 0; k < 60; k++)
            {
                int r1 = rnd.Next(1, Rows + 1), c1 = rnd.Next(1, Cols + 1);
                int r2 = Math.Min(Rows, r1 + rnd.Next(0, rnd.Next(2) == 0 ? 1 : 12));
                int c2 = Math.Min(Cols, c1 + rnd.Next(0, 10));
                string? f = rnd.Next(8) == 0 ? null : templates[rnd.Next(templates.Length)];
                for (int r = r1; r <= r2; r++)
                    for (int c = c1; c <= c2; c++) g[r, c] = f;
            }
        }

        var engine = new TraceEngine();
        engine.UpdateBook(new WorkbookModel("Book.xlsx", sheets));
        foreach (var s in sheets)
        {
            var w = engine.BeginSheet("Book.xlsx", s);
            var g = grids[s];
            for (int r = 1; r <= Rows; r++)
                for (int c = 1; c <= Cols; c++)
                    if (g[r, c] != null) w.Add(r, c, g[r, c]!);
            w.Commit();
        }
        var stats = engine.Stats();
        Assert.True(stats.Blocks < stats.FormulaCells, "copied formulas should compress into blocks");

        for (int q = 0; q < 400; q++)
        {
            string qs = sheets[rnd.Next(2)];
            int r1 = rnd.Next(1, Rows + 1), c1 = rnd.Next(1, Cols + 1);
            var target = rnd.Next(3) == 0
                ? new CellRect(r1, c1, Math.Min(Rows, r1 + rnd.Next(0, 5)), Math.Min(Cols, c1 + rnd.Next(0, 5)))
                : CellRect.Cell(r1, c1);
            if (rnd.Next(20) == 0) target = new CellRect(1, c1, Limits.MaxRows, c1);

            var expected = new HashSet<(string, int, int)>();
            foreach (var s in sheets)
            {
                var g = grids[s];
                for (int r = 1; r <= Rows; r++)
                    for (int c = 1; c <= Cols; c++)
                    {
                        if (g[r, c] == null) continue;
                        var t = TemplateParser.Parse(g[r, c]!);
                        foreach (var rf in t.Refs)
                        {
                            string tsheet = rf.PrefixIndex >= 0 ? t.Prefixes[rf.PrefixIndex].Sheet1! : s;
                            if (tsheet == qs && rf.Area.Resolve(r, c).Intersects(target)) expected.Add((s, r, c));
                        }
                    }
            }

            var actual = new HashSet<(string, int, int)>();
            foreach (var item in engine.Dependents("Book.xlsx", qs, target).Items)
                for (int r = item.Rect.Row1; r <= item.Rect.Row2; r++)
                    for (int c = item.Rect.Col1; c <= item.Rect.Col2; c++)
                        actual.Add((item.Sheet, r, c));

            Assert.True(expected.SetEquals(actual),
                $"query {qs}!{target}: missing [{string.Join(",", expected.Except(actual).Take(5))}] extra [{string.Join(",", actual.Except(expected).Take(5))}]");
        }
    }

    [Fact]
    public void WrappingReferencesAreIndexedPerCell()
    {
        var engine = new TraceEngine();
        engine.UpdateBook(new WorkbookModel("B", new[] { "S" }));
        var w = engine.BeginSheet("B", "S");
        for (int c = 1; c <= 3; c++) w.Add(1, c, "=R[-1]C"); // row 1 looks "up" to the last row
        w.Commit();
        var items = engine.Dependents("B", "S", CellRect.Cell(Limits.MaxRows, 2)).Items;
        Assert.Equal("B1", Assert.Single(items).Rect.ToString());
    }

    [Fact]
    public void ReindexingASheetReplacesItsRules()
    {
        var engine = new TraceEngine();
        engine.UpdateBook(new WorkbookModel("B", new[] { "S" }));
        var w = engine.BeginSheet("B", "S");
        w.Add(1, 2, "=RC[-1]");
        w.Commit();
        Assert.Single(engine.Dependents("B", "S", CellRect.Cell(1, 1)).Items);

        engine.MarkSheetDirty("B", "S");
        Assert.Equal(1, engine.Dependents("B", "S", CellRect.Cell(1, 1)).UnindexedSheets);
        w = engine.BeginSheet("B", "S");
        w.Add(1, 2, "=R5C5");
        w.Commit();
        Assert.Empty(engine.Dependents("B", "S", CellRect.Cell(1, 1)).Items);
        Assert.Single(engine.Dependents("B", "S", CellRect.Cell(5, 5)).Items);
    }
}
