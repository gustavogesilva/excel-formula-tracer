using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Running;
using Tracer.Core.Model;
using Tracer.Core.Parsing;
using Tracer.Core.Tracing;

namespace Tracer.Bench;

/// <summary>
/// <c>dotnet run -c Release</c> runs the scenario (end-to-end timings and retained memory on the synthetic model);
/// <c>dotnet run -c Release -- bdn</c> runs the BenchmarkDotNet micro-benchmarks.
/// Excel COM read time is not included: it can only be measured inside Excel on Windows.
/// </summary>
public static class Program
{
    public static void Main(string[] args)
    {
        if (args.Length > 0 && args[0] == "bdn")
        {
            BenchmarkRunner.Run<CoreBenchmarks>(args: args.Skip(1).ToArray());
            return;
        }
        Scenario();
    }

    private static void Scenario()
    {
        var model = new SyntheticModel();
        int cells = model.FormulaCells;
        var unique = model.Cells.Values.SelectMany(l => l).Select(c => c.Formula).Distinct().ToArray();
        Console.WriteLine($"Synthetic model: {model.Sheets.Count} sheets, {cells:N0} formula cells, {unique.Length:N0} distinct R1C1 formulas");

        // 1. Parse throughput (distinct formulas, no cache).
        for (int w = 0; w < 3; w++) foreach (var f in unique) TemplateParser.Parse(f);
        var sw = Stopwatch.StartNew();
        int reps = 0;
        while (sw.ElapsedMilliseconds < 1000) { foreach (var f in unique) TemplateParser.Parse(f); reps++; }
        sw.Stop();
        double perSec = unique.Length * (double)reps / sw.Elapsed.TotalSeconds;
        Console.WriteLine($"Parse:  {perSec / 1e6:F2} M formulas/s ({1e9 / perSec:F0} ns/formula)");

        // 2. Full index build from cold (cache lookups for every cell + parse of distinct formulas + bind + rules).
        TraceEngine Build()
        {
            var e = new TraceEngine();
            e.UpdateBook(model.Model());
            foreach (var sheet in model.Sheets)
            {
                var w = e.BeginSheet(SyntheticModel.Book, sheet);
                foreach (var (r, c, f) in model.Cells[sheet]) w.Add(r, c, f);
                w.Commit();
            }
            return e;
        }
        Build(); Build(); // JIT warm-up
        var times = new List<double>();
        TraceEngine engine = null!;
        for (int i = 0; i < 5; i++)
        {
            GC.Collect();
            sw.Restart();
            engine = Build();
            sw.Stop();
            times.Add(sw.Elapsed.TotalMilliseconds);
        }
        times.Sort();
        var st = engine.Stats();
        Console.WriteLine($"Index build: median {times[2]:F0} ms, min {times[0]:F0} ms over 5 runs ({cells / times[2] / 1e3:F2} M cells/s)");
        Console.WriteLine($"  {st.FormulaCells:N0} cells -> {st.Blocks:N0} blocks -> {st.Rules:N0} rules, {st.Templates:N0} templates ({(double)st.FormulaCells / st.Blocks:F1} cells/block)");

        // 3. Retained memory of a built engine.
        engine = null!;
        GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
        long before = GC.GetTotalMemory(true);
        engine = Build();
        long after = GC.GetTotalMemory(true);
        Console.WriteLine($"Retained memory: {(after - before) / 1048576.0:F1} MB for {cells:N0} formula cells (includes the template cache)");

        // 4. Dependents query latency on random single cells (inputs and calc cells).
        var rnd = new Random(1);
        var targets = new List<(string Sheet, CellRect Rect)>();
        for (int i = 0; i < 2000; i++)
        {
            string s = model.Sheets[rnd.Next(model.Sheets.Count)];
            targets.Add((s, CellRect.Cell(rnd.Next(2, 300), rnd.Next(3, 88))));
        }
        foreach (var (s, r) in targets.Take(200)) engine.Dependents(SyntheticModel.Book, s, r);
        var lat = new List<double>();
        long hits = 0;
        foreach (var (s, r) in targets)
        {
            long t0 = Stopwatch.GetTimestamp();
            hits += engine.Dependents(SyntheticModel.Book, s, r).Items.Count;
            lat.Add((Stopwatch.GetTimestamp() - t0) * 1000.0 / Stopwatch.Frequency);
        }
        lat.Sort();
        Console.WriteLine($"Dependents query: p50 {lat[lat.Count / 2] * 1000:F0} µs, p99 {lat[lat.Count * 99 / 100] * 1000:F0} µs, max {lat[^1] * 1000:F0} µs ({targets.Count} queries, {hits / (double)targets.Count:F1} items avg)");

        // Worst case: an input cell referenced from everywhere via names.
        sw.Restart();
        var inputDeps = engine.Dependents(SyntheticModel.Book, "Inputs", CellRect.Cell(2, 3));
        sw.Stop();
        Console.WriteLine($"Dependents of a widely used input (Inputs!C2): {sw.Elapsed.TotalMilliseconds * 1000:F0} µs, {inputDeps.Items.Count} blocks");

        // 5. Precedents (engine part only; in Excel add one COM read per cell + one per value shown).
        var cellsList = model.Cells["Calc7"];
        var plat = new List<double>();
        for (int i = 0; i < 2000; i++)
        {
            var (r, c, f) = cellsList[rnd.Next(cellsList.Count)];
            long t0 = Stopwatch.GetTimestamp();
            engine.Precedents(SyntheticModel.Book, "Calc7", r, c, f);
            plat.Add((Stopwatch.GetTimestamp() - t0) * 1000.0 / Stopwatch.Frequency);
        }
        plat.Sort();
        Console.WriteLine($"Precedents (engine): p50 {plat[plat.Count / 2] * 1000:F0} µs, p99 {plat[plat.Count * 99 / 100] * 1000:F0} µs");
    }
}

[MemoryDiagnoser]
[SimpleJob(warmupCount: 2, iterationCount: 8)]
public class CoreBenchmarks
{
    private SyntheticModel _model = null!;
    private string[] _unique = null!;
    private TraceEngine _engine = null!;

    [GlobalSetup]
    public void Setup()
    {
        _model = new SyntheticModel();
        _unique = _model.Cells.Values.SelectMany(l => l).Select(c => c.Formula).Distinct().ToArray();
        _engine = BuildIndex();
    }

    [Benchmark(Description = "Parse all distinct formulas")]
    public int ParseDistinct()
    {
        int n = 0;
        foreach (var f in _unique) n += TemplateParser.Parse(f).Refs.Length;
        return n;
    }

    [Benchmark(Description = "Build full dependents index")]
    public TraceEngine BuildIndex()
    {
        var e = new TraceEngine();
        e.UpdateBook(_model.Model());
        foreach (var sheet in _model.Sheets)
        {
            var w = e.BeginSheet(SyntheticModel.Book, sheet);
            foreach (var (r, c, f) in _model.Cells[sheet]) w.Add(r, c, f);
            w.Commit();
        }
        return e;
    }

    [Benchmark(Description = "Dependents of one calc cell")]
    public int QueryCalcCell() => _engine.Dependents(SyntheticModel.Book, "Calc5", CellRect.Cell(120, 40)).Items.Count;

    [Benchmark(Description = "Precedents of one cell")]
    public int Precedents() => _engine.Precedents(SyntheticModel.Book, "Calc5", 120, 40, "=SUM(R[-10]C:R[-1]C)+Calc2!R7C*(1+Input3)").Items.Count;
}
