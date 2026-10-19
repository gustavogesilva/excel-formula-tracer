using System;
using System.Collections.Generic;
using Tracer.Core.Model;

namespace Tracer.Core.Graph;

/// <summary>
/// "Formulas in block B (on SourceSheet) reference Area, evaluated at each of their own cells."
/// One rule covers a whole block of copied formulas, so the index grows with distinct formulas, not cells.
/// </summary>
internal struct Rule
{
    public int SourceSheet;
    public CellRect Block;
    public Area Area;
    /// <summary>Bounding box of everything the rule can point at on the target sheet.</summary>
    public CellRect Target;
    public int TemplateId;
}

/// <summary>Formula cells (a sub-rectangle of one block) that reference the queried cells.</summary>
public readonly struct DependentHit
{
    public readonly int SourceSheet;
    public readonly CellRect Cells;
    public readonly CellRect Block;
    public readonly int TemplateId;

    public DependentHit(int sourceSheet, CellRect cells, CellRect block, int templateId)
    {
        SourceSheet = sourceSheet; Cells = cells; Block = block; TemplateId = templateId;
    }
}

/// <summary>Reverse index from target sheet to rules, answering "who references these cells?".</summary>
public sealed class DependencyIndex
{
    private readonly Dictionary<int, List<Rule>> _byTarget = new();
    private readonly Dictionary<int, HashSet<int>> _targetsOfSource = new();

    public int RuleCount { get; private set; }

    /// <summary>Adds a rule for one block and one (bound) reference, splitting per cell if the reference wraps the sheet edge.</summary>
    internal void Add(int sourceSheet, CellRect block, int templateId, int targetSheet, Area area)
    {
        if (!area.Row1.StaysInside(block.Row1, block.Row2, Limits.MaxRows) || !area.Row2.StaysInside(block.Row1, block.Row2, Limits.MaxRows) ||
            !area.Col1.StaysInside(block.Col1, block.Col2, Limits.MaxCols) || !area.Col2.StaysInside(block.Col1, block.Col2, Limits.MaxCols))
        {
            for (int r = block.Row1; r <= block.Row2; r++)
                for (int c = block.Col1; c <= block.Col2; c++)
                    AddRule(sourceSheet, CellRect.Cell(r, c), templateId, targetSheet, Area.Absolute(area.Resolve(r, c)));
            return;
        }
        AddRule(sourceSheet, block, templateId, targetSheet, area);
    }

    private void AddRule(int sourceSheet, CellRect block, int templateId, int targetSheet, Area area)
    {
        // Each bound is monotone in the anchor, so the extremes occur at the block's corners.
        var a = area.Resolve(block.Row1, block.Col1);
        var b = area.Resolve(block.Row2, block.Col2);
        var target = new CellRect(Math.Min(a.Row1, b.Row1), Math.Min(a.Col1, b.Col1), Math.Max(a.Row2, b.Row2), Math.Max(a.Col2, b.Col2));

        if (!_byTarget.TryGetValue(targetSheet, out var list)) _byTarget[targetSheet] = list = new List<Rule>();
        list.Add(new Rule { SourceSheet = sourceSheet, Block = block, Area = area, Target = target, TemplateId = templateId });
        if (!_targetsOfSource.TryGetValue(sourceSheet, out var targets)) _targetsOfSource[sourceSheet] = targets = new HashSet<int>();
        targets.Add(targetSheet);
        RuleCount++;
    }

    /// <summary>Drops every rule that came from formulas on <paramref name="sourceSheet"/>.</summary>
    public void RemoveSource(int sourceSheet)
    {
        if (!_targetsOfSource.TryGetValue(sourceSheet, out var targets)) return;
        foreach (int t in targets)
            if (_byTarget.TryGetValue(t, out var list))
                RuleCount -= list.RemoveAll(r => r.SourceSheet == sourceSheet);
        _targetsOfSource.Remove(sourceSheet);
    }

    /// <summary>All formula cells referencing any cell of <paramref name="target"/> on <paramref name="targetSheet"/>.</summary>
    public List<DependentHit> Query(int targetSheet, CellRect target)
    {
        var hits = new List<DependentHit>();
        if (!_byTarget.TryGetValue(targetSheet, out var list)) return hits;

        foreach (var rule in list)
        {
            if (!rule.Target.Intersects(target)) continue;
            if (!Solve(rule.Area.Row1, rule.Area.Row2, target.Row1, target.Row2, rule.Block.Row1, rule.Block.Row2, out int r1, out int r2)) continue;
            if (!Solve(rule.Area.Col1, rule.Area.Col2, target.Col1, target.Col2, rule.Block.Col1, rule.Block.Col2, out int c1, out int c2)) continue;
            hits.Add(new DependentHit(rule.SourceSheet, new CellRect(r1, c1, r2, c2), rule.Block, rule.TemplateId));
        }
        return Dedupe(hits);
    }

    /// <summary>
    /// Anchors a in [lo, hi] for which the referenced span [min(f1(a), f2(a)), max(f1(a), f2(a))] overlaps [t1, t2],
    /// where each f is either a constant (absolute bound) or a + d (relative bound). Both min and max are
    /// non-decreasing in a, so the answer is one contiguous interval with a closed form.
    /// </summary>
    internal static bool Solve(Bound b1, Bound b2, int t1, int t2, int lo, int hi, out int x, out int y)
    {
        const int NegInf = int.MinValue / 2, PosInf = int.MaxValue / 2;

        // min(f1, f2) <= t2  <=>  f1 <= t2 || f2 <= t2  <=>  a <= max(up(b1), up(b2))
        static int Up(Bound b, int t2) => b.IsRelative ? t2 - b.Value : (b.Value <= t2 ? PosInf : NegInf);
        // max(f1, f2) >= t1  <=>  f1 >= t1 || f2 >= t1  <=>  a >= min(down(b1), down(b2))
        static int Down(Bound b, int t1) => b.IsRelative ? t1 - b.Value : (b.Value >= t1 ? NegInf : PosInf);

        int aMax = Math.Max(Up(b1, t2), Up(b2, t2));
        int aMin = Math.Min(Down(b1, t1), Down(b2, t1));
        x = Math.Max(lo, aMin);
        y = Math.Min(hi, aMax);
        return x <= y;
    }

    /// <summary>
    /// Several references in one block can hit the same cells; keep each block's hits minimal.
    /// Different blocks never share cells, so only hits from the same block are compared.
    /// </summary>
    private static List<DependentHit> Dedupe(List<DependentHit> hits)
    {
        if (hits.Count < 2) return hits;
        var byBlock = new Dictionary<(int, CellRect), List<DependentHit>>();
        var order = new List<(int, CellRect)>();
        foreach (var h in hits)
        {
            var key = (h.SourceSheet, h.Block);
            if (!byBlock.TryGetValue(key, out var group))
            {
                byBlock[key] = group = new List<DependentHit>(1);
                order.Add(key);
            }
            group.Add(h);
        }
        if (order.Count == hits.Count) return hits;

        var keep = new List<DependentHit>(hits.Count);
        foreach (var key in order)
        {
            var group = byBlock[key];
            for (int i = 0; i < group.Count; i++)
            {
                bool covered = false;
                for (int j = 0; j < group.Count && !covered; j++)
                {
                    if (i == j || !group[j].Cells.Contains(group[i].Cells)) continue;
                    // Drop i if another hit strictly contains it; of exact duplicates keep the first.
                    covered = !group[j].Cells.Equals(group[i].Cells) || j < i;
                }
                if (!covered) keep.Add(group[i]);
            }
        }
        return keep;
    }
}
