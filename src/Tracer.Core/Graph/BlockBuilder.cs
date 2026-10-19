using System.Collections.Generic;
using Tracer.Core.Model;

namespace Tracer.Core.Graph;

/// <summary>A rectangle of cells that all hold the same R1C1 formula.</summary>
public readonly struct FormulaBlock
{
    public readonly CellRect Cells;
    public readonly int TemplateId;

    public FormulaBlock(CellRect cells, int templateId)
    {
        Cells = cells;
        TemplateId = templateId;
    }
}

/// <summary>
/// Groups formula cells, fed in row-major order, into rectangles of identical R1C1 formulas:
/// first horizontal runs (formulas copied across period columns), then runs stacked on consecutive rows
/// with the same columns (formulas copied down a list).
/// </summary>
public sealed class BlockBuilder
{
    private readonly List<(int R1, int C1, int R2, int C2, int Tid)> _blocks = new();
    private readonly Dictionary<long, int> _open = new();
    private int _runRow = -1, _runC1, _runC2, _runTid;

    public int CellCount { get; private set; }

    public void Add(int row, int col, int templateId)
    {
        CellCount++;
        if (row == _runRow && col == _runC2 + 1 && templateId == _runTid)
        {
            _runC2 = col;
            return;
        }
        Flush();
        _runRow = row;
        _runC1 = _runC2 = col;
        _runTid = templateId;
    }

    private void Flush()
    {
        if (_runRow < 0) return;
        long key = ((long)_runTid << 30) | ((long)_runC1 << 15) | (uint)_runC2;
        if (_open.TryGetValue(key, out int idx) && _blocks[idx].R2 == _runRow - 1)
        {
            var b = _blocks[idx];
            _blocks[idx] = (b.R1, b.C1, _runRow, b.C2, b.Tid);
        }
        else
        {
            _open[key] = _blocks.Count;
            _blocks.Add((_runRow, _runC1, _runRow, _runC2, _runTid));
        }
        _runRow = -1;
    }

    public List<FormulaBlock> Build()
    {
        Flush();
        var result = new List<FormulaBlock>(_blocks.Count);
        foreach (var b in _blocks) result.Add(new FormulaBlock(new CellRect(b.R1, b.C1, b.R2, b.C2), b.Tid));
        return result;
    }
}
