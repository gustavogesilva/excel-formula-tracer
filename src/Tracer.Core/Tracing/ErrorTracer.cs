using System.Collections.Generic;
using Tracer.Core.Model;
using Tracer.Core.Text;

namespace Tracer.Core.Tracing;

/// <summary>A cell's formula and error value, as the host reads them.</summary>
public readonly struct CellRead
{
    public CellRead(string? formulaR1C1, string? error)
    {
        FormulaR1C1 = formulaR1C1;
        Error = error;
    }

    /// <summary>The R1C1 formula, or null for a constant/empty cell.</summary>
    public string? FormulaR1C1 { get; }
    /// <summary>Error text such as <c>#REF!</c>, or null when the value is not an error.</summary>
    public string? Error { get; }
}

/// <summary>What an OFFSET/INDIRECT expression evaluates to: a range, or an error.</summary>
public readonly struct DynamicTarget
{
    public DynamicTarget(string book, string sheet, CellRect rect)
    {
        Book = book; Sheet = sheet; Rect = rect; Error = null;
    }

    public DynamicTarget(string? error)
    {
        Book = Sheet = string.Empty; Rect = default; Error = error ?? "#VALUE!";
    }

    public string Book { get; }
    public string Sheet { get; }
    public CellRect Rect { get; }
    public string? Error { get; }
    public bool IsRange => Error == null;
}

/// <summary>Cell access the error tracer needs; the add-in implements it over COM, tests over a grid.</summary>
public interface ICellReader
{
    CellRead Read(string book, string sheet, int row, int col);

    /// <summary>Cells of <paramref name="rect"/> whose value is exactly <paramref name="error"/> (one bulk read), at most <paramref name="max"/>.</summary>
    IEnumerable<(int Row, int Col)> ErrorCells(string book, string sheet, CellRect rect, string error, int max);

    /// <summary>Evaluates an A1 reference expression (OFFSET/INDIRECT) on a sheet.</summary>
    DynamicTarget Evaluate(string book, string sheet, string expressionA1);
}

/// <summary>Root cells of an error value.</summary>
public sealed class ErrorTraceResult
{
    public ErrorTraceResult(string? error, List<TraceItem> roots, int visited, bool truncated)
    {
        Error = error;
        Roots = roots;
        Visited = visited;
        Truncated = truncated;
    }

    /// <summary>The start cell's error, or null when it has no error value.</summary>
    public string? Error { get; }
    public List<TraceItem> Roots { get; }
    /// <summary>Cells carrying the error that were examined.</summary>
    public int Visited { get; }
    /// <summary>The search stopped at its cell budget; there may be more roots.</summary>
    public bool Truncated { get; }
}

public sealed partial class TraceEngine
{
    private readonly struct Pending
    {
        public Pending(string book, string sheet, int row, int col, int depth, CellRead read)
        {
            Book = book; Sheet = sheet; Row = row; Col = col; Depth = depth; Read = read;
        }

        public string Book { get; }
        public string Sheet { get; }
        public int Row { get; }
        public int Col { get; }
        public int Depth { get; }
        public CellRead Read { get; }
    }

    /// <summary>
    /// Walks back from a cell showing an error to the cells where that error starts. Errors propagate unchanged,
    /// so only precedents holding the same error are followed. A cell is a root when its formula has a broken
    /// reference (<c>#REF!</c>), it holds the error as a constant, an OFFSET/INDIRECT in it evaluates to the error,
    /// it uses an unknown name (<c>#NAME?</c>), or no precedent carries the error (its own function produced it).
    /// A cell can be a root and still lead to further roots, so the search continues past roots.
    /// </summary>
    public ErrorTraceResult TraceErrorRoots(string book, string sheet, int row, int col, ICellReader reader, int maxCells = 5000)
    {
        var first = reader.Read(book, sheet, row, col);
        var roots = new List<TraceItem>();
        if (first.Error == null) return new ErrorTraceResult(null, roots, 0, false);

        string error = first.Error;
        var visited = new HashSet<(int, int, int)>();
        var queue = new Queue<Pending>();
        bool truncated = false;

        bool Enqueue(string b, string s, int r, int c, int depth, CellRead read)
        {
            var key = (Sheets.Id(b, s), r, c);
            if (visited.Contains(key)) return true;
            if (visited.Count >= maxCells)
            {
                truncated = true;
                return false;
            }
            visited.Add(key);
            queue.Enqueue(new Pending(b, s, r, c, depth, read));
            return true;
        }

        Enqueue(book, sheet, row, col, 0, first);
        while (queue.Count > 0)
        {
            var node = queue.Dequeue();
            var reasons = new List<string>();
            string? formulaA1 = null;

            if (node.Read.FormulaR1C1 == null || Workspace.Book(node.Book) == null)
            {
                if (node.Read.FormulaR1C1 == null) reasons.Add($"{error} is typed into the cell as a value");
                else reasons.Add($"its workbook {node.Book} is not loaded");
            }
            else
            {
                var p = Precedents(node.Book, node.Sheet, node.Row, node.Col, node.Read.FormulaR1C1);
                formulaA1 = p.Formula.Text;
                bool carried = false;
                var closed = new List<string>();
                int nextDepth = node.Depth + 1;

                bool Follow(string b, string s, CellRect rect)
                {
                    bool any = false;
                    if (rect.IsSingleCell)
                    {
                        var r = reader.Read(b, s, rect.Row1, rect.Col1);
                        if (r.Error == error)
                        {
                            any = true;
                            if (!Enqueue(b, s, rect.Row1, rect.Col1, nextDepth, r)) return true;
                        }
                        return any;
                    }
                    int budget = System.Math.Max(1, maxCells - visited.Count + 1);
                    foreach (var (r, c) in reader.ErrorCells(b, s, rect, error, budget))
                    {
                        any = true;
                        if (!Enqueue(b, s, r, c, nextDepth, reader.Read(b, s, r, c))) break;
                    }
                    return any;
                }

                foreach (var item in p.Items)
                {
                    if (truncated) break;
                    switch (item.Kind)
                    {
                        case TraceItemKind.Broken:
                            if (error == "#REF!") reasons.Add($"its formula contains {item.Label}: a referenced cell, row, column or sheet was deleted");
                            break;
                        case TraceItemKind.Unresolved:
                            if (error == "#NAME?") reasons.Add($"{item.Label} is not a defined name or table");
                            break;
                        case TraceItemKind.External:
                            closed.Add(item.Label);
                            break;
                        case TraceItemKind.Dynamic:
                            if (item.Expression == null) break;
                            var dt = reader.Evaluate(node.Book, node.Sheet, item.Expression);
                            if (!dt.IsRange)
                            {
                                if (dt.Error == error) reasons.Add($"{item.Label} evaluates to {dt.Error}");
                            }
                            else
                            {
                                carried |= Follow(dt.Book, dt.Sheet, dt.Rect);
                            }
                            break;
                        case TraceItemKind.Name when item.Children != null:
                            foreach (var child in item.Children)
                                if (child.HasTarget) carried |= Follow(child.Book, child.Sheet, child.Rect);
                            break;
                        default:
                            if (item.HasTarget && !item.IsAddressOnly) carried |= Follow(item.Book, item.Sheet, item.Rect);
                            break;
                    }
                }

                if (!carried && reasons.Count == 0)
                    reasons.Add(closed.Count > 0
                        ? $"no open precedent carries {error}; it reads closed workbook data ({string.Join(", ", closed)})"
                        : $"{error} is produced here: no precedent carries it");
            }

            if (reasons.Count == 0) continue;
            var rect = CellRect.Cell(node.Row, node.Col);
            roots.Add(new TraceItem
            {
                Kind = TraceItemKind.Cell,
                Book = node.Book,
                Sheet = node.Sheet,
                Rect = rect,
                HasTarget = true,
                Label = A1.Qualified(node.Book, node.Sheet, rect, book, sheet),
                Expression = formulaA1,
                Detail = string.Join("; ", reasons) + (node.Depth == 0 ? " (the cell itself)" : $" ({node.Depth} step{(node.Depth == 1 ? "" : "s")} back)"),
            });
        }

        return new ErrorTraceResult(error, roots, visited.Count, truncated);
    }
}
