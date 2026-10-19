using System;
using System.Collections.Generic;
using Tracer.Core.Binding;

namespace Tracer.Bench;

/// <summary>
/// A financial-model-shaped workbook: sheets of line items (rows) projected across monthly periods (columns),
/// with formulas copied across each row, row-to-row links, subtotals, cross-sheet links, cumulative sums,
/// named inputs and the odd whole-column lookup. Deterministic for a given seed.
/// </summary>
public sealed class SyntheticModel
{
    public const string Book = "Model.xlsx";
    public readonly List<string> Sheets = new();
    public readonly List<NameDef> Names = new();
    /// <summary>Per sheet, formula cells in row-major order.</summary>
    public readonly Dictionary<string, List<(int Row, int Col, string Formula)>> Cells = new();

    public SyntheticModel(int sheets = 20, int rows = 300, int periods = 84, int seed = 42)
    {
        var rnd = new Random(seed);
        Sheets.Add("Inputs");
        for (int s = 1; s <= sheets; s++) Sheets.Add($"Calc{s}");
        for (int k = 1; k <= 50; k++) Names.Add(new NameDef($"Input{k}", null, $"=Inputs!R{k + 1}C3"));

        var inputs = new List<(int, int, string)>();
        Cells["Inputs"] = inputs; // constants only

        const int firstCol = 4; // A..C hold labels/units/opening values
        for (int s = 1; s <= sheets; s++)
        {
            var list = new List<(int, int, string)>();
            for (int r = 2; r <= rows; r++)
            {
                int kind = rnd.Next(100);
                string body;
                if (kind < 15) continue; // input/label row: constants
                else if (kind < 40) body = $"RC[-1]*(1+Input{rnd.Next(1, 51)})";
                else if (kind < 55) body = $"R[-{rnd.Next(1, 6)}]C*R[-{rnd.Next(1, 6)}]C";
                else if (kind < 65) body = $"SUM(R[-{rnd.Next(5, 15)}]C:R[-1]C)";
                else if (kind < 75 && s > 1) body = $"Calc{rnd.Next(1, s)}!R{rnd.Next(2, rows + 1)}C";
                else if (kind < 82) body = $"SUM(R{r - 1}C{firstCol}:R[-1]C)";
                else if (kind < 88) body = $"IF(R1C>Input{rnd.Next(1, 51)},R[-1]C,0)";
                else if (kind < 91) body = $"SUMIFS(Inputs!C5,Inputs!C2,R1C)";
                else if (kind < 94) body = $"OFFSET(RC[-1],0,-{rnd.Next(1, 3)})";
                else body = $"R[-1]C+Calc{rnd.Next(1, sheets + 1)}!R{rnd.Next(2, rows + 1)}C*1.05";

                // First period usually differs (opening balance), the rest are copies.
                list.Add((r, firstCol, $"=R{r}C3"));
                for (int c = firstCol + 1; c < firstCol + periods; c++) list.Add((r, c, "=" + body));
            }
            Cells[$"Calc{s}"] = list;
        }
    }

    public WorkbookModel Model() => new(Book, Sheets, Names);

    public int FormulaCells
    {
        get
        {
            int n = 0;
            foreach (var l in Cells.Values) n += l.Count;
            return n;
        }
    }
}
