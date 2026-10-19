using System;

namespace Tracer.Core.Model;

/// <summary>A concrete, normalized rectangle of cells (1-based, inclusive).</summary>
public readonly struct CellRect : IEquatable<CellRect>
{
    public readonly int Row1, Col1, Row2, Col2;

    public CellRect(int row1, int col1, int row2, int col2)
    {
        if (row1 <= row2) { Row1 = row1; Row2 = row2; } else { Row1 = row2; Row2 = row1; }
        if (col1 <= col2) { Col1 = col1; Col2 = col2; } else { Col1 = col2; Col2 = col1; }
    }

    public static CellRect Cell(int row, int col) => new(row, col, row, col);

    public int Rows => Row2 - Row1 + 1;
    public int Cols => Col2 - Col1 + 1;
    public long Count => (long)Rows * Cols;
    public bool IsSingleCell => Row1 == Row2 && Col1 == Col2;
    public bool IsWholeColumns => Row1 == 1 && Row2 == Limits.MaxRows;
    public bool IsWholeRows => Col1 == 1 && Col2 == Limits.MaxCols;

    public bool Contains(int row, int col) => row >= Row1 && row <= Row2 && col >= Col1 && col <= Col2;
    public bool Contains(CellRect o) => o.Row1 >= Row1 && o.Row2 <= Row2 && o.Col1 >= Col1 && o.Col2 <= Col2;
    public bool Intersects(CellRect o) => o.Row1 <= Row2 && o.Row2 >= Row1 && o.Col1 <= Col2 && o.Col2 >= Col1;

    public CellRect Intersect(CellRect o) =>
        new(Math.Max(Row1, o.Row1), Math.Max(Col1, o.Col1), Math.Min(Row2, o.Row2), Math.Min(Col2, o.Col2));

    public bool Equals(CellRect o) => Row1 == o.Row1 && Col1 == o.Col1 && Row2 == o.Row2 && Col2 == o.Col2;
    public override bool Equals(object? obj) => obj is CellRect r && Equals(r);
    public override int GetHashCode() => unchecked(((Row1 * 397 ^ Col1) * 397 ^ Row2) * 397 ^ Col2);
    public override string ToString() => Text.A1.Address(this, absolute: false);
}
