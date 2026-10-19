namespace Tracer.Core.Model;

public enum AreaKind : byte
{
    /// <summary><c>R1C1</c> or <c>R1C1:R2C2</c>.</summary>
    Cells,
    /// <summary>Whole rows, <c>R1</c> or <c>R1:R3</c>.</summary>
    Rows,
    /// <summary>Whole columns, <c>C2</c> or <c>C2:C4</c>.</summary>
    Cols,
}

/// <summary>
/// A reference as written in an R1C1 formula: four bounds, each absolute or relative to the formula cell.
/// Whole rows/columns are normalized to full-sheet bounds on the other axis.
/// </summary>
public readonly struct Area
{
    public readonly Bound Row1, Col1, Row2, Col2;
    public readonly AreaKind Kind;

    public Area(Bound row1, Bound col1, Bound row2, Bound col2, AreaKind kind = AreaKind.Cells)
    {
        Row1 = row1; Col1 = col1; Row2 = row2; Col2 = col2; Kind = kind;
    }

    public static Area Cell(Bound row, Bound col) => new(row, col, row, col);

    public static Area Rows(Bound row1, Bound row2) =>
        new(row1, Bound.Abs(1), row2, Bound.Abs(Limits.MaxCols), AreaKind.Rows);

    public static Area Cols(Bound col1, Bound col2) =>
        new(Bound.Abs(1), col1, Bound.Abs(Limits.MaxRows), col2, AreaKind.Cols);

    public static Area Absolute(CellRect r) =>
        new(Bound.Abs(r.Row1), Bound.Abs(r.Col1), Bound.Abs(r.Row2), Bound.Abs(r.Col2));

    public bool IsSingleCell => Kind == AreaKind.Cells && Row1.Equals(Row2) && Col1.Equals(Col2);
    public bool IsFullyAbsolute => !Row1.IsRelative && !Row2.IsRelative && !Col1.IsRelative && !Col2.IsRelative;

    /// <summary>The cells this reference points at when the formula sits at (row, col).</summary>
    public CellRect Resolve(int row, int col) => new(
        Row1.Resolve(row, Limits.MaxRows), Col1.Resolve(col, Limits.MaxCols),
        Row2.Resolve(row, Limits.MaxRows), Col2.Resolve(col, Limits.MaxCols));

    public override string ToString() => Text.A1.R1C1(this);
}
