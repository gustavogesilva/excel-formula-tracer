namespace Tracer.Core.Model;

/// <summary>Sheet dimensions (Excel 2007+).</summary>
public static class Limits
{
    public const int MaxRows = 1_048_576;
    public const int MaxCols = 16_384;
}

/// <summary>
/// One edge of a reference along one axis: either an absolute 1-based index (<c>R5</c>)
/// or an offset from the formula's own cell (<c>R[-1]</c>, or <c>R</c> for offset 0).
/// </summary>
public readonly struct Bound : System.IEquatable<Bound>
{
    public readonly int Value;
    public readonly bool IsRelative;

    public Bound(int value, bool isRelative)
    {
        Value = value;
        IsRelative = isRelative;
    }

    public static Bound Abs(int index) => new(index, false);
    public static Bound Rel(int offset) => new(offset, true);

    /// <summary>Resolves against the formula cell's index on this axis, wrapping around the sheet edge like Excel does.</summary>
    public int Resolve(int anchor, int max)
    {
        if (!IsRelative) return Value;
        int v = anchor + Value;
        if (v < 1 || v > max) v = ((v - 1) % max + max) % max + 1;
        return v;
    }

    /// <summary>True when resolving at any anchor in [lo, hi] stays on the sheet without wrapping.</summary>
    public bool StaysInside(int lo, int hi, int max) => !IsRelative || (lo + Value >= 1 && hi + Value <= max);

    public bool Equals(Bound other) => Value == other.Value && IsRelative == other.IsRelative;
    public override bool Equals(object? obj) => obj is Bound b && Equals(b);
    public override int GetHashCode() => IsRelative ? ~Value : Value;
    public override string ToString() => IsRelative ? $"[{Value}]" : Value.ToString();
}
