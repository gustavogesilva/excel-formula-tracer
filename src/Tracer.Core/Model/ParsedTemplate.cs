using System;

namespace Tracer.Core.Model;

/// <summary>
/// The sheet/workbook qualifier in front of a reference or name, e.g. <c>'My Sheet'!</c>,
/// <c>Sheet1:Sheet3!</c>, <c>'C:\dir\[Book.xlsx]Data'!</c> or <c>Book.xlsx!</c>.
/// </summary>
public sealed class Prefix
{
    /// <summary>Workbook file name from <c>[Book.xlsx]</c>, or null for the formula's own workbook.</summary>
    public string? Book { get; }
    /// <summary>Directory before <c>[Book.xlsx]</c> for closed external workbooks, else null.</summary>
    public string? Path { get; }
    /// <summary>First (or only) sheet. For <c>Book.xlsx!Name</c> this holds the token before the bang, which may be a workbook name.</summary>
    public string? Sheet1 { get; }
    /// <summary>Last sheet of a 3D reference, else null.</summary>
    public string? Sheet2 { get; }

    public Prefix(string? book, string? path, string? sheet1, string? sheet2)
    {
        Book = book; Path = path; Sheet1 = sheet1; Sheet2 = sheet2;
    }

    public bool Is3D => Sheet2 != null;
    public override string ToString() => $"{(Path != null ? Path : "")}{(Book != null ? "[" + Book + "]" : "")}{Sheet1}{(Sheet2 != null ? ":" + Sheet2 : "")}!";
}

[Flags]
public enum RefFlags : byte
{
    None = 0,
    /// <summary><c>R1C1#</c> spill reference.</summary>
    Spill = 1,
    /// <summary>Argument of OFFSET/INDIRECT: the address is used, the cell may not be.</summary>
    InDynamicCall = 2,
    /// <summary><c>Sheet1!#REF!</c> or similar: points at nothing.</summary>
    Broken = 4,
}

/// <summary>A cell/range reference found in a formula.</summary>
public readonly struct RefToken
{
    public readonly Area Area;
    /// <summary>Index into <see cref="ParsedTemplate.Prefixes"/>, or -1 for the formula's own sheet.</summary>
    public readonly int PrefixIndex;
    /// <summary>Span of the whole token in the R1C1 text, prefix included.</summary>
    public readonly int Start, Length;
    /// <summary>Where the address part begins (after the prefix).</summary>
    public readonly int AreaStart;
    public readonly RefFlags Flags;

    public RefToken(Area area, int prefixIndex, int start, int length, int areaStart, RefFlags flags)
    {
        Area = area; PrefixIndex = prefixIndex; Start = start; Length = length; AreaStart = areaStart; Flags = flags;
    }

    public int End => Start + Length;
}

/// <summary>A defined name (or a call to a LAMBDA stored in a name) used in a formula.</summary>
public readonly struct NameToken
{
    public readonly string Name;
    public readonly int PrefixIndex;
    public readonly int Start, Length;
    /// <summary>The name is called like a function: <c>MyLambda(1, 2)</c>.</summary>
    public readonly bool IsCall;

    public NameToken(string name, int prefixIndex, int start, int length, bool isCall)
    {
        Name = name; PrefixIndex = prefixIndex; Start = start; Length = length; IsCall = isCall;
    }

    public int End => Start + Length;
}

[Flags]
public enum TableArea : byte
{
    None = 0,
    Headers = 1,
    Data = 2,
    Totals = 4,
    All = Headers | Data | Totals,
    ThisRow = 8,
}

/// <summary>A structured reference such as <c>Table1[Price]</c> or <c>[@Qty]</c>.</summary>
public readonly struct StructToken
{
    /// <summary>Table name, or null when unqualified (the table containing the formula cell).</summary>
    public readonly string? Table;
    public readonly TableArea Area;
    /// <summary>Column range, or null for all columns.</summary>
    public readonly string? FirstColumn, LastColumn;
    public readonly int PrefixIndex;
    public readonly int Start, Length;

    public StructToken(string? table, TableArea area, string? firstColumn, string? lastColumn, int prefixIndex, int start, int length)
    {
        Table = table; Area = area; FirstColumn = firstColumn; LastColumn = lastColumn;
        PrefixIndex = prefixIndex; Start = start; Length = length;
    }

    public int End => Start + Length;
}

/// <summary>A function call; kept for OFFSET/INDIRECT evaluation and LAMBDA-name lookup.</summary>
public readonly struct CallToken
{
    /// <summary>Upper-case name without <c>_xlfn.</c>/<c>_xlws.</c> prefixes; empty for an inline LAMBDA invocation.</summary>
    public readonly string Name;
    /// <summary>Span from the function name through the closing parenthesis.</summary>
    public readonly int Start, Length;

    public CallToken(string name, int start, int length)
    {
        Name = name; Start = start; Length = length;
    }

    public int End => Start + Length;
    public bool IsDynamic => Name == "OFFSET" || Name == "INDIRECT";
}

/// <summary>
/// Everything the tracer needs from one distinct R1C1 formula. Copies of a formula share the same R1C1 text,
/// so one instance serves every cell holding it.
/// </summary>
public sealed class ParsedTemplate
{
    public int Id { get; internal set; } = -1;
    public string Formula { get; }
    public RefToken[] Refs { get; }
    public NameToken[] Names { get; }
    public StructToken[] StructRefs { get; }
    public CallToken[] Calls { get; }
    public Prefix[] Prefixes { get; }
    /// <summary>Spans of numeric literals written into the formula (hard-coded inputs).</summary>
    public int[] NumberStarts { get; }
    public bool HasDynamic { get; }
    public bool HasUnqualifiedStruct { get; }

    public ParsedTemplate(string formula, RefToken[] refs, NameToken[] names, StructToken[] structRefs,
        CallToken[] calls, Prefix[] prefixes, int[] numberStarts)
    {
        Formula = formula; Refs = refs; Names = names; StructRefs = structRefs; Calls = calls;
        Prefixes = prefixes; NumberStarts = numberStarts;
        foreach (var c in calls) if (c.IsDynamic) { HasDynamic = true; break; }
        foreach (var s in structRefs) if (s.Table == null) { HasUnqualifiedStruct = true; break; }
    }

    public Prefix? PrefixOf(int prefixIndex) => prefixIndex < 0 ? null : Prefixes[prefixIndex];

    public static readonly ParsedTemplate Empty = new(string.Empty, Array.Empty<RefToken>(), Array.Empty<NameToken>(),
        Array.Empty<StructToken>(), Array.Empty<CallToken>(), Array.Empty<Prefix>(), Array.Empty<int>());
}
