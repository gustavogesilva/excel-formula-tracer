using System.Collections.Generic;
using Tracer.Core.Model;
using Tracer.Core.Parsing;

namespace Tracer.Core.Binding;

public enum RefOrigin : byte
{
    /// <summary>A reference written in the formula.</summary>
    Direct,
    /// <summary>Reached through a defined name (including LAMBDA names called like functions).</summary>
    Name,
    /// <summary>A structured table reference.</summary>
    Table,
}

/// <summary>A reference resolved to a target workbook and sheet. Bounds may still be relative to the formula cell.</summary>
public readonly struct BoundRef
{
    public readonly string Book;
    public readonly string Sheet;
    public readonly Area Area;
    /// <summary>The target workbook is open and the sheet exists, so the target can be read and navigated.</summary>
    public readonly bool IsOpen;
    public readonly RefOrigin Origin;
    /// <summary>Index of the originating token in the top-level template (<see cref="ParsedTemplate.Refs"/>, <see cref="ParsedTemplate.Names"/> or <see cref="ParsedTemplate.StructRefs"/> per origin; -1 for name calls).</summary>
    public readonly int Token;
    public readonly RefFlags Flags;

    public BoundRef(string book, string sheet, Area area, bool isOpen, RefOrigin origin, int token, RefFlags flags)
    {
        Book = book; Sheet = sheet; Area = area; IsOpen = isOpen; Origin = origin; Token = token; Flags = flags;
    }
}

/// <summary>Where a formula lives: its workbook, sheet and cell (the anchor for relative bounds and <c>[@Col]</c>).</summary>
public readonly struct BindContext
{
    public readonly WorkbookModel Book;
    public readonly string Sheet;
    public readonly int Row, Col;

    public BindContext(WorkbookModel book, string sheet, int row, int col)
    {
        Book = book; Sheet = sheet; Row = row; Col = col;
    }
}

/// <summary>
/// Resolves a parsed formula's references against the open workbooks: sheet prefixes, external books,
/// 3D spans, defined names (recursively), LAMBDA-name calls and structured references.
/// </summary>
public sealed class Binder
{
    private const int MaxNameDepth = 8;
    private readonly WorkspaceModel _ws;
    private readonly TemplateCache _cache;

    public Binder(WorkspaceModel workspace, TemplateCache cache)
    {
        _ws = workspace;
        _cache = cache;
    }

    public TemplateCache Templates => _cache;

    /// <summary>Every reference the formula depends on, with names and tables expanded. Used by the dependents index.</summary>
    public void BindAll(ParsedTemplate t, in BindContext ctx, List<BoundRef> sink)
    {
        for (int i = 0; i < t.Refs.Length; i++) BindRef(t, i, ctx, sink, RefOrigin.Direct, i);
        for (int i = 0; i < t.StructRefs.Length; i++)
            if (BindStruct(t, i, ctx, out var r, i)) sink.Add(r);
        for (int i = 0; i < t.Names.Length; i++)
        {
            var (book, def) = ResolveName(t.Names[i].Name, t.PrefixOf(t.Names[i].PrefixIndex), ctx);
            if (book != null && def != null) BindNameDefinition(def, book, ctx, sink, i, 0);
        }
        foreach (var c in t.Calls)
        {
            if (c.IsDynamic) continue;
            var def = ctx.Book.FindName(c.Name, ctx.Sheet);
            if (def != null) BindNameDefinition(def, ctx.Book, ctx, sink, -1, 0);
        }
    }

    /// <summary>Binds reference token <paramref name="index"/>; a 3D reference yields one entry per sheet.</summary>
    public void BindRef(ParsedTemplate t, int index, in BindContext ctx, List<BoundRef> sink, RefOrigin origin, int token)
    {
        var r = t.Refs[index];
        if ((r.Flags & RefFlags.Broken) != 0) return;
        var p = t.PrefixOf(r.PrefixIndex);

        if (p == null)
        {
            sink.Add(new BoundRef(ctx.Book.Name, ctx.Sheet, r.Area, true, origin, token, r.Flags));
            return;
        }
        if (p.Sheet1 == null) return;

        var book = p.Book != null ? _ws.Book(p.Book) : ctx.Book;
        if (book == null)
        {
            // Closed external workbook: keep it so it can be shown (and indexed by name for when it opens).
            sink.Add(new BoundRef(p.Book!, p.Sheet1, r.Area, false, origin, token, r.Flags));
            return;
        }
        if (p.Is3D)
        {
            foreach (var s in book.SheetSpan(p.Sheet1, p.Sheet2!))
                sink.Add(new BoundRef(book.Name, s, r.Area, true, origin, token, r.Flags));
            return;
        }
        var sheet = book.Sheet(p.Sheet1);
        sink.Add(new BoundRef(book.Name, sheet ?? p.Sheet1, r.Area, sheet != null, origin, token, r.Flags));
    }

    /// <summary>Finds the definition of a name used with an optional prefix (<c>Sheet1!Rate</c>, <c>Book2.xlsx!Rate</c>).</summary>
    public (WorkbookModel? Book, NameDef? Def) ResolveName(string name, Prefix? p, in BindContext ctx)
    {
        if (p == null) return (ctx.Book, ctx.Book.FindName(name, ctx.Sheet));
        if (p.Book != null)
        {
            var b = _ws.Book(p.Book);
            if (b == null) return (null, null);
            return (b, p.Sheet1 != null ? b.FindLocalName(name, p.Sheet1) : b.FindName(name, null));
        }
        if (p.Sheet1 == null) return (null, null);
        if (ctx.Book.HasSheet(p.Sheet1))
            return (ctx.Book, ctx.Book.FindLocalName(name, p.Sheet1) ?? ctx.Book.FindName(name, null));
        var other = _ws.Book(p.Sheet1);
        return (other, other?.FindName(name, null));
    }

    /// <summary>Binds what a name refers to. Unqualified references inside a name float to the sheet using it.</summary>
    public void BindNameDefinition(NameDef def, WorkbookModel book, in BindContext ctx, List<BoundRef> sink, int token, int depth)
    {
        if (depth >= MaxNameDepth) return;
        var t = _cache.Get(def.RefersToR1C1);
        var inner = new BindContext(book, ReferenceEquals(book, ctx.Book) ? ctx.Sheet : def.ScopeSheet ?? ctx.Sheet, ctx.Row, ctx.Col);

        var tmp = new List<BoundRef>();
        for (int i = 0; i < t.Refs.Length; i++) BindRef(t, i, inner, tmp, RefOrigin.Name, token);
        for (int i = 0; i < t.StructRefs.Length; i++)
            if (BindStruct(t, i, inner, out var r, token)) tmp.Add(r);
        foreach (var r in tmp)
            sink.Add(new BoundRef(r.Book, r.Sheet, r.Area, r.IsOpen, RefOrigin.Name, token, r.Flags));

        foreach (var n in t.Names)
        {
            var (b, d) = ResolveName(n.Name, t.PrefixOf(n.PrefixIndex), inner);
            if (b != null && d != null && !ReferenceEquals(d, def)) BindNameDefinition(d, b, inner, sink, token, depth + 1);
        }
    }

    /// <summary>Resolves a structured reference to an area. <c>[@Col]</c> becomes a row-relative bound.</summary>
    public bool BindStruct(ParsedTemplate t, int index, in BindContext ctx, out BoundRef result, int token)
    {
        result = default;
        var s = t.StructRefs[index];
        var p = t.PrefixOf(s.PrefixIndex);
        var book = ctx.Book;
        if (p != null)
            book = (p.Book != null ? _ws.Book(p.Book) : p.Sheet1 != null ? _ws.Book(p.Sheet1) ?? ctx.Book : ctx.Book)!;
        if (book == null) return false;

        var table = s.Table != null ? book.FindTable(s.Table) : book.FindTableAt(ctx.Sheet, ctx.Row, ctx.Col);
        if (table == null) return false;

        int c1 = table.FirstCol, c2 = table.LastCol;
        if (s.FirstColumn != null)
        {
            c1 = table.ColumnIndex(s.FirstColumn);
            c2 = table.ColumnIndex(s.LastColumn ?? s.FirstColumn);
            if (c1 < 0 || c2 < 0) return false;
        }

        Bound r1, r2;
        if ((s.Area & TableArea.ThisRow) != 0)
        {
            r1 = r2 = Bound.Rel(0);
        }
        else
        {
            int lo = int.MaxValue, hi = int.MinValue;
            if ((s.Area & TableArea.Headers) != 0 && table.HeaderRow > 0) { lo = System.Math.Min(lo, table.HeaderRow); hi = System.Math.Max(hi, table.HeaderRow); }
            if ((s.Area & TableArea.Data) != 0 && table.LastDataRow >= table.FirstDataRow) { lo = System.Math.Min(lo, table.FirstDataRow); hi = System.Math.Max(hi, table.LastDataRow); }
            if ((s.Area & TableArea.Totals) != 0 && table.TotalsRow > 0) { lo = System.Math.Min(lo, table.TotalsRow); hi = System.Math.Max(hi, table.TotalsRow); }
            if (lo > hi) return false;
            r1 = Bound.Abs(lo);
            r2 = Bound.Abs(hi);
        }

        result = new BoundRef(book.Name, table.Sheet, new Area(r1, Bound.Abs(c1), r2, Bound.Abs(c2)), true, RefOrigin.Table, token, RefFlags.None);
        return true;
    }
}
