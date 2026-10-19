using System;
using System.Collections.Generic;
using System.Linq;
using Tracer.Core.Binding;
using Tracer.Core.Graph;
using Tracer.Core.Model;
using Tracer.Core.Parsing;
using Tracer.Core.Text;

namespace Tracer.Core.Tracing;

/// <summary>
/// Excel-independent tracing engine. The host (the Excel add-in, or tests) keeps workbook models current,
/// feeds sheet formulas when a sheet needs indexing, and asks for precedents/dependents.
/// Not thread-safe: use from one thread (Excel's main thread).
/// </summary>
public sealed partial class TraceEngine
{
    private sealed class SheetState
    {
        public string Book = string.Empty;
        public string Sheet = string.Empty;
        public bool Dirty;
        public int FormulaCells;
        public int Blocks;
        public long DynamicCells;
    }

    private readonly Dictionary<int, SheetState> _states = new();

    public TraceEngine()
    {
        Binder = new Binder(Workspace, Templates);
    }

    public WorkspaceModel Workspace { get; } = new();
    public TemplateCache Templates { get; } = new();
    public SheetRegistry Sheets { get; } = new();
    public DependencyIndex Index { get; } = new();
    public Binder Binder { get; }

    // ---- workbook / sheet lifecycle -----------------------------------------------------------

    /// <summary>Adds or refreshes a workbook's structure. If names, tables or sheet order changed, its sheets are re-indexed.</summary>
    public void UpdateBook(WorkbookModel model)
    {
        var old = Workspace.Book(model.Name);
        Workspace.Set(model);
        if (old == null) return;
        foreach (var s in old.Sheets.Where(s => !model.HasSheet(s)).ToList()) DropSheet(model.Name, s);
        if (old.Signature() != model.Signature()) MarkBookDirty(model.Name);
    }

    public void RemoveBook(string book)
    {
        foreach (var kv in _states.Where(kv => string.Equals(kv.Value.Book, book, StringComparison.OrdinalIgnoreCase)).ToList())
        {
            Index.RemoveSource(kv.Key);
            _states.Remove(kv.Key);
        }
        Workspace.Remove(book);
    }

    public void MarkSheetDirty(string book, string sheet)
    {
        if (Sheets.TryGetId(book, sheet, out int id) && _states.TryGetValue(id, out var st)) st.Dirty = true;
    }

    public void MarkBookDirty(string book)
    {
        foreach (var st in _states.Values)
            if (string.Equals(st.Book, book, StringComparison.OrdinalIgnoreCase)) st.Dirty = true;
    }

    public void MarkAllDirty()
    {
        foreach (var st in _states.Values) st.Dirty = true;
    }

    private void DropSheet(string book, string sheet)
    {
        if (!Sheets.TryGetId(book, sheet, out int id)) return;
        Index.RemoveSource(id);
        _states.Remove(id);
    }

    public bool IsIndexed(string book, string sheet) =>
        Sheets.TryGetId(book, sheet, out int id) && _states.TryGetValue(id, out var st) && !st.Dirty;

    /// <summary>Sheets of open workbooks that are dirty or have never been indexed.</summary>
    public List<(string Book, string Sheet)> SheetsNeedingIndex()
    {
        var list = new List<(string, string)>();
        foreach (var b in Workspace.Books)
            foreach (var s in b.Sheets)
                if (!IsIndexed(b.Name, s)) list.Add((b.Name, s));
        return list;
    }

    /// <summary>Starts (re)indexing a sheet. Feed every formula cell in row-major order, then <see cref="SheetWriter.Commit"/>.</summary>
    public SheetWriter BeginSheet(string book, string sheet) => new(this, book, sheet);

    public sealed class SheetWriter
    {
        private readonly TraceEngine _e;
        private readonly string _book, _sheet;
        private readonly BlockBuilder _blocks = new();

        internal SheetWriter(TraceEngine e, string book, string sheet)
        {
            _e = e;
            _book = book;
            _sheet = sheet;
        }

        /// <summary>Adds a formula cell. Must be called in row-major order (row ascending, then column ascending).</summary>
        public void Add(int row, int col, string formulaR1C1) => _blocks.Add(row, col, _e.Templates.Get(formulaR1C1).Id);

        public void Commit()
        {
            var e = _e;
            var model = e.Workspace.Book(_book) ?? throw new InvalidOperationException($"Workbook '{_book}' is not registered.");
            int sourceId = e.Sheets.Id(_book, _sheet);
            e.Index.RemoveSource(sourceId);

            var blocks = _blocks.Build();
            var st = new SheetState { Book = _book, Sheet = _sheet, FormulaCells = _blocks.CellCount, Blocks = blocks.Count };
            var byTemplate = new Dictionary<int, BoundRef[]>();
            var tmp = new List<BoundRef>();

            foreach (var block in blocks)
            {
                var t = e.Templates[block.TemplateId];
                if (t.HasDynamic) st.DynamicCells += block.Cells.Count;
                if (t.HasUnqualifiedStruct || !byTemplate.TryGetValue(block.TemplateId, out var refs))
                {
                    tmp.Clear();
                    e.Binder.BindAll(t, new BindContext(model, _sheet, block.Cells.Row1, block.Cells.Col1), tmp);
                    refs = tmp.ToArray();
                    if (!t.HasUnqualifiedStruct) byTemplate[block.TemplateId] = refs;
                }
                foreach (var r in refs)
                    e.Index.Add(sourceId, block.Cells, block.TemplateId, e.Sheets.Id(r.Book, r.Sheet), r.Area);
            }
            e._states[sourceId] = st;
        }
    }

    public (int FormulaCells, int Blocks, int Rules, int Templates) Stats()
    {
        int cells = 0, blocks = 0;
        foreach (var st in _states.Values) { cells += st.FormulaCells; blocks += st.Blocks; }
        return (cells, blocks, Index.RuleCount, Templates.Count);
    }

    // ---- precedents ---------------------------------------------------------------------

    /// <summary>Precedents of the formula at (row, col). Needs only the workbook model, not the index.</summary>
    public PrecedentResult Precedents(string book, string sheet, int row, int col, string formulaR1C1)
    {
        var model = Workspace.Book(book) ?? throw new InvalidOperationException($"Workbook '{book}' is not registered.");
        var t = Templates.Get(formulaR1C1);
        var ctx = new BindContext(model, sheet, row, col);
        var rendered = FormulaRenderer.ToA1(t, row, col);
        var items = new List<TraceItem>();
        var tmp = new List<BoundRef>();

        for (int i = 0; i < t.Refs.Length; i++)
        {
            var r = t.Refs[i];
            var (s, len) = rendered.MapSpan(r.Start, r.Length);
            if ((r.Flags & RefFlags.Broken) != 0)
            {
                items.Add(new TraceItem { Kind = TraceItemKind.Broken, Label = rendered.Text.Substring(s, len), SpanStart = s, SpanLength = len });
                continue;
            }
            tmp.Clear();
            Binder.BindRef(t, i, ctx, tmp, RefOrigin.Direct, i);
            foreach (var b in tmp)
            {
                var item = FromBound(b, row, col, model.Name, sheet);
                item.SpanStart = s;
                item.SpanLength = len;
                item.IsSpill = (r.Flags & RefFlags.Spill) != 0;
                item.IsAddressOnly = (r.Flags & RefFlags.InDynamicCall) != 0;
                items.Add(item);
            }
        }

        for (int i = 0; i < t.StructRefs.Length; i++)
        {
            var sr = t.StructRefs[i];
            var (s, len) = rendered.MapSpan(sr.Start, sr.Length);
            string text = t.Formula.Substring(sr.Start, sr.Length);
            if (Binder.BindStruct(t, i, ctx, out var b, i))
            {
                var item = FromBound(b, row, col, model.Name, sheet);
                item.Kind = TraceItemKind.Table;
                item.Label = text;
                item.Expression = A1.Qualified(b.Book, b.Sheet, item.Rect, model.Name, sheet);
                item.SpanStart = s;
                item.SpanLength = len;
                items.Add(item);
            }
            else
            {
                items.Add(new TraceItem { Kind = TraceItemKind.Unresolved, Label = text, SpanStart = s, SpanLength = len });
            }
        }

        for (int i = 0; i < t.Names.Length; i++)
        {
            var n = t.Names[i];
            var (s, len) = rendered.MapSpan(n.Start, n.Length);
            var (nameBook, def) = Binder.ResolveName(n.Name, t.PrefixOf(n.PrefixIndex), ctx);
            string text = t.Formula.Substring(n.Start, n.Length);
            items.Add(def == null || nameBook == null
                ? new TraceItem
                {
                    Kind = t.PrefixOf(n.PrefixIndex)?.Book != null ? TraceItemKind.External : TraceItemKind.Unresolved,
                    Label = text, SpanStart = s, SpanLength = len,
                }
                : NameItem(def, nameBook, ctx, text, s, len, i));
        }

        foreach (var c in t.Calls)
        {
            if (c.IsDynamic)
            {
                var (s, len) = rendered.MapSpan(c.Start, c.Length);
                string expr = rendered.Text.Substring(s, len);
                items.Add(new TraceItem { Kind = TraceItemKind.Dynamic, Label = expr, Expression = expr, SpanStart = s, SpanLength = len });
                continue;
            }
            var def = model.FindName(c.Name, sheet);
            if (def == null) continue;
            int start = rendered.Map(c.Start);
            items.Add(NameItem(def, model, ctx, def.Name, start, def.Name.Length, -1));
        }

        items.Sort((a, b) => a.SpanStart.CompareTo(b.SpanStart));
        return new PrecedentResult(t, rendered, items);
    }

    private TraceItem NameItem(NameDef def, WorkbookModel nameBook, in BindContext ctx, string label, int spanStart, int spanLength, int token)
    {
        var refs = new List<BoundRef>();
        Binder.BindNameDefinition(def, nameBook, ctx, refs, token, 0);
        var children = new List<TraceItem>(refs.Count);
        foreach (var b in refs) children.Add(FromBound(b, ctx.Row, ctx.Col, ctx.Book.Name, ctx.Sheet));
        var defText = FormulaRenderer.ToA1(Templates.Get(def.RefersToR1C1), ctx.Row, ctx.Col).Text;
        var item = new TraceItem
        {
            Kind = TraceItemKind.Name,
            Label = label,
            Expression = defText,
            SpanStart = spanStart,
            SpanLength = spanLength,
            Children = children,
        };
        if (children.Count == 1 && children[0].HasTarget)
        {
            item.Book = children[0].Book;
            item.Sheet = children[0].Sheet;
            item.Rect = children[0].Rect;
            item.HasTarget = true;
        }
        return item;
    }

    private static TraceItem FromBound(BoundRef b, int row, int col, string contextBook, string contextSheet)
    {
        var rect = b.Area.Resolve(row, col);
        return new TraceItem
        {
            Kind = !b.IsOpen ? TraceItemKind.External : rect.IsSingleCell ? TraceItemKind.Cell : TraceItemKind.Range,
            Book = b.Book,
            Sheet = b.Sheet,
            Rect = rect,
            HasTarget = b.IsOpen,
            Label = A1.Qualified(b.Book, b.Sheet, rect, contextBook, contextSheet),
        };
    }

    // ---- dependents -----------------------------------------------------------------------

    /// <summary>
    /// Formula cells that reference any cell in <paramref name="target"/>, grouped into blocks of copied formulas.
    /// Only indexed sheets are searched; call <see cref="SheetsNeedingIndex"/> and index them first.
    /// </summary>
    public DependentResult Dependents(string book, string sheet, CellRect target)
    {
        var items = new List<TraceItem>();
        long dynamicCells = 0;
        int unindexed = 0;
        foreach (var st in _states.Values)
        {
            dynamicCells += st.DynamicCells;
            if (st.Dirty) unindexed++;
        }
        foreach (var b in Workspace.Books)
            foreach (var s in b.Sheets)
                if (!Sheets.TryGetId(b.Name, s, out int sid) || !_states.ContainsKey(sid)) unindexed++;

        if (!Sheets.TryGetId(book, sheet, out int targetId))
            return new DependentResult(items, dynamicCells, unindexed);

        foreach (var hit in Index.Query(targetId, target))
        {
            var (hb, hs) = Sheets[hit.SourceSheet];
            var t = Templates[hit.TemplateId];
            items.Add(new TraceItem
            {
                Kind = hit.Cells.IsSingleCell ? TraceItemKind.Cell : TraceItemKind.Range,
                Book = hb,
                Sheet = hs,
                Rect = hit.Cells,
                HasTarget = true,
                Label = A1.Qualified(hb, hs, hit.Cells, book, sheet),
                Expression = FormulaRenderer.ToA1(t, hit.Cells.Row1, hit.Cells.Col1).Text,
            });
        }

        items.Sort((a, b) =>
        {
            int c = SameRank(a.Book, book).CompareTo(SameRank(b.Book, book));
            if (c == 0) c = string.Compare(a.Book, b.Book, StringComparison.OrdinalIgnoreCase);
            if (c == 0) c = SheetOrder(a.Book, a.Sheet).CompareTo(SheetOrder(b.Book, b.Sheet));
            if (c == 0) c = a.Rect.Row1.CompareTo(b.Rect.Row1);
            if (c == 0) c = a.Rect.Col1.CompareTo(b.Rect.Col1);
            return c;
        });
        return new DependentResult(items, dynamicCells, unindexed);
    }

    private static int SameRank(string a, string b) => string.Equals(a, b, StringComparison.OrdinalIgnoreCase) ? 0 : 1;

    private int SheetOrder(string book, string sheet)
    {
        var m = Workspace.Book(book);
        if (m == null) return int.MaxValue;
        for (int i = 0; i < m.Sheets.Count; i++)
            if (string.Equals(m.Sheets[i], sheet, StringComparison.OrdinalIgnoreCase)) return i;
        return int.MaxValue;
    }
}
