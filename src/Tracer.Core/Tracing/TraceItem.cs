using System.Collections.Generic;
using Tracer.Core.Model;

namespace Tracer.Core.Tracing;

public enum TraceItemKind : byte
{
    /// <summary>A single cell.</summary>
    Cell,
    /// <summary>A multi-cell range (expandable into its cells).</summary>
    Range,
    /// <summary>A defined name; <see cref="TraceItem.Children"/> holds what it refers to.</summary>
    Name,
    /// <summary>A structured table reference.</summary>
    Table,
    /// <summary>An OFFSET/INDIRECT call; <see cref="TraceItem.Expression"/> is A1 text for the host to evaluate.</summary>
    Dynamic,
    /// <summary>A reference into a workbook that is not open.</summary>
    External,
    /// <summary><c>#REF!</c> in the formula.</summary>
    Broken,
    /// <summary>A name or table that could not be resolved.</summary>
    Unresolved,
}

/// <summary>One row of the trace window: a precedent of the traced cell, or a block of dependents.</summary>
public sealed class TraceItem
{
    public TraceItemKind Kind { get; set; }
    /// <summary>Target workbook (precedents) or formula workbook (dependents).</summary>
    public string Book { get; set; } = string.Empty;
    public string Sheet { get; set; } = string.Empty;
    public CellRect Rect { get; set; }
    /// <summary>False when there is nothing to navigate to (unresolved name, dynamic not yet evaluated).</summary>
    public bool HasTarget { get; set; }
    /// <summary>What the list shows in its first column, e.g. <c>Inputs!$B$2</c> or <c>Rate</c>.</summary>
    public string Label { get; set; } = string.Empty;
    /// <summary>Span in the A1 rendering of the traced formula, for highlighting; -1 for none.</summary>
    public int SpanStart { get; set; } = -1;
    public int SpanLength { get; set; }
    /// <summary>Dynamic: A1 expression to evaluate. Name: its definition in A1. Dependents: the formula in A1.</summary>
    public string? Expression { get; set; }
    public bool IsSpill { get; set; }
    /// <summary>The cell's address only feeds OFFSET's base argument (its value is not used).</summary>
    public bool IsAddressOnly { get; set; }
    public List<TraceItem>? Children { get; set; }

    public override string ToString() => $"{Kind} {Label}";
}

/// <summary>Precedents of one formula cell.</summary>
public sealed class PrecedentResult
{
    public PrecedentResult(ParsedTemplate template, Text.RenderedFormula formula, List<TraceItem> items)
    {
        Template = template;
        Formula = formula;
        Items = items;
    }

    public ParsedTemplate Template { get; }
    /// <summary>The formula in A1 notation; item spans point into this text.</summary>
    public Text.RenderedFormula Formula { get; }
    public List<TraceItem> Items { get; }
}

/// <summary>Dependents of a cell or range.</summary>
public sealed class DependentResult
{
    public DependentResult(List<TraceItem> items, long dynamicFormulaCells, int unindexedSheets)
    {
        Items = items;
        DynamicFormulaCells = dynamicFormulaCells;
        UnindexedSheets = unindexedSheets;
    }

    public List<TraceItem> Items { get; }
    /// <summary>Formula cells using OFFSET/INDIRECT anywhere in the index: they may depend on the target without saying so.</summary>
    public long DynamicFormulaCells { get; }
    /// <summary>Sheets that were dirty or never indexed when queried (results may be incomplete).</summary>
    public int UnindexedSheets { get; }
}
