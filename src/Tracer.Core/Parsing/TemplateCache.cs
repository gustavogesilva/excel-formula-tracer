using System;
using System.Collections.Generic;
using Tracer.Core.Model;

namespace Tracer.Core.Parsing;

/// <summary>
/// Interns parsed formulas by their R1C1 text. Copied formulas share one R1C1 string,
/// so a model with 500k formula cells typically needs only a few thousand parses.
/// Not thread-safe; owned by the Excel main thread.
/// </summary>
public sealed class TemplateCache
{
    private readonly Dictionary<string, ParsedTemplate> _byText = new(StringComparer.Ordinal);
    private readonly List<ParsedTemplate> _byId = new();

    public int Count => _byId.Count;
    public long Lookups { get; private set; }

    public ParsedTemplate Get(string formulaR1C1)
    {
        Lookups++;
        if (_byText.TryGetValue(formulaR1C1, out var t)) return t;
        t = TemplateParser.Parse(formulaR1C1);
        t.Id = _byId.Count;
        _byId.Add(t);
        _byText.Add(formulaR1C1, t);
        return t;
    }

    public ParsedTemplate this[int id] => _byId[id];

    public void Clear()
    {
        _byText.Clear();
        _byId.Clear();
        Lookups = 0;
    }
}
