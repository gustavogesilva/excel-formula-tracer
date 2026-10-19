using System;
using System.Collections.Generic;

namespace Tracer.Core.Graph;

/// <summary>Interns (workbook, sheet) pairs to small integer ids. Case-insensitive like Excel.</summary>
public sealed class SheetRegistry
{
    private readonly Dictionary<string, int> _ids = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<(string Book, string Sheet)> _names = new();

    public int Id(string book, string sheet)
    {
        string key = book + "\u0001" + sheet;
        if (_ids.TryGetValue(key, out int id)) return id;
        id = _names.Count;
        _names.Add((book, sheet));
        _ids.Add(key, id);
        return id;
    }

    public bool TryGetId(string book, string sheet, out int id) => _ids.TryGetValue(book + "\u0001" + sheet, out id);

    public (string Book, string Sheet) this[int id] => _names[id];

    public int Count => _names.Count;
}
