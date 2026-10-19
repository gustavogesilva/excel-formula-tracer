using System;
using System.Collections.Generic;
using System.Linq;

namespace Tracer.Setup;

/// <summary>
/// Excel's auto-open list: values <c>OPEN</c>, <c>OPEN1</c>, <c>OPEN2</c>… under <c>HKCU\Software\Microsoft\Office\&lt;ver&gt;\Excel\Options</c>,
/// with data <c>/R "C:\path\addin.xll"</c> for XLLs. Excel stops reading at the first gap, so lists are always
/// rewritten as consecutive values. Pure logic (no registry access) so it can be unit-tested anywhere.
/// </summary>
internal static class OpenValues
{
    /// <summary>Value name for a position: 0 → <c>OPEN</c>, 3 → <c>OPEN3</c>.</summary>
    public static string NameAt(int index) => index == 0 ? "OPEN" : "OPEN" + index;

    /// <summary>Position of an OPEN value name (<c>OPEN</c> → 0, <c>OPEN7</c> → 7), or -1 for other names.</summary>
    public static int IndexOf(string valueName)
    {
        if (!valueName.StartsWith("OPEN", StringComparison.OrdinalIgnoreCase)) return -1;
        if (valueName.Length == 4) return 0;
        string digits = valueName.Substring(4);
        return digits.All(char.IsDigit) && int.TryParse(digits, out int i) && i > 0 ? i : -1;
    }

    /// <summary>The OPEN entries of a key, in Excel's order, including any stranded after a gap.</summary>
    public static List<string> Ordered(IEnumerable<KeyValuePair<string, string>> values) =>
        values.Select(v => (Index: IndexOf(v.Key), v.Value))
              .Where(v => v.Index >= 0)
              .OrderBy(v => v.Index)
              .Select(v => v.Value)
              .ToList();

    public static string Entry(string xllPath) => "/R \"" + xllPath + "\"";

    /// <summary>The file path in an OPEN value, without switches such as <c>/R</c> and without quotes.</summary>
    public static string PathOf(string data)
    {
        string s = data.Trim();
        while (s.StartsWith("/", StringComparison.Ordinal))
        {
            int space = s.IndexOf(' ');
            if (space < 0) return string.Empty;
            s = s.Substring(space + 1).TrimStart();
        }
        return s.Trim('"');
    }

    /// <summary>
    /// True for any Formula Tracer XLL (installed copy or a hand-loaded <c>FormulaTracer64-packed.xll</c>),
    /// given an OPEN value or a plain path (Add-in Manager value names are paths).
    /// </summary>
    public static bool IsOurs(string valueOrPath)
    {
        string path = PathOf(valueOrPath);
        int slash = Math.Max(path.LastIndexOf('\\'), path.LastIndexOf('/'));
        string file = slash >= 0 ? path.Substring(slash + 1) : path;
        return file.StartsWith("FormulaTracer", StringComparison.OrdinalIgnoreCase)
            && file.EndsWith(".xll", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>The list after installing: other add-ins keep their order, ours goes last (once).</summary>
    public static List<string> Install(IEnumerable<string> current, string xllPath)
    {
        var list = Uninstall(current);
        list.Add(Entry(xllPath));
        return list;
    }

    /// <summary>The list after removing every Formula Tracer entry.</summary>
    public static List<string> Uninstall(IEnumerable<string> current) => current.Where(v => !IsOurs(v)).ToList();
}
