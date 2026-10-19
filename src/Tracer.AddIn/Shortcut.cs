using System;
using System.Collections.Generic;
using System.Linq;

namespace Tracer.AddIn;

/// <summary>
/// A keyboard shortcut such as <c>Ctrl+Shift+Q</c>, convertible to Excel's <c>Application.OnKey</c> syntax and to the
/// WinForms <c>Keys</c> value (virtual-key code plus modifier flags). Pure: no WinForms or Excel types, unit-tested.
/// Punctuation keys assume the US keyboard layout.
/// </summary>
internal readonly struct Shortcut : IEquatable<Shortcut>
{
    // Same values as System.Windows.Forms.Keys.Shift / Control / Alt.
    public const int ShiftFlag = 0x10000, CtrlFlag = 0x20000, AltFlag = 0x40000;

    private sealed class KeyInfo
    {
        public KeyInfo(string name, int vk, string onKey)
        {
            Name = name;
            Vk = vk;
            OnKey = onKey;
        }

        public string Name { get; }
        public int Vk { get; }
        public string OnKey { get; }
    }

    private static readonly List<KeyInfo> Table = BuildTable();
    private static readonly Dictionary<string, string> Aliases = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Del"] = "Delete", ["Ins"] = "Insert", ["PgUp"] = "PageUp", ["PgDn"] = "PageDown", ["Back"] = "Backspace",
    };

    public Shortcut(int virtualKey, bool ctrl, bool shift, bool alt)
    {
        VirtualKey = virtualKey;
        Ctrl = ctrl;
        Shift = shift;
        Alt = alt;
    }

    public static Shortcut None => default;

    public int VirtualKey { get; }
    public bool Ctrl { get; }
    public bool Shift { get; }
    public bool Alt { get; }
    public bool IsEmpty => VirtualKey == 0;
    public bool IsFunctionKey => VirtualKey >= 0x70 && VirtualKey <= 0x7B;

    /// <summary>The WinForms <c>KeyEventArgs.KeyData</c> value for this shortcut.</summary>
    public int KeyData => IsEmpty ? 0 : VirtualKey | (Ctrl ? CtrlFlag : 0) | (Shift ? ShiftFlag : 0) | (Alt ? AltFlag : 0);

    /// <summary>From a WinForms <c>KeyData</c>; <see cref="None"/> when the key is not one a shortcut can use.</summary>
    public static Shortcut FromKeyData(int keyData)
    {
        int vk = keyData & 0xFFFF;
        if (Find(vk) == null) return None;
        return new Shortcut(vk, (keyData & CtrlFlag) != 0, (keyData & ShiftFlag) != 0, (keyData & AltFlag) != 0);
    }

    /// <summary>Parses <c>Ctrl+Shift+Q</c>, <c>Alt+F2</c>, <c>Ctrl+[</c>; empty text is <see cref="None"/>.</summary>
    public static bool TryParse(string? text, out Shortcut shortcut)
    {
        shortcut = None;
        if (string.IsNullOrWhiteSpace(text)) return true;
        var parts = text!.Split('+').Select(p => p.Trim()).ToArray();
        if (parts.Any(p => p.Length == 0)) return false;

        bool ctrl = false, shift = false, alt = false;
        for (int i = 0; i < parts.Length - 1; i++)
        {
            switch (parts[i].ToUpperInvariant())
            {
                case "CTRL": case "CONTROL": ctrl = true; break;
                case "SHIFT": shift = true; break;
                case "ALT": alt = true; break;
                default: return false;
            }
        }
        string keyName = parts[parts.Length - 1];
        if (Aliases.TryGetValue(keyName, out var canonical)) keyName = canonical;
        var key = Table.FirstOrDefault(k => string.Equals(k.Name, keyName, StringComparison.OrdinalIgnoreCase));
        if (key == null) return false;
        shortcut = new Shortcut(key.Vk, ctrl, shift, alt);
        return true;
    }

    /// <summary>The key string for <c>Application.OnKey</c>, e.g. <c>^+q</c>, <c>^{[}</c>, <c>%{F2}</c>.</summary>
    public string ToOnKey()
    {
        if (IsEmpty) return string.Empty;
        return (Ctrl ? "^" : "") + (Shift ? "+" : "") + (Alt ? "%" : "") + Find(VirtualKey)!.OnKey;
    }

    /// <summary>Why this shortcut cannot be used, or null when it can.</summary>
    public string? Problem()
    {
        if (IsEmpty) return null;
        if (!Ctrl && !Alt && !IsFunctionKey) return "Shortcuts need Ctrl or Alt (F1–F12 may be used alone), so typing in cells is never captured.";
        return null;
    }

    /// <summary>A well-known Excel command this would replace while the add-in is loaded, or null.</summary>
    public string? ExcelConflict()
    {
        if (!Ctrl || Shift || Alt) return null;
        return VirtualKey switch
        {
            'C' => "Copy",
            'V' => "Paste",
            'X' => "Cut",
            'Z' => "Undo",
            'Y' => "Redo",
            'S' => "Save",
            'A' => "Select All",
            'F' => "Find",
            'P' => "Print",
            'N' => "New Workbook",
            'O' => "Open",
            'W' => "Close Workbook",
            _ => null,
        };
    }

    public override string ToString()
    {
        if (IsEmpty) return string.Empty;
        return (Ctrl ? "Ctrl+" : "") + (Shift ? "Shift+" : "") + (Alt ? "Alt+" : "") + Find(VirtualKey)!.Name;
    }

    public bool Equals(Shortcut other) => KeyData == other.KeyData;
    public override bool Equals(object? obj) => obj is Shortcut s && Equals(s);
    public override int GetHashCode() => KeyData;

    private static KeyInfo? Find(int vk) => Table.FirstOrDefault(k => k.Vk == vk);

    private static List<KeyInfo> BuildTable()
    {
        var t = new List<KeyInfo>();
        for (char c = 'A'; c <= 'Z'; c++) t.Add(new KeyInfo(c.ToString(), c, char.ToLowerInvariant(c).ToString()));
        for (char c = '0'; c <= '9'; c++) t.Add(new KeyInfo(c.ToString(), c, c.ToString()));
        for (int i = 1; i <= 12; i++) t.Add(new KeyInfo("F" + i, 0x6F + i, "{F" + i + "}"));
        t.Add(new KeyInfo("Backspace", 0x08, "{BACKSPACE}"));
        t.Add(new KeyInfo("Delete", 0x2E, "{DEL}"));
        t.Add(new KeyInfo("Insert", 0x2D, "{INSERT}"));
        t.Add(new KeyInfo("Home", 0x24, "{HOME}"));
        t.Add(new KeyInfo("End", 0x23, "{END}"));
        t.Add(new KeyInfo("PageUp", 0x21, "{PGUP}"));
        t.Add(new KeyInfo("PageDown", 0x22, "{PGDN}"));
        t.Add(new KeyInfo("Up", 0x26, "{UP}"));
        t.Add(new KeyInfo("Down", 0x28, "{DOWN}"));
        t.Add(new KeyInfo("Left", 0x25, "{LEFT}"));
        t.Add(new KeyInfo("Right", 0x27, "{RIGHT}"));
        // US-layout punctuation (VK_OEM_*). OnKey needs braces around [ and ] (and + ^ % ~ ( ) { }, not offered here).
        t.Add(new KeyInfo(";", 0xBA, ";"));
        t.Add(new KeyInfo("=", 0xBB, "="));
        t.Add(new KeyInfo(",", 0xBC, ","));
        t.Add(new KeyInfo("-", 0xBD, "-"));
        t.Add(new KeyInfo(".", 0xBE, "."));
        t.Add(new KeyInfo("/", 0xBF, "/"));
        t.Add(new KeyInfo("`", 0xC0, "`"));
        t.Add(new KeyInfo("[", 0xDB, "{[}"));
        t.Add(new KeyInfo("\\", 0xDC, "\\"));
        t.Add(new KeyInfo("]", 0xDD, "{]}"));
        t.Add(new KeyInfo("'", 0xDE, "'"));
        return t;
    }
}

/// <summary>The actions that can have global shortcuts.</summary>
internal enum TraceAction
{
    TracePrecedents,
    TraceDependents,
    TraceError,
    JumpBack,
}

/// <summary>Each action's primary and alternate shortcut, with persistence to settings key/value pairs.</summary>
internal sealed class ShortcutMap
{
    public static readonly TraceAction[] Actions =
        { TraceAction.TracePrecedents, TraceAction.TraceDependents, TraceAction.TraceError, TraceAction.JumpBack };

    private readonly Dictionary<TraceAction, Shortcut[]> _keys = Actions.ToDictionary(a => a, _ => new Shortcut[2]);

    public static ShortcutMap Arixcel() => Make("Ctrl+Q", "Ctrl+Shift+Q", "", "Ctrl+Backspace");
    public static ShortcutMap Brackets() => Make("Ctrl+[", "Ctrl+]", "", "Ctrl+Backspace");

    /// <summary>Slot 0 is the primary shortcut, slot 1 the alternate.</summary>
    public Shortcut this[TraceAction action, int slot]
    {
        get => _keys[action][slot];
        set => _keys[action][slot] = value;
    }

    public static string Label(TraceAction a) => a switch
    {
        TraceAction.TracePrecedents => "Trace precedents",
        TraceAction.TraceDependents => "Trace dependents",
        TraceAction.TraceError => "Trace error to its root",
        _ => "Jump back",
    };

    /// <summary>The Excel macro (registered by <c>Commands</c>) each action runs.</summary>
    public static string Macro(TraceAction a) => a switch
    {
        TraceAction.TracePrecedents => "FT_TracePrecedents",
        TraceAction.TraceDependents => "FT_TraceDependents",
        TraceAction.TraceError => "FT_TraceError",
        _ => "FT_JumpBack",
    };

    /// <summary>All assigned shortcuts.</summary>
    public IEnumerable<(TraceAction Action, Shortcut Key)> All() =>
        from a in Actions from k in _keys[a] where !k.IsEmpty select (a, k);

    /// <summary>The action bound to a WinForms <c>KeyData</c>, if any.</summary>
    public TraceAction? Find(int keyData)
    {
        foreach (var (a, k) in All())
            if (k.KeyData == keyData) return a;
        return null;
    }

    /// <summary>"Ctrl+Q, Ctrl+[" style text for one action (empty when unassigned).</summary>
    public string Describe(TraceAction a) => string.Join(", ", _keys[a].Where(k => !k.IsEmpty));

    /// <summary>Problems that block saving: unusable keys and keys assigned twice.</summary>
    public List<string> Problems()
    {
        var problems = new List<string>();
        foreach (var (a, k) in All())
            if (k.Problem() is { } p) problems.Add($"{Label(a)}: {k}. {p}");
        foreach (var dup in All().GroupBy(x => x.Key).Where(g => g.Count() > 1))
            problems.Add($"{dup.Key} is assigned to {string.Join(" and ", dup.Select(x => Label(x.Action)).Distinct())}.");
        return problems;
    }

    public ShortcutMap Clone()
    {
        var c = new ShortcutMap();
        foreach (var a in Actions)
            for (int s = 0; s < 2; s++) c[a, s] = this[a, s];
        return c;
    }

    public void Write(IDictionary<string, string> kv)
    {
        foreach (var a in Actions)
        {
            kv["Shortcut." + a] = this[a, 0].ToString();
            kv["Shortcut." + a + ".Alt"] = this[a, 1].ToString();
        }
    }

    /// <summary>
    /// Reads saved shortcuts. Without any (settings from 0.1.2 or earlier) it maps the old <c>Keys</c> preset:
    /// Arixcel (default), Brackets, or Both (Ctrl+Q / Ctrl+Shift+Q with Ctrl+[ / Ctrl+] as alternates).
    /// Unparseable entries fall back to that default.
    /// </summary>
    public static ShortcutMap Read(IReadOnlyDictionary<string, string> kv)
    {
        kv.TryGetValue("Keys", out var preset);
        var map = string.Equals(preset, "Brackets", StringComparison.OrdinalIgnoreCase) ? Brackets() : Arixcel();
        if (string.Equals(preset, "Both", StringComparison.OrdinalIgnoreCase))
        {
            var b = Brackets();
            map[TraceAction.TracePrecedents, 1] = b[TraceAction.TracePrecedents, 0];
            map[TraceAction.TraceDependents, 1] = b[TraceAction.TraceDependents, 0];
        }

        foreach (var a in Actions)
        {
            if (kv.TryGetValue("Shortcut." + a, out var primary) && Shortcut.TryParse(primary, out var p)) map[a, 0] = p;
            if (kv.TryGetValue("Shortcut." + a + ".Alt", out var alternate) && Shortcut.TryParse(alternate, out var s)) map[a, 1] = s;
        }
        return map;
    }

    private static ShortcutMap Make(string prec, string dep, string error, string back)
    {
        var m = new ShortcutMap();
        string[] keys = { prec, dep, error, back };
        for (int i = 0; i < Actions.Length; i++)
        {
            Shortcut.TryParse(keys[i], out var k);
            m[Actions[i], 0] = k;
        }
        return m;
    }
}
