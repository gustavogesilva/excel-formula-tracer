using System.Collections.Generic;
using Tracer.AddIn;
using Xunit;

namespace Tracer.Core.Tests;

public class ShortcutTests
{
    private static Shortcut P(string text)
    {
        Assert.True(Shortcut.TryParse(text, out var s), text);
        return s;
    }

    [Theory]
    [InlineData("Ctrl+Q", "^q")]
    [InlineData("Ctrl+Shift+Q", "^+q")]
    [InlineData("ctrl+shift+alt+p", "^+%p")]
    [InlineData("Ctrl+[", "^{[}")]
    [InlineData("Ctrl+]", "^{]}")]
    [InlineData("Ctrl+Backspace", "^{BACKSPACE}")]
    [InlineData("Alt+F2", "%{F2}")]
    [InlineData("F9", "{F9}")]
    [InlineData("Ctrl+Shift+PgDn", "^+{PGDN}")]
    [InlineData("Ctrl+;", "^;")]
    [InlineData("Ctrl+7", "^7")]
    public void OnKeyStrings(string text, string onKey) => Assert.Equal(onKey, P(text).ToOnKey());

    [Theory]
    [InlineData("ctrl+shift+q", "Ctrl+Shift+Q")]
    [InlineData("Alt+Ctrl+del", "Ctrl+Alt+Delete")]
    [InlineData("Control+[", "Ctrl+[")]
    [InlineData("", "")]
    public void FormatRoundTrip(string text, string formatted)
    {
        Assert.Equal(formatted, P(text).ToString());
        Assert.Equal(P(text), P(formatted));
    }

    [Theory]
    [InlineData("Ctrl+")]
    [InlineData("Hyper+Q")]
    [InlineData("Ctrl+Esc")]
    [InlineData("Ctrl+F13")]
    [InlineData("Ctrl+QQ")]
    public void RejectsUnknownKeys(string text) => Assert.False(Shortcut.TryParse(text, out _));

    [Fact]
    public void KeyDataMatchesWinFormsValues()
    {
        // Keys.Control | Keys.Shift | Keys.Q == 0x20000 | 0x10000 | 0x51
        Assert.Equal(0x20000 | 0x10000 | 0x51, P("Ctrl+Shift+Q").KeyData);
        // Keys.Control | Keys.OemOpenBrackets (0xDB)
        Assert.Equal(0x20000 | 0xDB, P("Ctrl+[").KeyData);
        Assert.Equal(P("Ctrl+Shift+Q"), Shortcut.FromKeyData(0x20000 | 0x10000 | 0x51));
        Assert.True(Shortcut.FromKeyData(0x20000 | 0x11).IsEmpty); // Ctrl alone (ControlKey) is not a shortcut
    }

    [Theory]
    [InlineData("Q", false)]
    [InlineData("Shift+Q", false)]
    [InlineData("Ctrl+Q", true)]
    [InlineData("Alt+Q", true)]
    [InlineData("F8", true)]
    public void NeedsCtrlOrAltUnlessFunctionKey(string text, bool ok) => Assert.Equal(ok, P(text).Problem() == null);

    [Fact]
    public void WarnsAboutCommonExcelKeys()
    {
        Assert.Equal("Copy", P("Ctrl+C").ExcelConflict());
        Assert.Null(P("Ctrl+Shift+C").ExcelConflict());
        Assert.Null(P("Ctrl+Q").ExcelConflict());
    }

    [Fact]
    public void MapDefaultsAndLookup()
    {
        var m = ShortcutMap.Arixcel();
        Assert.Equal("Ctrl+Q", m.Describe(TraceAction.TracePrecedents));
        Assert.Equal("", m.Describe(TraceAction.TraceError));
        Assert.Equal(TraceAction.JumpBack, m.Find(P("Ctrl+Backspace").KeyData));
        Assert.Null(m.Find(P("Ctrl+[").KeyData));
        Assert.Empty(m.Problems());
    }

    [Fact]
    public void DuplicatesAreProblems()
    {
        var m = ShortcutMap.Arixcel();
        m[TraceAction.TraceError, 1] = P("Ctrl+Q");
        var problem = Assert.Single(m.Problems());
        Assert.Contains("Ctrl+Q is assigned to Trace precedents and Trace error to its root", problem);
    }

    [Fact]
    public void PersistenceRoundTrip()
    {
        var m = ShortcutMap.Arixcel();
        m[TraceAction.TraceError, 0] = P("Ctrl+Shift+E");
        m[TraceAction.TracePrecedents, 1] = P("Ctrl+[");
        m[TraceAction.JumpBack, 0] = Shortcut.None; // explicitly cleared stays cleared
        var kv = new Dictionary<string, string>();
        m.Write(kv);
        var back = ShortcutMap.Read(kv);
        foreach (var a in ShortcutMap.Actions)
            Assert.Equal(m.Describe(a), back.Describe(a));
    }

    [Theory]
    [InlineData(null, "Ctrl+Q", "Ctrl+Shift+Q")]
    [InlineData("Arixcel", "Ctrl+Q", "Ctrl+Shift+Q")]
    [InlineData("Brackets", "Ctrl+[", "Ctrl+]")]
    [InlineData("Both", "Ctrl+Q, Ctrl+[", "Ctrl+Shift+Q, Ctrl+]")]
    public void MigratesOldPresets(string? preset, string prec, string dep)
    {
        var kv = new Dictionary<string, string>();
        if (preset != null) kv["Keys"] = preset;
        var m = ShortcutMap.Read(kv);
        Assert.Equal(prec, m.Describe(TraceAction.TracePrecedents));
        Assert.Equal(dep, m.Describe(TraceAction.TraceDependents));
        Assert.Equal("Ctrl+Backspace", m.Describe(TraceAction.JumpBack));
    }
}
