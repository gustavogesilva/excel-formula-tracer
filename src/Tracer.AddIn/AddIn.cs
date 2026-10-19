using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using Tracer.AddIn.Excel;
using Tracer.AddIn.UI;

namespace Tracer.AddIn;

/// <summary>Add-in lifecycle: hotkeys, Excel events, and pre-warming the window at the first idle moment.</summary>
public sealed class AddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        try
        {
            Application.EnableVisualStyles();
        }
        catch (InvalidOperationException)
        {
            // Another component already created windows in this process; classic styles are fine.
        }

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            Hotkeys.Apply(Settings.Current.Keys);
            EventSync.Attach();
            TraceController.Instance.Prewarm();
            EventSync.SchedulePrebuild();
        });
    }

    public void AutoClose()
    {
        try
        {
            Hotkeys.Clear();
            EventSync.Detach();
        }
        catch (COMException)
        {
        }
    }
}

/// <summary>Macros bound to hotkeys and the ribbon. Each catches everything: an add-in must never take Excel down.</summary>
public static class Commands
{
    [ExcelCommand(Name = "FT_TracePrecedents")]
    public static void TracePrecedents() => Run(() => TraceController.Instance.TraceActive(TraceMode.Precedents));

    [ExcelCommand(Name = "FT_TraceDependents")]
    public static void TraceDependents() => Run(() => TraceController.Instance.TraceActive(TraceMode.Dependents));

    [ExcelCommand(Name = "FT_JumpBack")]
    public static void JumpBack() => Run(() => TraceController.Instance.JumpBack());

    [ExcelCommand(Name = "FT_RebuildIndex")]
    public static void RebuildIndex() => Run(() => TraceController.Instance.RebuildIndex());

    internal static void Run(Action action)
    {
        try
        {
            action();
        }
        catch (Exception ex)
        {
            MessageBox.Show("Formula Tracer hit an error:\n\n" + ex.Message, "Formula Tracer", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}

/// <summary>Binds shortcut keys with <c>Application.OnKey</c> so they can be changed at runtime.</summary>
internal static class Hotkeys
{
    private static string[] _bound = Array.Empty<string>();

    public static void Apply(KeyPreset preset)
    {
        Clear();
        var app = ExcelHost.App;
        var keys = new System.Collections.Generic.List<string>();
        void Bind(string key, string macro)
        {
            try
            {
                app.OnKey(key, macro);
                keys.Add(key);
            }
            catch (COMException)
            {
            }
        }

        if (preset != KeyPreset.Brackets)
        {
            Bind("^q", "FT_TracePrecedents");
            Bind("^+q", "FT_TraceDependents");
        }
        if (preset != KeyPreset.Arixcel)
        {
            // Brackets must be escaped in OnKey strings.
            Bind("^{[}", "FT_TracePrecedents");
            Bind("^{]}", "FT_TraceDependents");
        }
        Bind("^{BACKSPACE}", "FT_JumpBack");
        _bound = keys.ToArray();
    }

    public static void Clear()
    {
        var app = ExcelHost.App;
        foreach (var k in _bound)
        {
            try
            {
                app.OnKey(k);
            }
            catch (COMException)
            {
            }
        }
        _bound = Array.Empty<string>();
    }
}
