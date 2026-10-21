using System;
using System.Collections.Generic;
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
            try
            {
                Hotkeys.Apply(Settings.Current.Shortcuts);
                EventSync.Attach();
                TraceController.Instance.Prewarm();
                EventSync.SchedulePrebuild();
            }
            catch (Exception ex)
            {
                try
                {
                    ExcelHost.App.StatusBar = "Formula Tracer could not start: " + ex.Message;
                }
                catch (Exception)
                {
                }
            }
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
    public static void TracePrecedents() => Run("tracing precedents", () => TraceController.Instance.TraceActive(TraceMode.Precedents));

    [ExcelCommand(Name = "FT_TraceDependents")]
    public static void TraceDependents() => Run("tracing dependents", () => TraceController.Instance.TraceActive(TraceMode.Dependents));

    [ExcelCommand(Name = "FT_TraceError")]
    public static void TraceError() => Run("tracing the error", () => TraceController.Instance.TraceActive(TraceMode.Errors));

    /// <summary>Target of Excel's Ctrl+Z after an edit made in the trace window (registered with OnUndo).</summary>
    [ExcelCommand(Name = "FT_UndoEdit")]
    public static void UndoEdit() => Run("undoing the edit", () => TraceController.Instance.UndoFromExcel());

    [ExcelCommand(Name = "FT_JumpBack")]
    public static void JumpBack() => Run("jumping back", () => TraceController.Instance.JumpBack());

    [ExcelCommand(Name = "FT_RebuildIndex")]
    public static void RebuildIndex() => Run("rebuilding the index", () => TraceController.Instance.RebuildIndex());

    internal const string OutOfMemoryText =
        "Excel ran out of memory while Formula Tracer was reading the workbook. Close other workbooks and try again; 64-bit Excel avoids this.";

    /// <param name="what">What was being done, for the message (e.g. "tracing dependents").</param>
    internal static void Run(string what, Action action)
    {
        try
        {
            action();
        }
        catch (Exception ex)
        {
            string message = ChunkedReads.IsOutOfMemory(ex) ? OutOfMemoryText : ex.Message;
            MessageBox.Show($"Formula Tracer hit an error while {what}:\n\n{message}", "Formula Tracer", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}

/// <summary>Binds the configured shortcuts with <c>Application.OnKey</c>, so they can be changed at runtime.</summary>
internal static class Hotkeys
{
    private static string[] _bound = Array.Empty<string>();

    /// <summary>Rebinds every shortcut in <paramref name="map"/>. Returns the ones Excel refused (e.g. "Ctrl+; (Trace error)").</summary>
    public static List<string> Apply(ShortcutMap map)
    {
        Clear();
        var app = ExcelHost.App;
        var bound = new List<string>();
        var refused = new List<string>();
        foreach (var (action, key) in map.All())
        {
            string onKey = key.ToOnKey();
            try
            {
                app.OnKey(onKey, ShortcutMap.Macro(action));
                bound.Add(onKey);
            }
            catch (COMException)
            {
                refused.Add($"{key} ({ShortcutMap.Label(action)})");
            }
        }
        _bound = bound.ToArray();
        return refused;
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
