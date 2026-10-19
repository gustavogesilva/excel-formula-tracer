using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using XlRange = Microsoft.Office.Interop.Excel.Range;

namespace Tracer.AddIn.Excel;

/// <summary>
/// Keeps the dependents index honest with cheap invalidation only (no reading on edit):
/// an edit marks its sheet dirty; structural edits mark everything dirty, because Excel rewrites
/// references in other sheets and workbooks without raising events for them.
/// Dirty sheets are re-read lazily on the next dependents trace, or at idle when pre-building is on.
/// </summary>
internal static class EventSync
{
    // Keep the Application RCW alive, or the COM event sinks are collected and events silently stop.
    private static Application? _app;
    private static bool _prebuildQueued;

    public static void Attach()
    {
        _app = ExcelHost.App;
        _app.SheetChange += OnSheetChange;
        _app.WorkbookOpen += OnWorkbookOpen;
        _app.WorkbookActivate += OnWorkbookActivate;
        _app.WorkbookNewSheet += OnWorkbookNewSheet;
        _app.WorkbookAfterSave += OnWorkbookAfterSave;
    }

    public static void Detach()
    {
        if (_app == null) return;
        _app.SheetChange -= OnSheetChange;
        _app.WorkbookOpen -= OnWorkbookOpen;
        _app.WorkbookActivate -= OnWorkbookActivate;
        _app.WorkbookNewSheet -= OnWorkbookNewSheet;
        _app.WorkbookAfterSave -= OnWorkbookAfterSave;
        _app = null;
    }

    private static void OnSheetChange(object sh, XlRange target)
    {
        try
        {
            if (sh is not Worksheet ws) return;
            string book = ((Workbook)ws.Parent).Name;
            var engine = ExcelHost.Instance.Engine;
            bool structural = target.Rows.Count == ws.Rows.Count || target.Columns.Count == ws.Columns.Count;
            if (structural)
            {
                // Inserted/deleted rows or columns shift references everywhere, including other workbooks.
                engine.MarkAllDirty();
                ExcelHost.Instance.InvalidateStructure(book);
            }
            else
            {
                engine.MarkSheetDirty(book, ws.Name);
            }
        }
        catch (Exception)
        {
            // Never let an exception escape into Excel's event dispatch; worst case the index is re-read later.
        }
    }

    private static void OnWorkbookOpen(Workbook wb) => SchedulePrebuild();

    private static void OnWorkbookActivate(Workbook wb) => SchedulePrebuild();

    private static void OnWorkbookNewSheet(Workbook wb, object sh)
    {
        try
        {
            ExcelHost.Instance.InvalidateStructure(wb.Name);
        }
        catch (Exception)
        {
        }
    }

    private static void OnWorkbookAfterSave(Workbook wb, bool success) => SchedulePrebuild();

    /// <summary>Indexes pending sheets one per idle callback so Excel stays responsive between sheets.</summary>
    public static void SchedulePrebuild()
    {
        if (_prebuildQueued || !Settings.Current.Prebuild) return;
        _prebuildQueued = true;
        try
        {
            QueuePrebuildStep();
        }
        catch (Exception)
        {
            _prebuildQueued = false;
        }
    }

    private static void QueuePrebuildStep()
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            _prebuildQueued = false;
            try
            {
                if (ExcelHost.Instance.IndexNextPendingSheet()) SchedulePrebuild();
            }
            catch (Exception)
            {
                // Excel was busy or a workbook closed mid-way; the next trace indexes synchronously.
            }
        });
    }
}
