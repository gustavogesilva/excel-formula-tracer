using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using XlRange = Microsoft.Office.Interop.Excel.Range;

namespace Tracer.AddIn.Excel;

/// <summary>Where the user was (workbook, sheet, selection, active cell, scroll), so Esc / Ctrl+Backspace can return there.</summary>
internal sealed class SavedPosition
{
    private Workbook _book = null!;
    private Worksheet _sheet = null!;
    private string _selection = "A1";
    private string _activeCell = "A1";
    private int _scrollRow = 1, _scrollCol = 1;

    public string Description => $"{_book.Name} › {_sheet.Name}!{_activeCell}";

    public static SavedPosition? Capture()
    {
        try
        {
            var app = ExcelHost.App;
            if (app.ActiveSheet is not Worksheet ws) return null;
            var p = new SavedPosition { _book = (Workbook)ws.Parent, _sheet = ws };
            if (app.Selection is XlRange sel) p._selection = sel.Address[false, false];
            if (app.ActiveCell is XlRange ac) p._activeCell = ac.Address[false, false];
            var wnd = app.ActiveWindow;
            if (wnd != null)
            {
                p._scrollRow = wnd.ScrollRow;
                p._scrollCol = wnd.ScrollColumn;
            }
            return p;
        }
        catch (COMException)
        {
            return null;
        }
    }

    public bool Restore()
    {
        try
        {
            var app = ExcelHost.App;
            bool screen = app.ScreenUpdating;
            app.ScreenUpdating = false;
            try
            {
                _book.Activate();
                _sheet.Activate();
                _sheet.Range[_selection].Select();
                _sheet.Range[_activeCell].Activate();
                var wnd = app.ActiveWindow;
                if (wnd != null)
                {
                    wnd.ScrollRow = _scrollRow;
                    wnd.ScrollColumn = _scrollCol;
                }
            }
            finally
            {
                app.ScreenUpdating = screen;
            }
            return true;
        }
        catch (COMException)
        {
            return false; // workbook or sheet was closed/deleted
        }
    }
}
