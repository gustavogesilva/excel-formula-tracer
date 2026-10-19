# Windows test checklist

The engine is unit-tested on Linux, but nothing that touches Excel could be run there. Work through this list in
desktop Excel (Microsoft 365, and Excel 2016/2019 if you support them) with `FormulaTracer64-packed.xll` loaded.
Items marked **(assumption)** are behaviours the code relies on that were not verified.

## Installer (`FormulaTracerSetup.exe`)

- [ ] Running the downloaded exe shows SmartScreen ("More info › Run anyway") but **no** UAC/admin prompt.
- [ ] The dialog shows the detected Excel version, bitness and path, and "not installed".
- [ ] With Excel open, Install asks you to close Excel; Retry continues once it is closed.
- [ ] After Install:
  - [ ] `%LOCALAPPDATA%\FormulaTracer` holds `FormulaTracer.xll` and `FormulaTracerSetup.exe`;
  - [ ] `HKCU\Software\Microsoft\Office\16.0\Excel\Options` has an `OPEN…` value `/R "…\FormulaTracer.xll"`;
  - [ ] other add-ins' `OPEN` values are unchanged and still consecutive.
- [ ] Start Excel:
  - [ ] the **Tracer** tab is on the ribbon;
  - [ ] File › Options › Add-ins lists *Formula Tracer* as active;
  - [ ] no "blocked add-in" warning appears (the file carries no Mark-of-the-Web);
  - [ ] the tab is still there after closing and reopening Excel.
- [ ] 32-bit Excel gets the 32-bit add-in. Check that it loads.
- [ ] A copy you loaded by hand earlier (`…\Downloads\FormulaTracer64-packed.xll`) is replaced, not loaded twice.
- [ ] Running the installer again shows "Update" and keeps everything working.
- [ ] Settings › Apps lists *Formula Tracer for Excel 0.1.0*. Uninstall from there, then confirm:
  - [ ] the folder disappears within a few seconds;
  - [ ] Excel starts without the add-in;
  - [ ] the `OPEN` values are renumbered without gaps.
- [ ] `FormulaTracerSetup.exe /install /quiet` and `/uninstall /quiet` work silently. Exit code 2 means Excel was running.
- [ ] **Close** and **Esc** close the installer dialog.
- [ ] Properties › Details of `FormulaTracerSetup.exe` show company "Gustavo Silva" and version 0.1.3. Settings › Apps
      shows publisher "Gustavo Silva".

## Setup

- [ ] Load the `.xll`. Expect no security error beyond the usual "enable this add-in" prompt.
- [ ] The **Tracer** ribbon tab appears.
- [ ] Excel starts as fast as it did without the add-in. The window is created at the first idle moment, not
      during load.

## Keyboard and window (assumption: Excel's message loop delivers these keys to a modeless WinForms window)

- [ ] Ctrl+Q on a formula cell opens the window. Expect:
  - the bold top row is the traced cell and is selected;
  - the arguments are listed under it;
  - the formula's references are colour-coded.
- [ ] One-argument formula (`=Sheet2!B5`): ↓ **or** → from the top row moves Excel to Sheet2!B5, switching the
      tab, without Enter.
- [ ] ↑/↓ move the list, and Excel's selection follows, including onto other sheets and other open workbooks.
- [ ] After each move, keyboard focus stays in the trace window.
- [ ] → on an argument drills into it, and its own top row is selected.
- [ ] ← on an argument returns to the top row (Excel follows); ← on the top row goes back one level.
- [ ] Tab toggles precedents/dependents.
- [ ] **Traces that lead nowhere show a message on the status line and no error dialog:**
  - [ ] → into a constant cell, then →, ←, Tab, E;
  - [ ] → on an empty range;
  - [ ] a cell with `=1+2` (no references);
  - [ ] dependents of a cell nothing refers to;
  - [ ] → on `#REF!`, a closed-workbook row, or an unknown name;
  - [ ] → on OFFSET pointing off the sheet;
  - [ ] keys pressed while editing a cell in Excel;
  - [ ] closing a workbook while the window shows its cells, then pressing ↑/↓.
- [ ] Enter jumps and closes; Esc restores the original sheet, selection and scroll position.
- [ ] Ctrl+C copies the address. Confirm that Excel does not intercept it.
- [ ] Ctrl+Backspace (in Excel, after an Enter-jump) returns to the previous position.
- [ ] **Tracer › Shortcuts…:**
  - [ ] set Trace precedents to Ctrl+Alt+P: it works at once in Excel and inside the trace window, and Ctrl+Q no
        longer traces;
  - [ ] set Trace error to Ctrl+Shift+E: it traces the active cell's error;
  - [ ] assigning the same key twice is refused;
  - [ ] Ctrl+C asks for confirmation;
  - [ ] a letter without Ctrl/Alt is refused with a beep;
  - [ ] the ribbon tooltips show the new keys;
  - [ ] the keys survive an Excel restart.
- [ ] **Excel brackets** preset binds Ctrl+[ / Ctrl+] **(assumption: OnKey accepts `^{[}`)**. If Excel refuses
      them, the dialog says so after OK.
- [ ] Moving to a far-away cell scrolls it into view **(assumption: `Application.Goto(range, false)` scrolls when
      the target is off-screen)**.
- [ ] The window stays above Excel when navigating into another workbook window (SDI).

## Precedents coverage

- [ ] Same-sheet, cross-sheet and cross-workbook references, ranges, whole columns (`SUM(B:B)`) and whole rows.
- [ ] Closed external workbook: the row says "(closed)" and isn't navigable.
- [ ] 3D reference `SUM(Jan:Mar!B2)` shows one row per sheet.
- [ ] Defined names:
  - [ ] Workbook- and sheet-scoped names resolve.
  - [ ] Nested names resolve (`Double = Rate*2`).
  - [ ] **(assumption: `Name.RefersToR1C1` offsets are relative to the using cell)** A relative name (defined
        while a cell is selected, e.g. `Prev = A1` from B2) resolves to the cell above-left of each using cell.
- [ ] A LAMBDA name called as `MyFn(A1)` shows `MyFn` and its body's references.
- [ ] Tables: `Table1[Col]`, `[@Col]` inside the table, `Table1[[#Headers],[Col]]`, and a table with its header
      row hidden or a totals row.
- [ ] `OFFSET(A1,2,3)`: → evaluates it and the row shows the resolved address. `INDIRECT("B5")` behaves the same.
- [ ] Spill `A1#`: → expands to the spill range.
- [ ] LET locals are not listed as names: `=LET(x, A1, x*2)` lists only A1.
- [ ] A cell with a constant shows "No formula", and Tab shows its dependents.
- [ ] **(assumption: spilled child cells report an empty `Formula2R1C1`)** Dependents of a cell referenced by a
      dynamic-array formula list the anchor only, not every spilled cell.
- [ ] A text cell entered as `'=abc` is not treated as a formula.

## Error roots (E in the window, or the Trace error ribbon button)

- [ ] **Deleted sheet.** Reference Sheet2!A1 from a chain of formulas, then delete Sheet2. Expect:
  - the formula reads `=#REF!A1`;
  - E on the last cell of the chain lists the cell holding `#REF!A1`, with "contains #REF!";
  - the Ctrl+Q list shows it as a broken row, **not** as A1 on the current sheet **(assumption: Excel writes
    `#REF!A1` for refs to deleted sheets)**.
- [ ] **Deleted row.** A formula referencing a deleted row (`=#REF!*2`) is listed as the root.
- [ ] **Error inside a range.** `SUM(B1:B100)` with one `#REF!` inside: E lists only that one cell.
- [ ] **Function-produced error.** E on `=INDEX(A1:A3,5)` (`#REF!`) or on a VLOOKUP `#N/A` lists that cell, with
      "produced here".
- [ ] **OFFSET off the sheet.** `=SUM(OFFSET(A1,-5,0))` is listed with "evaluates to #REF!" **(assumption:
      `Worksheet.Evaluate` returns the error code for an OFFSET that goes off the sheet)**.
- [ ] **Two independent roots** are both listed.
- [ ] **Mixed errors.** A cell showing `#N/A` whose precedents hold both `#N/A` and `#REF!` follows only the `#N/A` path.
- [ ] **No error.** E on a cell with no error says so.
- [ ] **Ribbon icon.** The Trace error button shows an icon **(assumption: imageMso `TraceError` exists; an
      unknown imageMso just leaves the button without an icon)**.

## Dependents and index upkeep

- [ ] Ctrl+Shift+Q on an input cell lists dependents grouped into blocks (e.g. `Calc!$D$5:$O$5`).
      → on a block lists its cells.
- [ ] Dependents across sheets and across two open workbooks.
- [ ] Edit a formula to reference a new cell, then Ctrl+Shift+Q on that cell: it is listed (the dirty sheet was re-read).
- [ ] Insert a row above referenced cells, then trace again: results are correct (everything was marked dirty).
- [ ] Rename a sheet, then trace again: results are correct.
- [ ] Close a workbook: its dependents disappear from results.
- [ ] Select a range (several cells), then Ctrl+Shift+Q lists dependents of any cell in it.
- [ ] The status line notes OFFSET/INDIRECT formulas when the workbook has them.

## Timings (the status line shows ms when ShowTimings is on)

Record these for a real model, including its size (formula cells, sheets):

| Measurement | Target | Measured |
|---|---|---|
| Ctrl+Q to window shown | < 30 ms | |
| ↓ to Excel selection moved | < 20 ms | |
| First Ctrl+Shift+Q with index-at-idle off (full index incl. COM reads) | < 1.5 s per 500k formula cells | |
| Ctrl+Shift+Q after editing one cell (one sheet re-read) | | |
| Ctrl+Shift+Q with index warm | < 5 ms | |
| Rebuild index (Tracer tab; the Excel status bar shows cells/blocks/rules/ms) | | |

If the full index is slow, the cost will be in the `Formula2R1C1` bulk reads. Try a different `ChunkCells`
in `ExcelHost`.
