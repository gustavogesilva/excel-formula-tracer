# Windows test checklist

The engine is unit-tested on Linux, but nothing that touches Excel could be run there. Work through this list in
desktop Excel (Microsoft 365, and Excel 2016/2019 if you support them) with `FormulaTracer64-packed.xll` loaded.
Items marked **(assumption)** are behaviours the code relies on that were not verified.

## Setup

- [ ] Load the `.xll`. Expect no security error beyond the usual "enable this add-in" prompt.
- [ ] The **Tracer** ribbon tab appears.
- [ ] Excel starts as fast as it did without the add-in. The window is created at the first idle moment, not
      during load.

## Keyboard and window (assumption: Excel's message loop delivers these keys to a modeless WinForms window)

- [ ] Ctrl+Q on a formula cell opens the window, with the first precedent selected and the formula's references colour-coded.
- [ ] ↑/↓ move the list, and Excel's selection follows, including onto other sheets and other open workbooks.
- [ ] After each move, keyboard focus stays in the trace window.
- [ ] → drills into the selected row; ← goes back; Tab toggles precedents/dependents.
- [ ] Enter jumps and closes; Esc restores the original sheet, selection and scroll position.
- [ ] Ctrl+C copies the address. Confirm that Excel does not intercept it.
- [ ] Ctrl+Backspace (in Excel, after an Enter-jump) returns to the previous position.
- [ ] The ribbon's Shortcuts option set to `Ctrl+[ / Ctrl+]` binds those keys **(assumption: OnKey accepts `^{[}`)**.
      If it doesn't, try `^[` in `Hotkeys.Apply`.
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
