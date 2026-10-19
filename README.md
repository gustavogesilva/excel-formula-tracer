# excel-formula-tracer

A simple, fast, and free formula tracer add-in for Excel (Windows desktop), in the spirit of Arixcel Explorer:
one keystroke opens a window listing a cell's precedents or dependents across sheets and workbooks, and
Excel's selection follows as you arrow through the list.

## Install

1. Download `FormulaTracer64-packed.xll` (64-bit Excel) or `FormulaTracer-packed.xll` (32-bit) from the latest
   CI run's `FormulaTracer-xll` artifact, or build it yourself (below).
2. Open it in Excel (double-click, or File › Options › Add-ins › Manage Excel Add-ins › Browse to keep it loaded).

Nothing else is installed: no installer, admin rights, registry entries or server. It runs on the .NET Framework 4.8
that ships with Windows 10/11.

## Keys

| Where | Key | Action |
|---|---|---|
| Excel | Ctrl+Q | Trace precedents of the active cell |
| Excel | Ctrl+Shift+Q | Trace dependents of the selection |
| Excel | Ctrl+Backspace | Jump back to where you jumped from (100 steps) |
| Window | ↑ ↓ PgUp PgDn | Move through the list; Excel follows |
| Window | → / double-click | Drill in: precedents of that cell, cells of a range, parts of a name, evaluate OFFSET/INDIRECT |
| Window | ← / Backspace | Back one level |
| Window | Tab | Switch between precedents and dependents of the same cell |
| Window | Ctrl+Q / Ctrl+Shift+Q | Precedents / dependents of the selected row |
| Window | E | Trace the error (`#REF!`, `#N/A`, …) in the selected row, or in the traced cell, back to its root cell(s) |
| Window | Enter | Jump to the selected row and close |
| Window | Esc | Return to where you started and close |
| Window | Ctrl+C | Copy the selected address |

The Tracer ribbon tab has a **Trace error** button for the active cell. It also switches the shortcuts to Ctrl+[ / Ctrl+], which replaces Excel's built-in Go To Precedents/Dependents, turns index-at-idle on or off, and rebuilds the index.

## What it understands

Same-sheet, cross-sheet and cross-workbook references; closed external workbooks (shown, not navigable); 3D
references (`Jan:Dec!A1`); whole rows/columns; defined names (workbook- and sheet-scoped, relative, nested,
LAMBDA names called like functions); structured references (`Table1[Col]`, `[@Col]`, `[#Headers]`, …);
spill references (`A1#`); LET/LAMBDA locals (never mistaken for names); OFFSET/INDIRECT (evaluated on demand
for precedents); references to deleted sheets (`#REF!A1`, shown as broken rather than as a same-sheet cell).

### Tracing an error to its root

**E** (or **Trace error**) walks back from a cell showing an error and lists the cell(s) where that error starts,
each with the reason.

- **What it follows.** Errors pass through formulas unchanged, so only precedents holding the *same* error are
  followed. Ranges are scanned with one bulk read.
- **When a cell counts as a root:**
  - its formula contains `#REF!` (a referenced cell, row, column or sheet was deleted);
  - the error is typed in as a value;
  - an OFFSET/INDIRECT in it evaluates to the error (e.g. it points off the sheet);
  - it uses an unknown name (`#NAME?`);
  - none of its precedents carry the error, so its own function produced it (INDEX out of range, VLOOKUP `#N/A`,
    division by zero).
- **Several roots.** All roots are listed; → on a root shows its precedents.
- **Limits.** The search stops after 5,000 cells. A circular error with no root is reported as such.

## How it is fast

- **No dependence on Excel's trace API.** `Range.DirectPrecedents`/`DirectDependents` only see the active sheet,
  so formulas are parsed directly.
- **Bulk reads.** A sheet's formulas come back as one `object[,]` per 64k cells (`Range.Formula2R1C1`), instead of
  one COM call per cell.
- **Parse each distinct formula once.** Copied formulas have identical R1C1 text, so a 423k-cell model has about
  1.5k distinct formulas to parse.
- **Compressed dependents index.** Copied formulas are grouped into rectangular blocks, and the index stores one
  rule per (block, reference). A query solves, in O(1) per rule, which cells of a block point at the target.
- **Lazy, incremental upkeep.** An edit only marks its sheet dirty. Dirty sheets are re-read on the next dependents
  trace, or one per idle tick in the background.

Engine timings from `dotnet run -c Release --project bench/Tracer.Bench`:

- Workload: a synthetic 21-sheet model with 423,276 formula cells.
- Machine: one Linux container with 4 vCPUs.
- Excel COM read time is **not** included; it can only be measured inside Excel.

| Operation | Result |
|---|---|
| Parse | 2.0 M distinct formulas/s (~0.5 µs each) |
| Full dependents index build | 31 ms median (423k cells → 10k blocks → 16.6k rules) |
| Index memory | 2.3 MB retained |
| Dependents query, single cell | p50 24 µs, p99 83 µs |
| Precedents (engine part) | p50 5 µs |

Correctness of the compressed index is checked against a per-cell brute force on random grids
(`DependencyIndexTests.IndexMatchesBruteForceOracle`).

## Layout

```
src/Tracer.Core       netstandard2.0, no Excel references: parser, binder, index, trace engine
src/Tracer.AddIn      net48 Excel-DNA add-in: COM access, window, hotkeys, ribbon, event sync
tests/Tracer.Core.Tests
bench/Tracer.Bench
docs/WINDOWS_CHECKLIST.md   manual test plan for Excel
```

## Build

Requires the .NET 8 SDK. This works on Windows or Linux: Excel-DNA 1.9 packs `.xll` files on both.

```
dotnet test tests/Tracer.Core.Tests
dotnet build src/Tracer.AddIn -c Release
# -> src/Tracer.AddIn/bin/Release/net48/publish/FormulaTracer64-packed.xll
```

## Known limitations

- Dependents reached only through OFFSET/INDIRECT are not indexed (the same as Arixcel). The window says how many
  such formulas exist.
- Excel rewrites references in other sheets after cut/paste and row/column inserts without raising events for
  them. Whole-row/column edits re-index everything, but a single-cell cut/paste can leave other sheets stale until
  their next edit. Use **Rebuild index** if results look wrong.
- Legacy CSE array formulas spanning several cells are indexed per cell, which is correct only for their
  top-left cell.
- Windows desktop Excel only (no Mac/web).

## License

MIT
