# Server rack inventory — VBA installation and maintenance

This project consists of three Excel worksheets, four standard VBA modules, three worksheet events, and one dynamic lookup formula.

## Installation — desktop Excel

1. **Make a backup** of the working `.xlsm` file before replacing code.
2. Open **Alt+F11** (Visual Basic Editor).
3. Under **Modules**, add or replace the **entire contents** of these four standard modules:
   - [modRackData.bas](modRackData.bas) — shared limits, table guards, text/U parsing (**NEW REQUIRED MODULE**).
   - [modRackSearch.bas](modRackSearch.bas) — exclusive search inputs, including multi-cell paste.
   - [modRackRegistry.bas](modRackRegistry.bas) — automatic registration and manual rack audits.
   - [modRackElevation.bas](modRackElevation.bas) — renders 1–42U elevation.
4. Under **Microsoft Excel Objects**, replace each sheet's existing `Worksheet_Change` procedure with the appropriate snippet (do **not** put these in standard modules):
   - **Rack Lookup** → [RackLookup_SheetCode.txt](RackLookup_SheetCode.txt).
   - **Server Inventory** → [ServerInventory_SheetCode.txt](ServerInventory_SheetCode.txt).
   - **Rack Info** → [RackInfo_SheetCode.txt](RackInfo_SheetCode.txt).
   Only one `Worksheet_Change` may exist per worksheet. Preserve other, unrelated sheet event procedures if present.
5. In **Rack Lookup!E3**, use the **full** dynamic formula from [server-rack-lookup-formula.md](../server-rack-lookup-formula.md). Keep the spill destination E3:I free. It cannot be placed inside a Table.
6. Run **Debug > Compile VBAProject**, then save and reopen if appropriate. Ensure VBA editor error trapping is **Break on Unhandled Errors**.
7. Run `SyncRackRegistry` once from **Alt+F8** to register rack IDs that were present before worksheet events were installed; then run `ValidateRackRegistry` once to check duplicates and bad heights.

## Workbook layout — exact names

| Worksheet | Excel table | Required content |
| --- | --- | --- |
| **Server Inventory** | `tblInventory` | `Rack ID`, `U Position`, `Hostname`, `Serial Number`, `Asset Tag`; optional `Height (U)` |
| **Rack Info** | `tblRacks` | `Rack ID`, `Barcode`, `Rack Height` |
| **Rack Lookup** | none required | C3 Rack ID, C4 U position, C6 Hostname, C8 Serial Number, C10 Asset Tag; E2:I2 results headers, E3 spill formula, K2:M48 rack elevation |

The `Height (U)` field describes **equipment height**, while `Rack Height` describes the **rack height**. Rack Height is blank for a newly registered rack, which means **42U** by default. The rack height must be an integer from **1 through 42**, never over 42.

The Brady printer can generate barcode artwork; the workbook does not generate barcodes. Rack IDs and barcode fields can be maintained independently.

## Changes and refreshes

- Scanning or editing a search input clears conflicting inputs then redraws the elevation; a pasted Rack ID + U pair is preserved. When multiple searches are pasted simultaneously, precedence is hostname → serial → asset → rack/U.
- Editing or pasting rows inside `tblInventory` redraws the current elevation; rack ID changes also synchronize new rack IDs into `tblRacks`.
- Editing `tblRacks[Rack ID]` or `tblRacks[Rack Height]` validates the rack registry and redraws the elevation.
- Rack synchronization **appends only**. It never deletes racks or modifies existing IDs, heights, or barcodes. The sync is case-insensitive and trims whitespace to avoid duplicate registrations.
- A mistaken new Rack ID in inventory is still treated as a new rack. Correcting the typo does **not** delete the incorrect registry row: review and remove it manually after confirming it is unused.
- Changing a formula result, importing data, or deleting an entire table can behave differently from typing inside a table; `Worksheet_Change` is not a general recalculation monitor. Use `SyncRackRegistry` / `RenderRackElevation` manually if needed.
- Clear Search button macro: `ClearRackSearch`. Manual redraw: `RenderRackElevation`. Manual registry audit: `ValidateRackRegistry`.

## Rendering / validation rules

- Visualization uses **Rack Lookup!K2:M48** for a 42U rack; code also clears K2:M126 so older taller diagrams are removed. Do not store other data in that reserved area.
- `U Position` is the **topmost U** a server occupies. `40` and `U40` work equivalently. Rack+U search highlights devices whose **topmost U** matches; lower occupied units of multi-U equipment are still drawn.
- Equipment `Height (U)` may be omitted or left blank; then it is treated as 1U. A nonblank invalid height is counted and **not** drawn, rather than silently defaulted to 1U.
- Overlapping equipment U slots are red, populated slots blue, matched slots yellow, and empty slots gray. Invalid U positions or equipment extending out of a rack are skipped and counted.
- Duplicate Rack IDs in `tblRacks` are reported by `ValidateRackRegistry` and marked in the drawing when a duplicated rack is selected. For backward compatibility the elevation uses the first matching registry record until duplicates are corrected.
- Blank Rack Height uses 42U. Invalid or missing rack metadata produces a warning in the drawing; invalid heights are not silently treated as verified. Missing `tblRacks` / required columns produce an explicit error rather than masking the broken structure.

## Resource safety

- Before iterating, the code checks that `tblInventory` has no more than **50,000** data rows and `tblRacks` has no more than **10,000**.
- **Do not resize a Table to an entire sheet/column.** Use **Table Design > Resize Table** to include only its actual records. Excel Tables expand automatically as you add real rows.
- An oversized Table causes a descriptive error instead of a long scan or out-of-memory condition.
- The code preserves and restores `Application.EnableEvents` and screen updating settings during modifications. Avoid interrupting execution without restoring those settings if Excel is paused in the debugger.

## Compatibility and limitations

- Requires desktop Microsoft Excel supporting **LET / FILTER / CHOOSE** dynamic arrays and VBA. The registry uses late-bound `Scripting.Dictionary` on Windows.
- Excel lookup text comparison is case-insensitive. VBA uses trimmed case-insensitive text keys. Store serial numbers / Rack IDs as **Text** if leading zeros are significant.
- Number strings for U should use the standard form `40` or `U40`. Rack metadata and equipment height must be whole numbers 1–42.
- This project is source code/documentation, **not a precompiled, tested .xlsm**. Complete the verification in [TEST_PLAN.md](TEST_PLAN.md) on a backup workbook before deployment.
- Keep actual organization inventory/hostnames/serials out of this public code repository.
