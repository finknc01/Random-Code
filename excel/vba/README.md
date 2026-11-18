# Server rack inventory — VBA installation and maintenance

This project starts with the Rack Lookup, Server Inventory, and Rack Info worksheets; shared VBA modules; **eight ActiveX buttons**; Add/Edit/Decom/Audit UserForms; worksheet events; and a dynamic lookup formula. The archive and audit-report sheets are created by their workflows.

## Installation — desktop Excel

1. **Make a backup** of the working `.xlsm` file before replacing code. For all eight buttons, start with the complete [EIGHT_BUTTON_SETUP.md](EIGHT_BUTTON_SETUP.md); the instructions below also document the original shared modules.
2. Open **Alt+F11** (Visual Basic Editor).
3. Under **Modules**, add or replace the **entire contents** of these five standard modules:
   - [modRackData.bas](modRackData.bas) — shared limits, table guards, text/U parsing (**NEW REQUIRED MODULE**).
   - [modRackSearch.bas](modRackSearch.bas) — exclusive search inputs, including multi-cell paste.
   - [modRackRegistry.bas](modRackRegistry.bas) — automatic registration and manual rack audits.
   - [modRackElevation.bas](modRackElevation.bas) — renders 1–42U elevation.
   - [modServerEntry.bas](modServerEntry.bas) — validates and appends a server, including multi-U collision detection (**REQUIRED FOR ADD SERVER**).
   For the original Add/Clear/Refresh ActiveX buttons, retain these standard modules:
   - [buttonClear.bas](buttonClear.bas) → module **buttonClear**, public macro **ClearRackLookupButton**.
   - [buttonRefresh.bas](buttonRefresh.bas) → module **buttonRefresh**, public macro **RefreshRackLookupButton**.
   - [buttonAddServer.bas](buttonAddServer.bas) → module **buttonAddServer**, public macro **OpenAddServerForm** (**requires `frmAddServer` UserForm**).
4. To use **Add Server**, build the `frmAddServer` UserForm and paste [frmAddServer_Code.txt](frmAddServer_Code.txt) into its code-behind. Follow the exact control names in [ADD_SERVER_FORM_SETUP.md](ADD_SERVER_FORM_SETUP.md). A UserForm is **not** a standard module. If you omit this feature, omit its button/Click handler too.
5. Under **Microsoft Excel Objects**, replace each sheet's existing `Worksheet_Change` procedure with the appropriate snippet (do **not** put these in standard modules):
   - **Rack Lookup** → [RackLookup_SheetCode.txt](RackLookup_SheetCode.txt).
   - **Server Inventory** → [ServerInventory_SheetCode.txt](ServerInventory_SheetCode.txt).
   - **Rack Info** → [RackInfo_SheetCode.txt](RackInfo_SheetCode.txt).
   Only one `Worksheet_Change` may exist per worksheet. The Rack Lookup sheet code also contains **ActiveX button Click events**. Preserve other unrelated sheet event procedures.
6. In **Rack Lookup!E3**, use the **full** dynamic formula from [server-rack-lookup-formula.md](../server-rack-lookup-formula.md). Keep the spill destination E3:I free. It cannot be placed inside a Table.
7. Run **Debug > Compile VBAProject**, then save and reopen if appropriate. Ensure VBA editor error trapping is **Break on Unhandled Errors**.
8. Run `SyncRackRegistry` once from **Alt+F8** to register rack IDs that were present before worksheet events were installed; then run `ValidateRackRegistry` once to check duplicates and bad heights.
9. On **Rack Lookup**, install the **eight ActiveX CommandButtons** and their Click handlers per [EIGHT_BUTTON_SETUP.md](EIGHT_BUTTON_SETUP.md) and [ACTIVEX_CONTROLS_SETUP.md](ACTIVEX_CONTROLS_SETUP.md). Do not use Assign Macro. Preserve the existing `Worksheet_Change` logic and any unrelated workbook events.

## Eight-button workflows

The latest additions implement **Edit Server**, **Decom Server**, **Audit Racks**, **Print Search** and **Print Racks**, alongside the established Add Server / Clear Search / Refresh buttons.

- [EIGHT_BUTTON_SETUP.md](EIGHT_BUTTON_SETUP.md) is the authoritative installation guide for all eight buttons, including the new standard modules `modServerActions`, `modRackAudit`, `modRackPrinting`, the `frmEditServer`, `frmDecomServer`, `frmRackAudit` forms, and the archive/audit table layouts.
- The new `RackLookup_SheetCode.txt` contains all eight worksheet Click events. Do not merge in duplicate `Worksheet_Change` procedures.
- **Decom Server** archives the complete inventory row in a `tblDecom` table on the `Decommissioned Servers` sheet before deleting the active row. An incompatible existing archive blocks the action.
- **Audit Racks** collects physical serial scans and creates a separate report with physical misses, unexpected scans and inventory-data warnings. It does not edit inventory.
- Both Print buttons open Print Preview rather than sending unattended print jobs.

## ActiveX compatibility

Microsoft 365 and Office 2024 disable ActiveX controls by default. Your organization must explicitly permit their use; if the buttons are blocked, **do not change managed security settings**. The old Form Control button macros remain in `buttonClear`, `buttonRefresh`, and `buttonAddServer` for an approved fallback. See [ACTIVEX_CONTROLS_SETUP.md](ACTIVEX_CONTROLS_SETUP.md).

## Workbook layout — exact names

| Worksheet | Excel table | Required content |
| --- | --- | --- |
| **Server Inventory** | `tblInventory` | `Rack ID`, `U Position`, `Hostname`, `Serial Number`, `Asset Tag`; `Height (U)` optional for the elevation renderer, **required for the Add Server form** |
| **Rack Info** | `tblRacks` | `Rack ID`, `Barcode`, `Rack Height` |
| **Rack Lookup** | none required | C3 Rack ID, C4 U position, C6 Hostname, C8 Serial Number, C10 Asset Tag; E2:I2 results headers, E3 spill formula, compact title/status in K2:M3, column headers K4:M4, rack elevation K5:M46 |

The `Height (U)` field describes **equipment height**, while `Rack Height` describes the **rack height**. Rack Height is blank for a newly registered rack, which means **42U** by default. The rack height must be an integer from **1 through 42**, never over 42.

The Brady printer can generate barcode artwork; the workbook does not generate barcodes. Rack IDs and barcode fields can be maintained independently.

## Changes and refreshes

- Scanning or editing a search input clears conflicting inputs then redraws the elevation; a pasted Rack ID + U pair is preserved. When multiple searches are pasted simultaneously, precedence is hostname → serial → asset → rack/U.
- Editing or pasting rows inside `tblInventory` redraws the current elevation; rack ID changes also synchronize new rack IDs into `tblRacks`.
- Editing `tblRacks[Rack ID]` or `tblRacks[Rack Height]` validates the rack registry and redraws the elevation.
- Rack synchronization **appends only**. It never deletes racks or modifies existing IDs, heights, or barcodes. The sync is case-insensitive and trims whitespace to avoid duplicate registrations.
- A mistaken new Rack ID in inventory is still treated as a new rack. Correcting the typo does **not** delete the incorrect registry row: review and remove it manually after confirming it is unused.
- Changing a formula result, importing data, or deleting an entire table can behave differently from typing inside a table; `Worksheet_Change` is not a general recalculation monitor. Use `SyncRackRegistry` / `RenderRackElevation` manually if needed.
- The **ActiveX Add Server** button opens `frmAddServer`, validates the selected registered rack and the full occupied U range, then appends one row to `tblInventory`. On success it displays and highlights the new device using C3 (rack) and C4 (bottom U). It preserves blank hostname/serial/asset values and does not modify other inventory columns. See [ADD_SERVER_FORM_SETUP.md](ADD_SERVER_FORM_SETUP.md).
- The **ActiveX Clear** button calls `buttonClear.ClearRackLookupButton`, which reuses `modRackSearch.ClearRackSearch`: it clears only C3, C4, C6, C8, C10, retains the E3 spill formula, redraws the empty rack view, and selects C3. The **ActiveX Refresh** button calls `buttonRefresh.RefreshRackLookupButton`, which reuses `modRackElevation.RenderRackElevation` without changing search inputs. Manual registry audit: `ValidateRackRegistry`.

## Rendering / validation rules

- Visualization uses a two-row, filled title/status area in **Rack Lookup!K2:M3** (dark title, pale-blue summary), column headings in **K4:M4**, and a 42U rack in **K5:M46**. Shorter racks use fewer rows; code clears K2:M126 on every redraw, including previous merged blocks. Do not store other data in that reserved area.
- `U Position` is the **bottom (lowest) U** a server occupies. Equipment fills upward: a 2U server at U40 occupies U40–U41, while a 4U server at U28 occupies U28–U31. `40` and `U40` work equivalently. Rack+U search matches the **bottom U** stored in inventory and highlights the entire device block.
- Equipment `Height (U)` may be omitted or left blank; then it is treated as 1U. A nonblank invalid height is counted and **not** drawn, rather than silently defaulted to 1U.
- Equipment spanning multiple U positions is drawn as one continuous block: hostname and serial cells are vertically merged over the occupied units, while column K retains individual U labels. A matched device highlights across its full height. The display unmerges old blocks before each redraw (including when switching to a shorter rack).
- A thick black border frames the entire visible panel from K2 through the last rack-unit row (K46 for 42U; K28 for 24U). A black double-line divider separates the column headings in row 4 from the first U-position row 5. Internal device outlines and U-position grid lines remain separate.
- Overlapping equipment U slots are red, populated slots blue, matched slots yellow, and empty slots gray. Conflicting rows remain unmerged so overlapping devices cannot be mistaken for a single device. Invalid U positions or equipment extending out of a rack are skipped and counted.
- Duplicate Rack IDs in `tblRacks` are reported by `ValidateRackRegistry` and marked in the drawing when a duplicated rack is selected. For backward compatibility the elevation uses the first matching registry record until duplicates are corrected.
- Blank Rack Height uses 42U. Invalid or missing rack metadata produces a warning in the drawing; invalid heights are not silently treated as verified. Missing `tblRacks` / required columns produce an explicit error rather than masking the broken structure.

## Resource safety

- Before iterating, the code checks that `tblInventory` has no more than **50,000** data rows and `tblRacks` has no more than **10,000**.
- **Do not resize a Table to an entire sheet/column.** Use **Table Design > Resize Table** to include only its actual records. Excel Tables expand automatically as you add real rows.
- An oversized Table causes a descriptive error instead of a long scan or out-of-memory condition.
- The code preserves and restores `Application.EnableEvents` and screen updating settings during modifications. Avoid interrupting execution without restoring those settings if Excel is paused in the debugger.

## Compatibility and limitations

- Requires desktop Microsoft Excel supporting **LET / FILTER / CHOOSE** dynamic arrays and VBA; the selected ActiveX worksheet controls also require that ActiveX be permitted by local or organization policy. The registry uses late-bound `Scripting.Dictionary` on Windows.
- Excel lookup text comparison is case-insensitive. VBA uses trimmed case-insensitive text keys. Store serial numbers / Rack IDs as **Text** if leading zeros are significant.
- Number strings for U should use the standard form `40` or `U40`. Rack metadata and equipment height must be whole numbers 1–42.
- This project is source code/documentation, **not a precompiled, tested .xlsm**. The UserForm must be built manually in desktop Excel and compiled there. Complete the verification in [TEST_PLAN.md](TEST_PLAN.md) on a backup workbook before deployment.
- Keep actual organization inventory/hostnames/serials out of this public code repository.
