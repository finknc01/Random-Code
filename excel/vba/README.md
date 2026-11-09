# VBA setup for server rack lookup, rack registry, and elevation

Works with the existing Excel lookup formula documented in [server-rack-lookup-formula.md](../server-rack-lookup-formula.md).

## Install or update

Save the workbook as **.xlsm**. Open the VBA editor with **Alt+F11**.

### Standard modules (under Modules)

- **modRackSearch** — import or paste [modRackSearch.bas](modRackSearch.bas). Handles search-field clearing.
- **modRackElevation** — replace the existing module code with [modRackElevation.bas](modRackElevation.bas). Draws the rack using the Rack Info worksheet.
- **modRackRegistry** — create a standard module with this exact name and paste [modRackRegistry.bas](modRackRegistry.bas). Adds missing rack IDs without altering heights already stored.

### Worksheet event code (under Microsoft Excel Objects)

Put each event in its own worksheet's code window, **not** in a standard module:

| Excel worksheet | Code file |
| --- | --- |
| Rack Lookup | [RackLookup_SheetCode.txt](RackLookup_SheetCode.txt) |
| Server Inventory | [ServerInventory_SheetCode.txt](ServerInventory_SheetCode.txt) |
| Rack Info | [RackInfo_SheetCode.txt](RackInfo_SheetCode.txt) |

Replace an existing `Worksheet_Change` procedure for the same worksheet rather than creating a duplicate. Do not overwrite other unrelated sheet events.

Select **Debug > Compile VBAProject** and save. If your workbook has existing racks not yet registered, run `SyncRackRegistry` once from **Alt+F8**.

## Workbook structure

- **Server Inventory** has `tblInventory` with at least: `Rack ID`, `U Position`, `Hostname`, `Serial Number`, `Asset Tag`. Optional: `Height (U)` (height of each piece of equipment).
- **Rack Info** has `tblRacks`, with `Rack ID`, `Barcode`, `Rack Height`. One row per rack. Rack IDs are ordinary table values, **not a spilled UNIQUE formula within the table**. Blank Rack Height means 42U.
- **Rack Lookup** inputs: C3 = Rack ID; C4 = U position; C6 = hostname; C8 = serial; C10 = asset. The formula in E3 spills results into E:I. The graphical rack elevation uses K2:M126.

## Automatic updates

- Changes to **Rack Lookup** search inputs clear conflicting fields and redraw the elevation.
- Changes anywhere inside `tblInventory` redraw the current elevation; changes specifically in `tblInventory[Rack ID]` also run `SyncRackRegistry`.
- Changes to `tblRacks[Rack ID]` or `tblRacks[Rack Height]` redraw the current elevation.
- Newly encountered rack IDs are appended to `tblRacks`. Existing rack rows, barcode information, and heights are not overwritten or removed. New Rack Height cells are blank, so they default to 42U in the display.
- The synchronization does **not** automatically repair typos in rack IDs or distinguish a valid new rack from an accidental ID. Review new rack IDs and consider input validation before broad deployment.
- **Do not resize the tables to entire worksheet columns**: an oversized `tblRacks` previously caused an Out of Memory error. The registry helper refuses tables above 10,000 rows as a safeguard.
- Excel's `Worksheet_Change` event responds to direct edits, not recalculations of formula-driven cells. If you later populate rack IDs using formulas, the event-driven registry will need a different refresh mechanism.

## Rack elevation details

- Rack height is looked up from `Rack Info!tblRacks` and can be 1–120U. Blank/invalid/unregistered heights fall back to 42U with a warning. (Correct invalid/unregistered values rather than treating that default as verified.)
- Equipment is placed by the **topmost U position** with optional `Height (U)` from `tblInventory`; without equipment height, the code assumes 1U.
- All equipment in a rack is shown; matching equipment is highlighted for a rack+U, hostname, serial, or asset search.
- Overlaps are marked red; missing or out-of-range positions are counted but not drawn.
- `RenderRackElevation` can be run from **Alt+F8** for manual refresh; `ClearRackSearch` clears the search and redraws the empty view.

## Cautions

Automatic rack registration writes to the workbook and may cause unintended new rack rows if the Rack ID is mistyped. Back up the workbook before enabling and consider data validation. The code is stored in GitHub but must still be compiled and tested inside desktop Excel.
