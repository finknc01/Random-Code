# VBA setup for server rack lookup and elevation

These modules work with the existing full Excel formula in [server-rack-lookup-formula.md](../server-rack-lookup-formula.md).

## Install

1. Save the Excel workbook as **.xlsm**.
2. Open Visual Basic Editor using **Alt+F11**.
3. Under **Insert > Module**, create a standard module named **modRackSearch** and put the contents of [modRackSearch.bas](modRackSearch.bas) inside it. If one is already installed, replace its contents with the latest version.
4. Create a second standard module named **modRackElevation** and put the contents of [modRackElevation.bas](modRackElevation.bas) inside it.
5. Under **Microsoft Excel Objects**, double-click the **Rack Lookup** worksheet. Replace the *existing* `Worksheet_Change` procedure with the content of [RackLookup_SheetCode.txt](RackLookup_SheetCode.txt). Do not put this worksheet event in a standard module. There must be only one `Worksheet_Change` event in this worksheet.
6. Run **Debug > Compile VBAProject**, then save.

## Workbook layout

- **Server Inventory** contains the Excel Table `tblInventory` with headers: `Rack ID`, `U Position`, `Hostname`, `Serial Number`, `Asset Tag`.
- **Rack Lookup** uses search inputs `C3` (rack), `C4` (U), `C6` (hostname), `C8` (serial), and `C10` (asset tag).
- The existing dynamic formula remains in **Rack Lookup!E3** and outputs into columns E through I; the VBA does not overwrite it.
- The graphical rack elevation is drawn using formatted worksheet cells in **Rack Lookup!K2:M66**, which should be reserved for this feature.
- **Rack Labels** is unchanged; the Brady printer can make the barcodes.

## Graphical rack behavior

- Rack ID **only** in C3: displays every server recorded in the rack, with no selection highlighted.
- Rack ID C3 **plus** U Position C4: displays the whole rack, highlighting the equipment at that U.
- Hostname C6, serial number C8, or asset tag C10: finds the matching inventory row, determines its rack, displays that entire rack, and highlights every matching device within that rack.
- If the search identifier matches servers in **more than one rack**, the elevation warns rather than arbitrarily choosing a rack. The normal E3 Excel formula may still return multiple result rows.
- Missing asset tags and hostnames do not prevent other servers from appearing. An entry with an invalid or out-of-range U position is not drawn and is counted as a warning.
- Blue rows are occupied, gray rows are empty, yellow rows are search matches, and red rows indicate more than one equipment record claiming the same U.

## Unit height and rack size

- The initial rack height is **42U**, controlled by the `RACK_UNITS = 42` constant near the top of `modRackElevation`. Change it to `48` for a 48U rack.
- The `U Position` field is treated as the **topmost** U occupied by the device. U positions may be numbers (e.g. `40`) or text with a `U` prefix (e.g. `U40`), but a consistent format is recommended.
- The code optionally recognizes an additional **`Height (U)`** column in `tblInventory`. If it is missing or empty, devices are shown as 1U. A 2U server at U40 occupies U40 and U39.
- **For accurate rack-space use, add and populate `Height (U)`** for multi-U equipment. Without it, the display is a device-position map, not a reliable free-U calculation.
- The current version assumes front/back equipment orientations are not tracked; overlapping records are flagged instead.

## Refreshing

Changing any of the five search inputs triggers `modRackElevation.RenderRackElevation` automatically. Editing **Server Inventory** alone does not automatically redraw the elevation; run `RenderRackElevation` from **Alt+F8** or make a Refresh button if you want an immediate update.

The optional `ClearRackSearch` macro from `modRackSearch` now also clears the visual and reselects C3. Assign it to a **Developer > Insert > Button (Form Control)** on Rack Lookup if desired.

## Scope

VBA manages clearing conflicting search fields and generating the graphical rack view. The formula in E3 still handles search results, and no inventory records are modified by these macros.
