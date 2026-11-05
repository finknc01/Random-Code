# VBA setup for server rack lookup

This directory contains the VBA code used with the full Excel lookup formula in [server-rack-lookup-formula.md](../server-rack-lookup-formula.md).

## Files and installation

1. Save your workbook as **.xlsm** (Excel macro-enabled workbook).
2. Open the Visual Basic Editor with **Alt+F11**.
3. Under **Insert > Module**, add a standard module and name it **modRackSearch** in the Properties window (F4). Enter the code from [modRackSearch.bas](modRackSearch.bas) into that module. You can also try importing the .bas file if permitted.
4. In Project Explorer, double-click the **Rack Lookup** worksheet object (not ThisWorkbook and not a standard module). Enter the event code from [RackLookup_SheetCode.txt](RackLookup_SheetCode.txt) there.
5. If you already created a `Worksheet_Change` event for Rack Lookup, **replace** that event; do not keep two copies. Only one `Worksheet_Change` procedure should exist in that worksheet module.
6. Select **Debug > Compile VBAProject**, save, and reopen with macros enabled if permitted by your organization.
7. Keep the existing Excel `LET` / `FILTER` formula in `Rack Lookup!E3`. The VBA does not populate result cells.

## Search inputs

| Cell | Searches for | Fields cleared on new entry |
| --- | --- | --- |
| C3 | Rack ID | C4, C6, C8, C10 |
| C4 | U Position (with C3 Rack ID) | C6, C8, C10 |
| C6 | Hostname | C3, C4, C8, C10 |
| C8 | Serial Number | C3, C4, C6, C10 |
| C10 | Asset Tag | C3, C4, C6, C8 |

Clearing an input cell does not trigger another clearing action. Blank inventory values (including missing asset tags) are handled by the Excel formula. The event responds to edits of one input cell at a time; it deliberately ignores multi-cell paste operations.

## Optional Clear Search button

Use **Developer > Insert > Button (Form Control)** on the Rack Lookup sheet and assign the macro `ClearRackSearch`. It clears C3/C4/C6/C8/C10 and selects C3.

## Scanner behavior

A USB/Bluetooth keyboard-emulation barcode scanner should enter the rack ID into C3. The scanner's code must match the inventory Rack ID exactly. If its Enter suffix advances Excel to another cell, change **File > Options > Advanced > After pressing Enter, move selection** or adjust the scanner configuration.

## Scope

These VBA files manage **input fields only**. They do not generate barcodes (Brady prints those), change master inventory records, or replace the existing formula results in E3:I. Other future features can be added in separate standard modules.
