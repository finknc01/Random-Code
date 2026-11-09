# Rack workbook regression test plan

Run on a **backup copy** of the Excel workbook after installing all four standard modules, all three worksheet events, and the updated E3 formula. All data used below should be fictitious; keep real inventory out of the public repository.

## Build and setup

- [ ] **Debug > Compile VBAProject** succeeds with no missing module, variable, or duplicate event procedure.
- [ ] All three sheets are named `Rack Lookup`, `Server Inventory`, `Rack Info`. Tables are named `tblInventory` and `tblRacks`.
- [ ] VBA error trapping is **Break on Unhandled Errors**.
- [ ] Excel tables contain only real records, not a whole worksheet.
- [ ] Run `SyncRackRegistry` once with Alt+F8; then run `ValidateRackRegistry`.

## Search and display

- [ ] Scan a registered 42U rack into C3. K7 is U42 and K48 is U1.
- [ ] Search a shorter 24U rack. K7 is U24, K30 is U1, and no older rack slots remain below.
- [ ] C3 alone displays all rows for the rack; C3+C4 highlights a device at that **top U**.
- [ ] C4 = `U40` and C4 = `40` produce equivalent results for a server stored with either convention.
- [ ] C6 hostname, C8 serial, C10 asset each locate the right rack; other search inputs clear.
- [ ] Search a non-existent identifier: no stale rack remains in the drawing.
- [ ] Paste C3:C4 (both nonblank) at once; both remain. Paste multiple nonblank search fields; priority is Hostname > Serial > Asset > Rack.
- [ ] Delete a search input: remaining non-conflicting search inputs are not accidentally cleared.
- [ ] A server with a blank hostname, serial number, or asset tag still draws; search formula shows blanks, not zeros.
- [ ] Search an identifier associated with equipment in two different racks: elevation shows a multiple-rack warning instead of choosing arbitrarily.

## Equipment positions / validation

- [ ] A 2U server with top U40 occupies U40 and U39. Search U40 highlights it.
- [ ] Two devices claiming the same U show red conflict rows.
- [ ] Invalid equipment U (0, 43, text) is counted and not drawn.
- [ ] Nonblank invalid equipment Height (U) (0, 43, text, 1.5) is counted and not drawn; a **blank** height draws 1U.
- [ ] A device extending below U1 is counted but not drawn.
- [ ] Rack Height 42 is valid. Rack Height blank defaults to 42 with a note.
- [ ] Rack Height 43, 0, nonnumeric, fractional values trigger registry validation and drawing warning; no rack renders above 42U.
- [ ] Duplicate Rack IDs differing only by case/whitespace are reported by `ValidateRackRegistry`.
- [ ] Editing a rack height automatically redraws the currently selected rack.

## Registration and automatic refresh

- [ ] Add fictitious new rack ID `DEMO-42` to `tblInventory`. It is appended once to `tblRacks` with blank Rack Height.
- [ ] Add another server with the ID ` demo-42 ` (different case/spaces). No duplicate registry row is created.
- [ ] Existing rack heights, barcodes, and Rack IDs remain unchanged after repeated syncs.
- [ ] Paste multiple inventory records with different new rack IDs. Each unique ID is appended once.
- [ ] Change an inventory U position while its rack is displayed. Diagram redraws without editing C3 again.
- [ ] Delete a server record or correct a Rack ID. Rack IDs already in the registry are not automatically removed.
- [ ] Change Rack Info's Rack Height. Display updates immediately.

## Resource safety

- [ ] In a disposable workbook only, enlarge `tblRacks` beyond 10,000 rows. Running registry validation or elevation reports an oversized-table error before scanning rows.
- [ ] Enlarge `tblInventory` beyond 50,000 rows in a disposable workbook. Elevation and registry sync stop with a descriptive error.
- [ ] After any handled error, check Immediate Window: `? Application.EnableEvents` should return `True` under normal operation.
- [ ] If VBA is manually stopped while events are disabled, explicitly restore `Application.EnableEvents = True` and `Application.ScreenUpdating = True` in the Immediate Window.

## Known limits

- The E3 lookup formula uses numeric conversion for C4; use canonical values such as 40 or U40. Exotic numeric expressions are not supported by the VBA position parser.
- Worksheet_Change does not trigger on formula recalculation or every external data refresh. Use manual `SyncRackRegistry` / `RenderRackElevation` after those operations.
- Manual Excel runtime execution is required; source review alone cannot verify compatibility with the user's installed Excel version.
