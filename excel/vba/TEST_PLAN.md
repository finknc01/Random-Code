# Rack workbook regression test plan

**Eight-button release:** Also follow [EIGHT_BUTTON_SETUP.md](EIGHT_BUTTON_SETUP.md), section 7, for Edit Server, Decom Server, physical Rack Audit, Print Search, and Print Racks. The original Add/Clear/Refresh checks below remain relevant.

Run on a **backup copy** of the Excel workbook after installing all five standard modules, all three worksheet events, and the updated E3 formula. All data used below should be fictitious; keep real inventory out of the public repository.

## Build and setup

- [ ] **Debug > Compile VBAProject** succeeds with no missing module, variable, or duplicate event procedure.
- [ ] All three sheets are named `Rack Lookup`, `Server Inventory`, `Rack Info`. Tables are named `tblInventory` and `tblRacks`.
- [ ] VBA error trapping is **Break on Unhandled Errors**.
- [ ] Excel tables contain only real records, not a whole worksheet.
- [ ] Run `SyncRackRegistry` once with Alt+F8; then run `ValidateRackRegistry`.

## ActiveX controls and compatibility

- [ ] Confirm the organization's Excel configuration permits ActiveX controls. If controls are disabled, **stop** and follow the approved IT policy; do not change Trust Center settings.
- [ ] Add exactly three ActiveX CommandButtons on Rack Lookup with names `cmdClear`, `cmdRefresh`, `cmdAddServer`. Design Mode must be OFF when testing clicks.
- [ ] `RackLookup_SheetCode.txt` handlers are pasted into **Sheet (Rack Lookup)**, not a standard module; exactly one `Worksheet_Change` remains.
- [ ] Buttons have consistent font/size, distinct BackColor, readable ForeColor, and visible press feedback; `TakeFocusOnClick` and `TabStop` are False.
- [ ] After testing the new controls, old Form Control buttons can be removed manually; the old `button*.bas` entry points remain intact.
- [ ] Close and reopen the backup workbook; all three controls still initialize and operate normally.

## Search and display

- [ ] Scan a registered 42U rack into C3. K2:M2 is a filled title bar, K3:M3 is a filled summary bar, K4:M4 contains the headers, K5 is U42 and K46 is U1.
- [ ] Thick black borders surround the rendered panel K2:M46, including the title/status/header, and a black double border runs across K:M exactly between rows 4 and 5.
- [ ] Search a shorter 24U rack. K5 is U24, K28 is U1, and no older rack slots or thick borders remain below row 28.
- [ ] C3 alone displays all rows for the rack; C3+C4 highlights a device whose **bottom U** matches C4.
- [ ] C4 = `U40` and C4 = `40` produce equivalent results for a server stored with either convention.
- [ ] C6 hostname, C8 serial, C10 asset each locate the right rack; other search inputs clear.
- [ ] Search a non-existent identifier: the summary shows the message in the filled K3:M3 bar and no stale rack remains in the drawing.
- [ ] Paste C3:C4 (both nonblank) at once; both remain. Paste multiple nonblank search fields; priority is Hostname > Serial > Asset > Rack.
- [ ] Delete a search input: remaining non-conflicting search inputs are not accidentally cleared.
- [ ] Create ActiveX control **`cmdClear`** on Rack Lookup with a `cmdClear_Click` handler calling `buttonClear.ClearRackLookupButton`. Click it after populating search fields; C3, C4, C6, C8, and C10 become blank, the E3 formula remains in place, the elevation clears, and focus returns to C3.
- [ ] Create ActiveX control **`cmdRefresh`** on Rack Lookup with a `cmdRefresh_Click` handler calling `buttonRefresh.RefreshRackLookupButton`. Click it with an active rack or server search; the rack redraws without clearing or changing any search values.
- [ ] A server with a blank hostname, serial number, or asset tag still draws; search formula shows blanks, not zeros.
- [ ] Rack-height, invalid-placement, and no-match warnings remain visible in the compact summary bar, without using K4 or K5 for messages.
- [ ] Search an identifier associated with equipment in two different racks: elevation shows a multiple-rack warning instead of choosing arbitrarily.

## Add Server form workflow

- [ ] Add a `Height (U)` column to `tblInventory`; add `modServerEntry`, `buttonAddServer`, and `frmAddServer` with all documented control names. Compile successfully.
- [ ] Create ActiveX control **`cmdAddServer`** with the `cmdAddServer_Click` handler in the Rack Lookup sheet object calling `buttonAddServer.OpenAddServerForm`. Button opens `frmAddServer` and Cancel leaves inventory unchanged.
- [ ] With rack ID in C3, open Add Server: the existing registered rack is preselected, but bottom U is blank. With a unique hostname/serial/asset search, the form preselects the matching rack.
- [ ] Save a fictitious 2U server at **bottom U40** in a 42U rack: the new `tblInventory` row has U Position = 40 and Height (U) = 2; the drawn block fills U40–U41 and highlights both.
- [ ] Save a 1U server immediately adjacent to an existing 2U server (e.g., bottom U39 next to U40–U41): allowed. Attempt U41 with the 2U in place: rejected; no row added.
- [ ] A 4U device at bottom U40 in a 42U rack is rejected as out of range. A device exceeding a shorter rack's height is also rejected.
- [ ] An invalid rack height, unregistered Rack ID, or duplicate Rack ID in `tblRacks` prevents saving.
- [ ] Existing records with invalid U Position, invalid nonblank Height (U), or a placement exceeding rack capacity block new additions **to that rack** until corrected.
- [ ] Blank hostname, serial, or asset tag individually is stored as a true blank. When all three are blank, the form asks confirmation; No cancels, Yes permits.
- [ ] A serial or asset tag with leading zeros remains text, including after save and lookup refresh.
- [ ] Successful save updates Rack Lookup C3/C4, clears C6/C8/C10, retains E3 lookup formula, refreshes elevation, and highlights the full block.
- [ ] Existing inventory rows and unrelated table columns remain intact. Save an invalid U, nonnumeric height, or overlapping placement and confirm row count is unchanged.
- [ ] Test on a backup workbook with an empty table, a full-height rack, and a shorter rack. Verify the form never silently creates an unregistered rack.

## Equipment positions / validation

- [ ] A 2U server with bottom U40 occupies U40 and U41 as one visually merged block in L:M; K still shows separate U41 and U40 markers. C4 = U40 highlights both slots; C4 = U41 does not match that server. The hostname and serial appear once, not once per U.
- [ ] A 4U server with bottom U28 spans U28–U31 in one continuous block; search U28 (or its hostname/serial) highlights all four U positions.
- [ ] Change a 4U device to 1U, or switch between a 42U and 24U rack: old merges and borders disappear; no stale blocks remain.
- [ ] Two adjacent 1U or multi-U devices remain distinct blocks and do not merge together.
- [ ] When a conflict intersects a multi-U device, the conflicted U is red and unmerged; any remaining uninterrupted portion of the device is still labeled.
- [ ] Two devices claiming the same U show red conflict rows.
- [ ] Invalid equipment U (0, 43, text) is counted and not drawn.
- [ ] Nonblank invalid equipment Height (U) (0, 43, text, 1.5) is counted and not drawn; a **blank** height draws 1U.
- [ ] A device whose height extends above the rack's last U (e.g., a 4U device with bottom U40 in a 42U rack) is counted but not drawn.
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
