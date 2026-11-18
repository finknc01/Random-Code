# RackScan: eight ActiveX buttons and three workflows

These files are **source code** for Excel desktop, not a preconfigured `.xlsm`.
Install on a **backup workbook** first. Do not commit real server inventory or
identifiers to GitHub.

## 1. Tables and button names

Existing workbook sheets/tables:

| Sheet | Table | Required headings |
| --- | --- | --- |
| Rack Lookup | none | Search cells C3/C4/C6/C8/C10, dynamic results E3:I, elevation K2:M46 |
| Server Inventory | `tblInventory` | Rack ID, U Position, Height (U), Hostname, Serial Number, Asset Tag |
| Rack Info | `tblRacks` | Rack ID, Barcode, Rack Height |

`U Position` is always the **bottom occupied U**. A 2U server at U40 occupies
U40 and U41. Rack Height blank defaults to 42U; maximum is 42U.

Create these **ActiveX CommandButtons** on `Rack Lookup` (Developer > Design
Mode > Insert > ActiveX CommandButton). Set `(Name)` exactly, not just Caption:

| Caption | `(Name)` | Standard module entry point |
| --- | --- | --- |
| Add Server | `cmdAddServer` | `buttonAddServer.OpenAddServerForm` |
| Edit Server | `cmdEditServer` | `buttonEditServer.OpenEditServerForm` |
| Decom Server | `cmdDecomServer` | `buttonDecomServer.OpenDecomServerForm` |
| Clear Search | `cmdClear` | `buttonClear.ClearRackLookupButton` |
| Refresh | `cmdRefresh` | `buttonRefresh.RefreshRackLookupButton` |
| Audit Racks | `cmdAuditRacks` | `buttonAuditRacks.OpenAuditRacksForm` |
| Print Search | `cmdPrintSearch` | `buttonPrintSearch.PrintSearchButton` |
| Print Racks | `cmdPrintRacks` | `buttonPrintRacks.PrintRacksButton` |

Paste `RackLookup_SheetCode.txt` into the **Rack Lookup sheet object**. It
contains the existing search `Worksheet_Change` and all eight Click events.
There must be exactly one `Worksheet_Change` and one handler per `(Name)`.
Preserve unrelated worksheet events from your working workbook.

## 2. Standard modules

Keep the existing `modRackData`, `modRackSearch`, `modRackElevation`,
`modRackRegistry`, `modServerEntry`, `buttonAddServer`, `buttonClear`, and
`buttonRefresh` modules.

Import these new modules under the same exact names as the filenames:

- `modServerActions.bas` - find one unambiguous server, edit it, and archive it.
- `modRackAudit.bas` - reconcile scans and perform rack data-quality checks.
- `modRackPrinting.bas` - print previews, restoring previous PageSetup settings.
- `buttonEditServer.bas`
- `buttonDecomServer.bas`
- `buttonAuditRacks.bas`
- `buttonPrintSearch.bas`
- `buttonPrintRacks.bas`

In the VBA editor's Properties window, confirm each standard module's
`(Name)` after importing. If Excel assigned a default name like `Module1`,
rename it to match the name shown above; the sheet code calls these modules
by name.

Existing `Add Server`, `Clear Search` and `Refresh` functionality is unchanged.
Buttons call reusable standard modules. They do not use Assign Macro.

## 3. Create frmEditServer

Insert a new UserForm and name it `frmEditServer`. Set its Caption to
`Edit Server | RackScan`. Add the controls below with exact `(Name)` values.
The label captions are suggestions; control names are mandatory.

| Type | `(Name)` | Caption / function |
| --- | --- | --- |
| Label | `lblTarget` | Identifies the row being edited |
| ComboBox | `cboRack` | Rack ID dropdown |
| TextBox | `txtUPosition` | Bottom U |
| TextBox | `txtHeightU` | Height (U) |
| TextBox | `txtHostname` | Hostname |
| TextBox | `txtSerial` | Serial Number |
| TextBox | `txtAssetTag` | Asset Tag |
| Label | `lblPlacement` | Current occupied U preview |
| Label | `lblMessage` | Validation / error message |
| CommandButton | `cmdSave` | Save Changes |
| CommandButton | `cmdCancel` | Cancel |

Copy `frmEditServer_Code.txt` to the form's code-behind. Include normal
labels beside the inputs, as with `frmAddServer`. Search for an existing
server on Rack Lookup before clicking **Edit Server**.

Search lookup priority is **hostname, serial, asset tag, then rack + bottom U**.
A search returning multiple inventory rows is rejected rather than guessing.
The form loads the matched row, allows all six fields to be edited, validates
the entire occupied U range (excluding its own existing row), and refreshes
the rack display on success. Other inventory columns remain intact.

## 4. Create frmDecomServer

Insert a UserForm named `frmDecomServer` with:

| Type | `(Name)` | Caption / function |
| --- | --- | --- |
| Label | `lblServer` | Large multiline server identification |
| TextBox | `txtReason` | Optional reason (max 255 chars) |
| Label | `lblMessage` | Error status |
| CommandButton | `cmdConfirm` | Decommission |
| CommandButton | `cmdCancel` | Cancel |

Copy `frmDecomServer_Code.txt` into its code-behind. The button requires a
single, unambiguous search result. The UserForm shows its identifying fields,
asks for confirmation, and **archives the entire source row before removing
the original from active inventory**.

On first successful use, if no `Decommissioned Servers` worksheet exists,
RackScan creates the sheet with a table named `tblDecom`. Its columns are
**all current `tblInventory` columns** plus `Decom Date` and `Decom Reason`.
If the sheet already exists, it **must** contain `tblDecom` with every
inventory heading and both metadata headings. An incompatible existing
sheet or table causes a clear error without moving any server data. Do not
rename or remove columns from an existing archive. Archived records are not
automatically restored or overwritten.

This intentionally uses a full snapshot rather than just the six display
fields, preserving any additional inventory columns in your workbook.

## 5. Create frmRackAudit

Insert a UserForm named `frmRackAudit` with:

| Type | `(Name)` | Caption / function |
| --- | --- | --- |
| ComboBox | `cboRack` | Rack ID dropdown |
| TextBox | `txtScan` | Serial number scanner input |
| ListBox | `lstScans` | Serial numbers already scanned |
| Label | `lblMessage` | Scan count and messages |
| CommandButton | `cmdAddScan` | Add Scan |
| CommandButton | `cmdFinish` | Finish Audit |
| CommandButton | `cmdCancel` | Cancel |

Copy `frmRackAudit_Code.txt` into the form's code-behind. Use a scanner that
sends **Enter** after each serial, or click **Add Scan**. Duplicates are
ignored; changing the rack resets scanned values. **Finish Audit** produces
`Rack Audit Report` showing:

- expected serials that were scanned (VERIFIED), not scanned (MISSING), or
  have blank serials (NO SERIAL / UNVERIFIABLE);
- unexpected scanned serials;
- inventory problems: invalid U/height, U overlap, rack-size overflow, and
  duplicate serials within the audited rack.

The report is **read-only relative to inventory**: it never changes server
records. The `Rack Audit Report` worksheet is generated/replaced on every
audit. To protect unrelated worksheets, an existing sheet of that name is
only overwritten when cell J1 contains `RACKSCAN_GENERATED_AUDIT`.

## 6. Print buttons

- **Print Search** previews the current E2:I search-results table, with
  current search criteria in a printed footer. If there is no printable
  result, it reports that instead.
- **Print Racks** refreshes the elevation and previews K2:M[bottom unit row]
  on a single portrait page. For a 42U rack this is K2:M46; for a 24U rack
  it is K2:M28. It refuses to print if no rack elevation is displayed.

Both commands show **Print Preview** and let the operator choose whether
to send the job to the printer. The original sheet's print area and other
changed PageSetup properties are restored after the preview closes.

## 7. Verify in Excel before production

1. Back up the `.xlsm`. Import all modules and create the three UserForms.
2. Confirm the eight ActiveX control names, paste the sheet code, and run
   **Debug > Compile VBAProject**. Turn Design Mode **off** for testing.
3. Use fictitious rack/server records only. Check Add/Clear/Refresh regression.
4. Edit a 2U server at bottom U40: U40–U41 occupied. Move it to a free rack;
   verify the original rack is no longer occupied. Try an overlap and a
   rack-height overflow: neither change should be saved.
5. Create two otherwise identical fictitious matches. Edit and Decom must
   refuse to pick a random record. A unique serial search should work.
6. Decommission one fictitious server. Verify all columns survive in
   `tblDecom` with the date/reason, and the source row is removed. An
   intentionally incompatible pre-existing archive should block the action.
7. Audit with one expected scan missing, one unexpected scan, a blank serial,
   and a bad U placement. Check report counts and flags; inventory unchanged.
8. Preview both print buttons, cancel the dialogs, then verify the pre-existing
   PageSetup and E3 lookup formula were not modified.
9. Confirm `Application.EnableEvents` is restored after an error. Save,
   close, and reopen the backup workbook and re-test all eight buttons.

**Compatibility:** ActiveX CommandButtons can be disabled by Office or
organization policy. Use only controls your managed Excel environment allows;
do not change organizational security settings to bypass restrictions.
