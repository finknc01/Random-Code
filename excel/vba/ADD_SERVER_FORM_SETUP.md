# Add Server — UserForm setup (desktop Excel)

This workflow adds a server from **Rack Lookup**, without typing directly into `tblInventory`. It uses a VBA **UserForm** opened by the **ActiveX Add Server CommandButton** on Rack Lookup. The UserForm itself contains standard Microsoft Forms (MSForms) controls.

## Before you start

1. **Back up your working `.xlsm` workbook.** Use fictitious test records in the copy.
2. Confirm these existing objects work: worksheets `Rack Lookup`, `Server Inventory`, `Rack Info`; tables `tblInventory` and `tblRacks`; shared module `modRackData`; and rack renderer `modRackElevation`.
3. Add a **Height (U)** column to `tblInventory` if it is missing, with the heading spelled exactly `Height (U)`. The existing rack display can read a missing/blank equipment height as 1U, but **Add Server requires this column** and writes a numeric height for each new record.
4. Ensure `tblRacks` contains the rack you want to select. Blank `Rack Height` means 42U. A nonblank height must be 1–42. The form intentionally does **not** create new racks.

## 1. Add the standard modules

In Excel press **Alt+F11**. In the Project Explorer, right-click your workbook, then choose **Insert > Module**:

| Standard module `(Name)` | Paste the contents of |
| --- | --- |
| `modServerEntry` | [modServerEntry.bas](modServerEntry.bas) |
| `buttonAddServer` | [buttonAddServer.bas](buttonAddServer.bas) |

For an existing module, open it and replace its contents only if you intend to update that particular module. **Do not paste UserForm event code into a standard module.**

The `OpenAddServerForm` routine in `buttonAddServer` is called from the worksheet's ActiveX `cmdAddServer_Click` handler. `modServerEntry` is responsible for validation, collision detection, append/rollback, and refreshing the Rack Lookup.

## 2. Build the UserForm

1. In the Visual Basic Editor, select your workbook in Project Explorer.
2. Choose **Insert > UserForm**.
3. In the **Properties** window (F4), set `(Name)` to **frmAddServer**; set `Caption` to **Add Server | RackScan**.
4. Make the form about **350 wide × 335 high** in the UserForm designer. Set `StartUpPosition` to **1 - CenterOwner**.
5. In the Toolbox (**View > Toolbox** if hidden), add the controls below. Set each control's **`(Name)`** exactly as shown, in the Properties window; change the visible `Caption` on labels and buttons as shown.

| Type | `(Name)` | Visible label or setting | Suggested position |
| --- | --- | --- | --- |
| Label | `lblRack` | Rack ID | x=18, y=16 |
| ComboBox | `cboRack` | Dropdown; list loaded from `tblRacks` | x=132, y=13, w=180 |
| Label | `lblU` | Bottom U Position | x=18, y=47 |
| TextBox | `txtUPosition` | Leave blank initially | x=132, y=44, w=80 |
| Label | `lblHeight` | Height (U) | x=18, y=78 |
| TextBox | `txtHeightU` | Starts at 1 | x=132, y=75, w=80 |
| Label | `lblHost` | Hostname (optional) | x=18, y=109 |
| TextBox | `txtHostname` | Blank allowed | x=132, y=106, w=180 |
| Label | `lblSerial` | Serial Number (optional) | x=18, y=140 |
| TextBox | `txtSerial` | Blank allowed | x=132, y=137, w=180 |
| Label | `lblAsset` | Asset Tag (optional) | x=18, y=171 |
| TextBox | `txtAssetTag` | Blank allowed | x=132, y=168, w=180 |
| Label | `lblPlacement` | Live occupied-U preview | x=18, y=205, w=305, h=30 |
| Label | `lblMessage` | Validation or error message | x=18, y=239, w=305, h=42 |
| CommandButton | `cmdSave` | **Add Server** | x=132, y=288, w=95 |
| CommandButton | `cmdCancel` | **Cancel** | x=238, y=288, w=75 |

To make it feel polished, try `lblPlacement.BackColor` = a pale blue and `lblMessage.ForeColor` = a dark red; keep enough text height for errors. These colors are cosmetic and do not affect validation. Set `TabIndex` in this order: rack, bottom U, height, hostname, serial, asset, Add Server, Cancel. The code sets the default Enter/Esc buttons.

**Naming matters:** `frmAddServer` is the UserForm's `(Name)`, not the visible `Caption`. The ComboBox must be `cboRack`, not `ComboBox1`, or compilation will fail.

## 3. Paste the form's code-behind

1. Select **frmAddServer** in Project Explorer.
2. Right-click > **View Code** (or press F7).
3. Paste the entire contents of [frmAddServer_Code.txt](frmAddServer_Code.txt) into that UserForm's code window. Do **not** put it in `ThisWorkbook`, a worksheet object, or a standard module.
4. Use **Debug > Compile VBAProject**. If it stops at an unknown control name, check the `(Name)` column in the table above.

The form fills its Rack ID dropdown from existing `tblRacks`. When possible it preselects the rack shown on Rack Lookup, including a rack found from a hostname/serial/asset search. It does **not** automatically select a U Position from an existing server: each new placement must be chosen explicitly.

## 4. Add the ActiveX button to Rack Lookup

Use **Developer > Insert > ActiveX Controls > Command Button** on Rack Lookup and set **`(Name) = cmdAddServer`**, **`Caption = Add Server`** while in Design Mode. Its click event must live in the **Rack Lookup worksheet object**, together with `cmdClear_Click` and `cmdRefresh_Click`. The new Click handler calls `buttonAddServer.OpenAddServerForm`.

See [ACTIVEX_CONTROLS_SETUP.md](ACTIVEX_CONTROLS_SETUP.md) for installing and styling all three ActiveX buttons, preserving the existing Worksheet_Change event, turning off Design Mode, and testing. ActiveX controls must be permitted by your organization's Office configuration.

## Behavior and safety

- **Bottom U:** a 2U server entered at U40 occupies **U40 and U41**, never U39–U40.
- **Rack range:** a 4U device at U40 is rejected in a 42U rack because it would require U43.
- **Collision check:** cannot add a device to any U already occupied in that rack. Adjacent devices are allowed.
- **Existing invalid positions:** if another record in that rack has an invalid bottom U/height or extends past the rack height, the addition is blocked until that record is corrected. This avoids pretending an unknown placement is empty.
- **Optional values:** Hostname, Serial Number, and Asset Tag may each be blank; if **all three** are blank, the form asks for confirmation. Text identifiers are stored as text, including leading zeros.
- **On save:** appends exactly one `tblInventory` row, updates Rack Lookup C3 (rack) and C4 (bottom U), clears conflicting search inputs, recalculates the E3 results, redraws the rack, and highlights the new device. The form leaves other inventory columns untouched.
- **On cancel or validation failure:** no inventory row is added. A write-time failure removes its newly created row.
- **Not part of this version:** editing an existing server, deleting a server, creating racks, or recording network connections. Those can be added as separate workflows later.

If ActiveX controls or UserForms are restricted by your organization's policy, stop and use the approved controls rather than changing Office security settings. Never upload real hostnames, serial numbers, or rack inventory to a public repository.

Finally, walk through the checks in [TEST_PLAN.md](TEST_PLAN.md) on your backup workbook.
