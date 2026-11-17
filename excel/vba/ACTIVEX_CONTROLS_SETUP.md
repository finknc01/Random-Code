# Rack Lookup — ActiveX button setup

RackScan uses **ActiveX CommandButtons on the Rack Lookup worksheet** for Clear, Refresh, and Add Server. The Add Server data-entry window is still `frmAddServer`, a VBA UserForm that already uses Microsoft Forms controls.

> **Compatibility gate:** Microsoft 365 and Office 2024 disable ActiveX controls by default. Your organization's configuration must permit the controls. If you cannot insert them or they remain disabled, **stop and ask IT**; do not change Trust Center settings or security policy to force them on. Keep the existing Form Control buttons as a working fallback.

## 1. Back up and confirm modules

1. Save a copy of the working `.xlsm` first.
2. Make sure the standard modules still exist:
   - `modRackSearch` (owns clearing the search).
   - `modRackElevation` (owns rendering the rack).
   - `modServerEntry` (owns Add Server validation).
   - `buttonClear`, `buttonRefresh`, and `buttonAddServer` (small entry-point wrappers).
3. For Add Server, the `frmAddServer` UserForm and all its controls must also exist. See [ADD_SERVER_FORM_SETUP.md](ADD_SERVER_FORM_SETUP.md).

## 2. Add the three worksheet buttons

In desktop Excel on **Rack Lookup**:

1. Choose **Developer > Design Mode**.
2. Go to **Developer > Insert > ActiveX Controls > Command Button**.
3. Draw the first button. While still in Design Mode, right-click it > **Properties**.
4. Change the **`(Name)` property**, then `Caption`. Repeat for the other two.

| Worksheet control | Exact `(Name)` | `Caption` | Existing VBA routine |
| --- | --- | --- | --- |
| Clear button | `cmdClear` | Clear Search | `buttonClear.ClearRackLookupButton` |
| Refresh button | `cmdRefresh` | Refresh Rack | `buttonRefresh.RefreshRackLookupButton` |
| Add Server button | `cmdAddServer` | Add Server | `buttonAddServer.OpenAddServerForm` |

These are **worksheet controls**, not additional VBA modules. Their `(Name)` is different from their `Caption`. You **do not** use Assign Macro with ActiveX buttons.

You cannot convert a Form Control button into an ActiveX control in place. Add the three new controls, verify that they work, and only then manually remove the specific old Form Control buttons if you want to.

## 3. Make them look consistent

In **Design Mode > Properties**, consider:

| Property | Clear | Refresh | Add Server |
| --- | --- | --- | --- |
| `BackColor` | Muted gray | Dark blue | Dark green |
| `ForeColor` | White | White | White |
| `Font` | Segoe UI, 10 pt, Bold | Same | Same |
| `TakeFocusOnClick` | False | False | False |
| `TabStop` | False | False | False |
| `SpecialEffect` | Raised/default | Raised/default | Raised/default |

All three can use the same width and height, with matching gaps. `TakeFocusOnClick = False` and `TabStop = False` help keep keyboard navigation practical for a USB/Bluetooth scanner workflow. ActiveX CommandButtons provide native button-state visual feedback when supported by the installed Excel/Windows configuration.

## 4. Put the Click procedures in the Rack Lookup sheet object

1. Press **Alt+F11**.
2. In **Microsoft Excel Objects**, double-click the sheet that is actually named **Rack Lookup**.
3. Replace only its old `Worksheet_Change` with the updated version in [RackLookup_SheetCode.txt](RackLookup_SheetCode.txt), and add all three `cmd..._Click` procedures from that file to the **same sheet object's code window**.
4. Remove any empty, auto-generated `CommandButton1_Click` stubs or outdated duplicate `cmdClear_Click` handlers in that sheet object. Do not remove unrelated sheet procedures.
5. Do **not** paste the Click procedures into `buttonClear`, `buttonRefresh`, `buttonAddServer`, `ThisWorkbook`, or `frmAddServer`. All three Click handlers belong directly to **Sheet (Rack Lookup)**.
6. Use **Debug > Compile VBAProject**.

The sheet must contain **exactly one** `Worksheet_Change` routine and exactly one Click routine for each of the three named controls.

## 5. Test the new controls

1. Return to Excel and switch **Developer > Design Mode OFF**.
2. Type a fictitious Rack ID into **C3**. Use **Refresh Rack**; the elevation should redraw without changing search values.
3. Use **Clear Search**; C3, C4, C6, C8, C10 should clear, E3's formula should remain, and focus should return to C3.
4. Use **Add Server**; `frmAddServer` should open. Cancel it and confirm the table is unchanged.
5. Enter a fictitious 2U server at bottom U40 in a test rack, save, and check that the rack display highlights U40–U41.
6. Close and reopen the backup workbook to check control initialization and behavior. If the workbook reports blocked ActiveX, stop here rather than changing your organization's settings.

Old standard modules `buttonClear.bas`, `buttonRefresh.bas`, and `buttonAddServer.bas` are **retained**. They are the routines the new ActiveX Click handlers call and make a Form Control fallback straightforward if ActiveX is later disabled.
