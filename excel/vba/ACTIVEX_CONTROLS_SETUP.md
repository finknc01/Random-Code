# RackScan — eight ActiveX CommandButtons

RackScan's current configuration has **eight ActiveX buttons** on the **Rack Lookup** sheet. See [EIGHT_BUTTON_SETUP.md](EIGHT_BUTTON_SETUP.md) for their exact control names, standard module entry points, three UserForms, archive schema, audit scanning workflow, printing behavior and full tests.

| Caption | ActiveX `(Name)` |
| --- | --- |
| Add Server | `cmdAddServer` |
| Edit Server | `cmdEditServer` |
| Decom Server | `cmdDecomServer` |
| Clear Search | `cmdClear` |
| Refresh | `cmdRefresh` |
| Audit Racks | `cmdAuditRacks` |
| Print Search | `cmdPrintSearch` |
| Print Racks | `cmdPrintRacks` |

Insert these with **Developer > Design Mode > Insert > ActiveX CommandButton**. Set the `(Name)` property, not just the caption. All Click events go into the **Rack Lookup worksheet object**, as shown in [RackLookup_SheetCode.txt](RackLookup_SheetCode.txt). ActiveX controls do not use Assign Macro.

To give them consistent styling, set `Font` to Segoe UI 10 Bold, `TakeFocusOnClick=False`, `TabStop=False`, and use a contrasting `ForeColor`. The Edit Server button can use muted yellow `BackColor=&H009FE0F5&` (BGR) with `ForeColor=&H00124D69&`. Choose darker accessible colors for other actions as appropriate.

**Compatibility:** Microsoft 365 and Office 2024 may disable ActiveX controls by default. Use only controls permitted by your organization's security policy, and do not change managed security settings to bypass it. Back up the workbook and confirm **Debug > Compile VBAProject** succeeds before replacing controls in the live workbook.
