# Random-Code

This repository currently contains **RackScan**, an Excel/VBA tool for server
inventory, rack lookup, and visual rack elevations.

## RackScan

- [Project overview and workbook layout](excel/vba/README.md)
- [Install all eight ActiveX buttons](excel/vba/EIGHT_BUTTON_SETUP.md)
- [Add Server UserForm setup](excel/vba/ADD_SERVER_FORM_SETUP.md)
- [VBA modules and worksheet code](excel/vba/)
- [Excel lookup formula](excel/server-rack-lookup-formula.md)
- [Regression checklist](excel/vba/TEST_PLAN.md)
- [Repository audit and remaining risks](excel/vba/REPOSITORY_AUDIT.md)

The repository stores **VBA source and documentation**, not a working `.xlsm`
workbook. Create the documented Excel tables, forms, and controls in desktop
Excel, and test in a backup workbook before using them with real equipment.

**Privacy:** Do not commit production hostnames, asset tags, serial numbers,
network information, or other internal inventory to this repository.
