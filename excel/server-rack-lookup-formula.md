# Server rack inventory lookup formula

Original full `LET` + `CHOOSE` + `FILTER` formula for the **Rack Lookup** worksheet. Kept in its full form rather than replaced by the shorter alternative.

## Workbook layout

- **Server Inventory** has an Excel Table named `tblInventory`.
- **Rack Lookup** search inputs: `C3` = Rack ID; `C4` = U Position (with Rack ID); `C6` = Hostname; `C8` = Serial Number; `C10` = Asset Tag.
- **Rack Lookup** output headers: `E2` = Rack ID; `F2` = U Position; `G2` = Hostname; `H2` = Serial Number; `I2` = Asset Tag.
- Put the formula in **`Rack Lookup!E3`**. Keep the spill area `E3:I` clear.
- The separate worksheet-change VBA event clears conflicting search inputs; this formula retrieves rows and formats blank inventory values.

## Formula — Rack Lookup!E3

```excel
=LET(
    rack,TRIM($C$3&""),
    pos,TRIM($C$4&""),
    host,TRIM($C$6&""),
    serial,TRIM($C$8&""),
    asset,TRIM($C$10&""),

    data,CHOOSE({1,2,3,4,5},
        tblInventory[Rack ID],
        tblInventory[U Position],
        tblInventory[Hostname],
        tblInventory[Serial Number],
        tblInventory[Asset Tag]
    ),

    matches,IF(host<>"",
        (tblInventory[Hostname]&"")=host,
        IF(serial<>"",
            (tblInventory[Serial Number]&"")=serial,
            IF(asset<>"",
                (tblInventory[Asset Tag]&"")=asset,
                ((tblInventory[Rack ID]&"")=rack)*
                IF(pos<>"",
                    (tblInventory[U Position]&"")=pos,
                    1
                )
            )
        )
    ),

    IF(AND(rack="",host="",serial="",asset=""),
        IF(pos<>"","Enter a Rack ID in C3",""),
        FILTER(
            IF(data="","",data),
            matches,
            "No matching servers"
        )
    )
)
```

## Expected behavior

- `C3` only: all matching servers in the rack.
- `C3` + `C4`: only servers at that U position in that rack.
- `C6` only: search by hostname.
- `C8` only: search by serial number.
- `C10` only: search by asset tag.
- No search input: blank output; U Position without Rack ID: prompt for Rack ID; no matches: `No matching servers`.
- Missing hostname, serial number, or asset tag **in inventory data** is displayed as a blank rather than zero.
- Because the VBA event enforces exclusive search methods, this formula does not independently flag simultaneous searches. Without that event, it prioritizes hostname, then serial number, then asset tag, then rack.
