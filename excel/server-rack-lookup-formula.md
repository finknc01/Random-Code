# Server rack inventory lookup formula

Full `LET` + `CHOOSE` + `FILTER` formula for **Rack Lookup!E3**, retaining the original five-column layout. This version trims inventory text and treats `40` and `U40` as the same top-of-device U position. It follows the same search priority as the rack elevation.

## Layout

- **Server Inventory** contains Excel Table `tblInventory` with `Rack ID`, `U Position`, `Hostname`, `Serial Number`, `Asset Tag`. Other columns may exist.
- **Rack Lookup** inputs: C3 Rack ID, C4 U Position (with rack), C6 Hostname, C8 Serial Number, C10 Asset Tag.
- Output headers E2:I2: Rack ID, U Position, Hostname, Serial Number, Asset Tag.
- Put the formula in **E3**, outside a table. Leave its spill destination E3:I clear.
- A blank search shows blank results. U alone prompts for a rack; unmatched searches say `No matching servers`.

## Formula — Rack Lookup!E3

```excel
=LET(
    rack,TRIM($C$3&""),
    pos,TRIM($C$4&""),
    host,TRIM($C$6&""),
    serial,TRIM($C$8&""),
    asset,TRIM($C$10&""),

    uText,TRIM(tblInventory[U Position]&""),
    uNumber,IFERROR(
        --IF(LEFT(UPPER(uText),1)="U",MID(uText,2,99),uText),
        -1
    ),
    posNumber,IFERROR(
        --IF(LEFT(UPPER(pos),1)="U",MID(pos,2,99),pos),
        -2
    ),

    data,CHOOSE({1,2,3,4,5},
        tblInventory[Rack ID],
        tblInventory[U Position],
        tblInventory[Hostname],
        tblInventory[Serial Number],
        tblInventory[Asset Tag]
    ),

    matches,IF(host<>"",
        TRIM(tblInventory[Hostname]&"")=host,
        IF(serial<>"",
            TRIM(tblInventory[Serial Number]&"")=serial,
            IF(asset<>"",
                TRIM(tblInventory[Asset Tag]&"")=asset,
                (TRIM(tblInventory[Rack ID]&"")=rack)*
                IF(pos<>"",
                    (uNumber=posNumber)*
                    (uNumber=INT(uNumber))*
                    (posNumber=INT(posNumber))*
                    (posNumber>=1)*
                    (posNumber<=42),
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

## Search behavior

- C3 only: all equipment assigned to that rack.
- C3 + C4: matching equipment whose **topmost occupied U** equals C4; `U40` and `40` are equivalent.
- C6, C8, or C10: match hostname, serial, or asset respectively; case-insensitive Excel text equality.
- Search precedence if several inputs remain populated: hostname, serial number, asset tag, rack/U.
- Leading/trailing spaces in inventory search fields are ignored.
- Missing hostname, serial number, and asset values in inventory remain visually blank.
- For multi-U equipment, C4 matches the **topmost U**, not every occupied position. The graphical elevation shows each occupied U.
- Rack heights can be at most **42U**. Equipment with invalid heights or positions may appear in the table search results but are excluded (with warning) from physical rack rendering.

## Validation notes

This is a full spill formula, **not** a formula for `tblRacks` or any other Excel Table. Verify that the formula compiles in your Excel version and that E3:I has no blocking values. The formula uses numeric conversion for U searches; stick to canonical whole-number U positions (`40` or `U40`) rather than decimal or exponential text formatting. Rack ID and serial number cells should be stored as Text where leading zeroes matter.
