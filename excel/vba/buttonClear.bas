Option Explicit

' Assigned to the Clear button on Rack Lookup.
' The shared search module owns the clearing and automatic redraw logic.
Public Sub ClearRackLookupButton()
    modRackSearch.ClearRackSearch
End Sub
