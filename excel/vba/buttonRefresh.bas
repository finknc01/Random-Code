Option Explicit

' Called by Rack Lookup's ActiveX cmdRefresh_Click event.
' Redraws without changing search fields; also usable manually.
Public Sub RefreshRackLookupButton()
    modRackElevation.RenderRackElevation
End Sub
