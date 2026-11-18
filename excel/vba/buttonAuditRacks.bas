Option Explicit

Public Sub OpenAuditRacksForm()
    On Error GoTo HandleError
    frmRackAudit.Show
    Exit Sub
HandleError:
    MsgBox "Cannot open Rack Audit: " & Err.Description, vbExclamation
End Sub
