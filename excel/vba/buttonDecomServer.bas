Option Explicit

Public Sub OpenDecomServerForm()
    Dim item As ListRow, issue As String
    On Error GoTo HandleError
    Set item = modServerActions.FindServerFromLookup(issue)
    If item Is Nothing Then
        MsgBox issue, vbExclamation, "Decom Server"
        Exit Sub
    End If
    Load frmDecomServer
    frmDecomServer.LoadServer item
    frmDecomServer.Show
    Exit Sub
HandleError:
    MsgBox "Cannot open Decom Server: " & Err.Description, vbExclamation
End Sub
