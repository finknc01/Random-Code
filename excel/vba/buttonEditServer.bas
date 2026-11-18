Option Explicit

Public Sub OpenEditServerForm()
    Dim item As ListRow, issue As String
    On Error GoTo HandleError
    Set item = modServerActions.FindServerFromLookup(issue)
    If item Is Nothing Then
        MsgBox issue, vbExclamation, "Edit Server"
        Exit Sub
    End If
    Load frmEditServer
    frmEditServer.LoadServer item
    frmEditServer.Show
    Exit Sub
HandleError:
    MsgBox "Cannot open Edit Server: " & Err.Description, vbExclamation
End Sub
