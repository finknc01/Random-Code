Option Explicit

' Clears conflicting search inputs after a scan, edit, or multi-cell paste.
' Priority for multi-cell pastes: hostname > serial > asset > rack/U.
' A pasted rack + U pair is kept together.
Public Sub HandleSearchChange(ByVal ws As Worksheet, ByVal Target As Range)
    Dim mode As String
    Dim preserveU As Boolean
    Dim previousEvents As Boolean

    If Application.Intersect(Target, ws.Range("C3,C4,C6,C8,C10")) Is Nothing Then Exit Sub

    If EditedWithValue(ws, Target, "C6") Then
        mode = "C6"
    ElseIf EditedWithValue(ws, Target, "C8") Then
        mode = "C8"
    ElseIf EditedWithValue(ws, Target, "C10") Then
        mode = "C10"
    ElseIf EditedWithValue(ws, Target, "C3") Then
        mode = "C3"
        preserveU = (Target.CountLarge > 1) And _
                    EditedWithValue(ws, Target, "C4")
    ElseIf EditedWithValue(ws, Target, "C4") Then
        mode = "C4"
    Else
        ' Deleting an input should redraw the view, not erase other inputs.
        Exit Sub
    End If

    previousEvents = Application.EnableEvents
    On Error GoTo HandleError
    Application.EnableEvents = False

    Select Case mode
        Case "C3"
            If preserveU Then
                ws.Range("C6,C8,C10").ClearContents
            Else
                ws.Range("C4,C6,C8,C10").ClearContents
            End If
        Case "C4"
            ws.Range("C6,C8,C10").ClearContents
        Case "C6"
            ws.Range("C3:C4,C8,C10").ClearContents
        Case "C8"
            ws.Range("C3:C4,C6,C10").ClearContents
        Case "C10"
            ws.Range("C3:C4,C6,C8").ClearContents
    End Select

CleanExit:
    Application.EnableEvents = previousEvents
    Exit Sub

HandleError:
    MsgBox "Search input error (" & Err.Number & "): " & _
           Err.Description, vbExclamation
    Resume CleanExit
End Sub

Private Function EditedWithValue(ByVal ws As Worksheet, _
                                 ByVal Target As Range, _
                                 ByVal cellAddress As String) As Boolean
    Dim inputCell As Range
    Set inputCell = ws.Range(cellAddress)

    If Application.Intersect(Target, inputCell) Is Nothing Then Exit Function
    EditedWithValue = (Len(modRackData.TextValue(inputCell.Value2)) > 0)
End Function

' Optional: assign to a Clear Search button.
Public Sub ClearRackSearch()
    Dim ws As Worksheet
    Dim previousEvents As Boolean

    previousEvents = Application.EnableEvents
    On Error GoTo HandleError
    Set ws = ThisWorkbook.Worksheets("Rack Lookup")

    Application.EnableEvents = False
    ws.Range("C3,C4,C6,C8,C10").ClearContents
    Application.EnableEvents = previousEvents

    ws.Activate
    ws.Range("C3").Select
    modRackElevation.RenderRackElevation
    Exit Sub

HandleError:
    Application.EnableEvents = previousEvents
    MsgBox "Could not clear search (" & Err.Number & "): " & _
           Err.Description, vbExclamation
End Sub
