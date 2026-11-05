Option Explicit

' Reusable search-input behavior for the Rack Lookup worksheet.
' Inventory retrieval stays in the Excel formula in E3.
Public Sub HandleSearchChange(ByVal ws As Worksheet, ByVal Target As Range)

    If Target.CountLarge <> 1 Then Exit Sub
    If Intersect(Target, ws.Range("C3,C4,C6,C8,C10")) Is Nothing Then Exit Sub
    If IsError(Target.Value2) Then Exit Sub
    If Len(Trim$(CStr(Target.Value2))) = 0 Then Exit Sub

    On Error GoTo HandleError
    Application.EnableEvents = False

    Select Case Target.Address(False, False)
        Case "C3"                 ' New rack scan/search
            ws.Range("C4,C6,C8,C10").ClearContents
        Case "C4"                 ' U position keeps rack ID
            ws.Range("C6,C8,C10").ClearContents
        Case "C6"                 ' Hostname
            ws.Range("C3:C4,C8,C10").ClearContents
        Case "C8"                 ' Serial number
            ws.Range("C3:C4,C6,C10").ClearContents
        Case "C10"                ' Asset tag
            ws.Range("C3:C4,C6,C8").ClearContents
    End Select

CleanExit:
    Application.EnableEvents = True
    Exit Sub

HandleError:
    MsgBox "Search input error: " & Err.Description, vbExclamation
    Resume CleanExit

End Sub

' Optional: assign this macro to a Clear Search button.
Public Sub ClearRackSearch()
    Dim ws As Worksheet

    On Error GoTo HandleError
    Set ws = ThisWorkbook.Worksheets("Rack Lookup")
    Application.EnableEvents = False
    ws.Range("C3,C4,C6,C8,C10").ClearContents
    Application.EnableEvents = True

    ws.Activate
    ws.Range("C3").Select
    modRackElevation.RenderRackElevation
    Exit Sub

HandleError:
    Application.EnableEvents = True
    MsgBox "Could not clear search: " & Err.Description, vbExclamation
End Sub
