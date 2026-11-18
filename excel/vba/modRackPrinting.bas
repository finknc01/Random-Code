Option Explicit

' Both buttons open Print Preview. They do not send an unattended print job.
' All temporary PageSetup settings are restored when preview closes.

Public Sub PrintSearchResults()
    Dim ws As Worksheet, spill As Range
    Dim lastRow As Long, printRange As String
    Dim detail As String, raw As String, firstResult As String
    Dim previousArea As String, oldHeader As String, oldFooter As String
    Dim oldOrientation As Long, oldZoom As Variant
    Dim oldWide As Variant, oldTall As Variant, oldCenter As Boolean
    Dim settingsChanged As Boolean, r As Long

    On Error GoTo HandleError
    Set ws = ThisWorkbook.Worksheets("Rack Lookup")
    ws.Range("E3").Calculate
    If IsError(ws.Range("E3").Value2) Then
        MsgBox "Search has an error or no printable results.", vbExclamation, "Print Search"
        Exit Sub
    End If

    On Error Resume Next
    Set spill = ws.Range("E3").SpillingToRange
    On Error GoTo HandleError
    If Not spill Is Nothing Then
        lastRow = spill.Row + spill.Rows.Count - 1
    Else
        lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    End If
    firstResult = modRackData.TextValue(ws.Range("E3").Value2)
    If lastRow < 3 Or Len(firstResult) = 0 Or _
       StrComp(firstResult, "No matching servers", vbTextCompare) = 0 Or _
       StrComp(firstResult, "Enter a Rack ID in C3", vbTextCompare) = 0 Then
        MsgBox "There are no search results to print.", vbInformation, "Print Search"
        Exit Sub
    End If

    detail = ""
    For r = 1 To 5
        Select Case r
            Case 1: raw = modRackData.TextValue(ws.Range("C3").Value2): If Len(raw) > 0 Then detail = detail & " Rack=" & raw
            Case 2: raw = modRackData.TextValue(ws.Range("C4").Value2): If Len(raw) > 0 Then detail = detail & " BottomU=" & raw
            Case 3: raw = modRackData.TextValue(ws.Range("C6").Value2): If Len(raw) > 0 Then detail = detail & " Host=" & raw
            Case 4: raw = modRackData.TextValue(ws.Range("C8").Value2): If Len(raw) > 0 Then detail = detail & " Serial=" & raw
            Case 5: raw = modRackData.TextValue(ws.Range("C10").Value2): If Len(raw) > 0 Then detail = detail & " Asset=" & raw
        End Select
    Next r
    printRange = ws.Range("E2:I" & lastRow).Address

    With ws.PageSetup
        previousArea = .PrintArea
        oldHeader = .CenterHeader
        oldFooter = .CenterFooter
        oldOrientation = .Orientation
        oldZoom = .Zoom
        oldWide = .FitToPagesWide
        oldTall = .FitToPagesTall
        oldCenter = .CenterHorizontally
        settingsChanged = True
        .PrintArea = printRange
        .CenterHeader = "RackScan - Search Results"
        .CenterFooter = Left$(Replace(Trim$(detail), "&", "&&"), 230)
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
    End With
    ws.PrintPreview

CleanExit:
    If settingsChanged Then
        On Error Resume Next
        With ws.PageSetup
            .PrintArea = previousArea
            .CenterHeader = oldHeader
            .CenterFooter = oldFooter
            .Orientation = oldOrientation
            .Zoom = oldZoom
            .FitToPagesWide = oldWide
            .FitToPagesTall = oldTall
            .CenterHorizontally = oldCenter
        End With
        On Error GoTo 0
    End If
    Exit Sub
HandleError:
    MsgBox "Print Search error (" & Err.Number & "): " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Public Sub PrintRackElevation()
    Dim ws As Worksheet
    Dim lastUnitRow As Long, r As Long
    Dim value As String, previousArea As String, oldHeader As String
    Dim oldOrientation As Long, oldZoom As Variant
    Dim oldWide As Variant, oldTall As Variant, oldCenter As Boolean
    Dim settingsChanged As Boolean

    On Error GoTo HandleError
    Set ws = ThisWorkbook.Worksheets("Rack Lookup")
    modRackElevation.RenderRackElevation
    For r = 5 To 46
        value = modRackData.TextValue(ws.Cells(r, "K").Value2)
        If Left$(UCase$(value), 1) = "U" Then
            If modRackData.ParseU(value) > 0 Then lastUnitRow = r
        End If
    Next r
    If lastUnitRow < 5 Then
        MsgBox "Search for a rack before printing its elevation.", vbInformation, "Print Racks"
        Exit Sub
    End If

    With ws.PageSetup
        previousArea = .PrintArea
        oldHeader = .CenterHeader
        oldOrientation = .Orientation
        oldZoom = .Zoom
        oldWide = .FitToPagesWide
        oldTall = .FitToPagesTall
        oldCenter = .CenterHorizontally
        settingsChanged = True
        .PrintArea = ws.Range("K2:M" & lastUnitRow).Address
        .CenterHeader = "RackScan - Rack Elevation"
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
    End With
    ws.PrintPreview

CleanExit:
    If settingsChanged Then
        On Error Resume Next
        With ws.PageSetup
            .PrintArea = previousArea
            .CenterHeader = oldHeader
            .Orientation = oldOrientation
            .Zoom = oldZoom
            .FitToPagesWide = oldWide
            .FitToPagesTall = oldTall
            .CenterHorizontally = oldCenter
        End With
        On Error GoTo 0
    End If
    Exit Sub
HandleError:
    MsgBox "Print Racks error (" & Err.Number & "): " & Err.Description, vbExclamation
    Resume CleanExit
End Sub
