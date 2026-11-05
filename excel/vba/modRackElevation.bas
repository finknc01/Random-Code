Option Explicit

' Rack elevation on Rack Lookup!K2:M126 (up to 120U).
' Rack-specific height comes from tblRacks on the Rack Labels worksheet.
Private Const DEFAULT_RACK_UNITS As Long = 42
Private Const MAX_RACK_UNITS As Long = 120
Private Const FIRST_UNIT_ROW As Long = 7

Public Sub RenderRackElevation()
    Dim ws As Worksheet, tbl As ListObject, entry As ListRow
    Dim rackCol As Long, uCol As Long, hostCol As Long
    Dim serialCol As Long, assetCol As Long, heightCol As Long
    Dim requestedRack As String, selectedRack As String, uSearch As String
    Dim searchValue As String, searchCol As Long, candidate As String
    Dim label() As String, serials() As String
    Dim used() As Boolean, highlighted() As Boolean, conflicts() As Boolean
    Dim u As Long, topU As Long, bottomU As Long, unitHeight As Long
    Dim outputRow As Long, deviceCount As Long, skipped As Long, hitCount As Long
    Dim isMatch As Boolean, matchedAnywhere As Boolean, multipleRacks As Boolean
    Dim hostName As String, serialNo As String, savedEvents As Boolean
    Dim savedScreen As Boolean
    Dim rowCells As Range
    Dim rackUnits As Long, heightNote As String

    savedEvents = Application.EnableEvents
    savedScreen = Application.ScreenUpdating
    On Error GoTo HandleError
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Set ws = ThisWorkbook.Worksheets("Rack Lookup")
    Set tbl = ThisWorkbook.Worksheets("Server Inventory").ListObjects("tblInventory")

    rackCol = tbl.ListColumns("Rack ID").Index
    uCol = tbl.ListColumns("U Position").Index
    hostCol = tbl.ListColumns("Hostname").Index
    serialCol = tbl.ListColumns("Serial Number").Index
    assetCol = tbl.ListColumns("Asset Tag").Index
    heightCol = OptionalColumn(tbl, "Height (U)")

    ' Only use this reserved section of the Rack Lookup worksheet.
    With ws.Range("K2:M" & (FIRST_UNIT_ROW + MAX_RACK_UNITS - 1))
        .ClearContents
        .Interior.Color = RGB(255, 255, 255)
        .Font.Color = RGB(40, 48, 58)
        .Font.Bold = False
        .Borders.LineStyle = xlNone
    End With
    ws.Columns("K").ColumnWidth = 9
    ws.Columns("L").ColumnWidth = 28
    ws.Columns("M").ColumnWidth = 23
    ws.Range("K2").Value = "RACK ELEVATION"
    ws.Range("K2").Font.Bold = True

    requestedRack = CellText(ws.Range("C3"))
    uSearch = CellText(ws.Range("C4"))

    ' Same search precedence as the formula in E3.
    If Len(CellText(ws.Range("C6"))) > 0 Then
        searchCol = hostCol
        searchValue = CellText(ws.Range("C6"))
    ElseIf Len(CellText(ws.Range("C8"))) > 0 Then
        searchCol = serialCol
        searchValue = CellText(ws.Range("C8"))
    ElseIf Len(CellText(ws.Range("C10"))) > 0 Then
        searchCol = assetCol
        searchValue = CellText(ws.Range("C10"))
    Else
        selectedRack = requestedRack
    End If

    ' For non-rack searches, find the rack that contains the matching server.
    If searchCol > 0 Then
        For Each entry In tbl.ListRows
            If StrComp(RowText(entry, searchCol), searchValue, vbTextCompare) = 0 Then
                matchedAnywhere = True
                candidate = RowText(entry, rackCol)
                If Len(candidate) = 0 Then
                    ws.Range("K3").Value = "Matching server has no Rack ID."
                    GoTo CleanExit
                End If
                If Len(selectedRack) = 0 Then
                    selectedRack = candidate
                ElseIf StrComp(selectedRack, candidate, vbTextCompare) <> 0 Then
                    multipleRacks = True
                    Exit For
                End If
            End If
        Next entry

        If multipleRacks Then
            ws.Range("K3").Value = "Matches found in multiple racks."
            ws.Range("K4").Value = "Use a unique identifier or a rack search."
            GoTo CleanExit
        End If
        If Not matchedAnywhere Then
            ws.Range("K3").Value = "No matching server found."
            GoTo CleanExit
        End If
    End If

    If Len(selectedRack) = 0 Then
        If Len(uSearch) > 0 Then
            ws.Range("K3").Value = "Enter a Rack ID in C3."
        Else
            ws.Range("K3").Value = "Scan a rack or search for a server."
        End If
        GoTo CleanExit
    End If

    rackUnits = GetRackUnits(selectedRack, heightNote)

    ReDim label(1 To rackUnits)
    ReDim serials(1 To rackUnits)
    ReDim used(1 To rackUnits)
    ReDim highlighted(1 To rackUnits)
    ReDim conflicts(1 To rackUnits)

    For Each entry In tbl.ListRows
        If StrComp(RowText(entry, rackCol), selectedRack, vbTextCompare) = 0 Then
            deviceCount = deviceCount + 1
            hostName = RowText(entry, hostCol)
            serialNo = RowText(entry, serialCol)
            If Len(hostName) = 0 Then hostName = "(hostname blank)"

            topU = ReadU(RowText(entry, uCol))
            unitHeight = 1
            If heightCol > 0 Then
                If IsNumeric(RowText(entry, heightCol)) Then
                    unitHeight = CLng(Val(RowText(entry, heightCol)))
                End If
            End If
            bottomU = topU - unitHeight + 1

            isMatch = False
            If searchCol > 0 Then
                isMatch = (StrComp(RowText(entry, searchCol), searchValue, vbTextCompare) = 0)
            ElseIf Len(uSearch) > 0 Then
                isMatch = (StrComp(RowText(entry, uCol), uSearch, vbTextCompare) = 0)
            End If
            If isMatch Then hitCount = hitCount + 1

            ' Invalid/out-of-range U values must not be drawn inaccurately.
            If topU < 1 Or topU > rackUnits Or bottomU < 1 _
               Or unitHeight < 1 Then
                skipped = skipped + 1
            Else
                For u = topU To bottomU Step -1
                    If used(u) Then
                        conflicts(u) = True
                        label(u) = "U POSITION CONFLICT"
                        serials(u) = ""
                    Else
                        If u = topU Then
                            label(u) = hostName & IIf(unitHeight > 1, _
                                " (" & unitHeight & "U)", "")
                            serials(u) = serialNo
                        Else
                            label(u) = "| " & hostName
                            serials(u) = ""
                        End If
                    End If
                    used(u) = True
                    If isMatch Then highlighted(u) = True
                Next u
            End If
        End If
    Next entry

    ws.Range("K3").Value = selectedRack
    ws.Range("L3").Value = rackUnits & "U rack | " & deviceCount & " records"
    If Len(heightNote) > 0 Then ws.Range("K4").Value = heightNote
    If skipped > 0 Then
        ws.Range("M3").Value = skipped & " invalid/out-of-range U"
    ElseIf (searchCol > 0 Or Len(uSearch) > 0) And hitCount = 0 Then
        ws.Range("M3").Value = "No highlighted match"
    End If

    ws.Range("K6").Value = "U"
    ws.Range("L6").Value = "Hostname"
    ws.Range("M6").Value = "Serial Number"
    With ws.Range("K6:M6")
        .Font.Bold = True
        .Interior.Color = RGB(40, 54, 68)
        .Font.Color = RGB(255, 255, 255)
    End With

    For u = rackUnits To 1 Step -1
        outputRow = FIRST_UNIT_ROW + rackUnits - u
        Set rowCells = ws.Range("K" & outputRow & ":M" & outputRow)
        ws.Cells(outputRow, "K").Value = "U" & u
        If used(u) Then
            ws.Cells(outputRow, "L").Value = label(u)
            ws.Cells(outputRow, "M").Value = serials(u)
            rowCells.Interior.Color = RGB(220, 234, 246)
        Else
            rowCells.Interior.Color = RGB(246, 247, 249)
        End If
        If highlighted(u) Then
            rowCells.Interior.Color = RGB(255, 221, 120)
            rowCells.Font.Bold = True
        End If
        If conflicts(u) Then
            rowCells.Interior.Color = RGB(255, 173, 173)
            rowCells.Font.Bold = True
        End If
        rowCells.Borders(xlEdgeBottom).LineStyle = xlContinuous
        rowCells.Borders(xlEdgeBottom).Color = RGB(219, 223, 228)
    Next u

CleanExit:
    Application.EnableEvents = savedEvents
    Application.ScreenUpdating = savedScreen
    Exit Sub

HandleError:
    MsgBox "Rack elevation error: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

' Rack Labels is the rack-level registry: one row per rack.
' Blank/unlisted heights default to 42U. Invalid heights are flagged.
' The optional Height (U) column on Server Inventory is DIFFERENT:
' it describes how many U a single device occupies.
Private Function GetRackUnits(ByVal rackId As String, _
                              ByRef note As String) As Long
    Dim racks As ListObject
    Dim entry As ListRow
    Dim idCol As Long, heightCol As Long
    Dim rawHeight As String
    Dim heightValue As Double
    Dim found As Boolean

    GetRackUnits = DEFAULT_RACK_UNITS

    On Error Resume Next
    Set racks = ThisWorkbook.Worksheets("Rack Labels").ListObjects("tblRacks")
    On Error GoTo 0

    If racks Is Nothing Then
        note = "tblRacks not found - defaulting to 42U"
        Exit Function
    End If

    idCol = OptionalColumn(racks, "Rack ID")
    heightCol = OptionalColumn(racks, "Rack Height")
    If idCol = 0 Or heightCol = 0 Then
        note = "Check tblRacks headers (Rack ID, Rack Height)"
        Exit Function
    End If

    For Each entry In racks.ListRows
        If StrComp(RowText(entry, idCol), rackId, vbTextCompare) = 0 Then
            If found Then
                note = "Duplicate Rack ID in tblRacks; using first"
                Exit Function
            End If

            found = True
            rawHeight = RowText(entry, heightCol)

            If Len(rawHeight) = 0 Then
                note = "Rack Height blank - defaulting to 42U"
            ElseIf Not IsNumeric(rawHeight) Then
                note = "Non-numeric Rack Height - defaulting to 42U"
            Else
                heightValue = CDbl(rawHeight)
                If heightValue < 1 Or heightValue > MAX_RACK_UNITS _
                   Or heightValue <> Fix(heightValue) Then
                    note = "Invalid Rack Height - defaulting to 42U"
                Else
                    GetRackUnits = CLng(heightValue)
                End If
            End If
        End If
    Next entry

    If Not found Then note = "Rack missing from tblRacks - defaulting to 42U"
End Function

Private Function OptionalColumn(ByVal tbl As ListObject, _
                                ByVal heading As String) As Long
    On Error Resume Next
    OptionalColumn = tbl.ListColumns(heading).Index
    On Error GoTo 0
End Function

Private Function CellText(ByVal cell As Range) As String
    If IsError(cell.Value2) Then Exit Function
    CellText = Trim$(CStr(cell.Value2))
End Function

Private Function RowText(ByVal entry As ListRow, ByVal col As Long) As String
    Dim v As Variant
    v = entry.Range.Cells(1, col).Value2
    If Not IsError(v) Then RowText = Trim$(CStr(v))
End Function

Private Function ReadU(ByVal value As String) As Long
    Dim s As String
    s = UCase$(Trim$(value))
    If Left$(s, 1) = "U" Then s = Mid$(s, 2)
    If Not IsNumeric(s) Then Exit Function
    If Val(s) <> Fix(Val(s)) Then Exit Function
    ReadU = CLng(Val(s))
End Function
