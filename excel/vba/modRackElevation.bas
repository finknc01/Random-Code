Option Explicit

Private Const FIRST_UNIT_ROW As Long = 7

' Draws the selected rack in Rack Lookup!K2:M48 (42U maximum).
' K2:M126 is cleared to remove remnants of older, taller displays.
Public Sub RenderRackElevation()
    Dim ws As Worksheet, inventory As ListObject, entry As ListRow
    Dim rackCol As Long, uCol As Long, hostCol As Long
    Dim serialCol As Long, assetCol As Long, heightCol As Long
    Dim selectedRack As String, candidate As String
    Dim uSearch As String, searchValue As String
    Dim searchCol As Long, requestedU As Long
    Dim labels() As String, serials() As String
    Dim occupied() As Boolean, highlighted() As Boolean, conflicts() As Boolean
    Dim slotOwner() As Long
    Dim u As Long, topU As Long, bottomU As Long, deviceHeight As Long
    Dim displayRow As Long, deviceCount As Long, skipped As Long, hitCount As Long
    Dim matchThisDevice As Boolean, anyMatch As Boolean, moreThanOneRack As Boolean
    Dim hostname As String, serialNumber As String, rawHeight As String
    Dim rackUnits As Long, heightNote As String
    Dim rowCells As Range
    Dim priorEvents As Boolean, priorScreen As Boolean

    priorEvents = Application.EnableEvents
    priorScreen = Application.ScreenUpdating
    On Error GoTo HandleError
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Set ws = ThisWorkbook.Worksheets("Rack Lookup")
    Set inventory = modRackData.GetTable("Server Inventory", "tblInventory")
    modRackData.CheckTableSize inventory, modRackData.MAX_INVENTORY_ROWS

    rackCol = modRackData.ColumnIndex(inventory, "Rack ID")
    uCol = modRackData.ColumnIndex(inventory, "U Position")
    hostCol = modRackData.ColumnIndex(inventory, "Hostname")
    serialCol = modRackData.ColumnIndex(inventory, "Serial Number")
    assetCol = modRackData.ColumnIndex(inventory, "Asset Tag")
    heightCol = modRackData.ColumnIndex(inventory, "Height (U)", False)

    ' Unmerge the device blocks from the previous drawing before clearing.
    ws.Range("L7:M126").UnMerge
    With ws.Range("K2:M126")
        .ClearContents
        .Interior.Pattern = xlNone
        .Font.Color = RGB(40, 48, 58)
        .Font.Bold = False
        .Borders.LineStyle = xlNone
    End With
    ws.Columns("K").ColumnWidth = 9
    ws.Columns("L").ColumnWidth = 28
    ws.Columns("M").ColumnWidth = 23
    ws.Range("K2").Value2 = "RACK ELEVATION"
    ws.Range("K2").Font.Bold = True

    selectedRack = modRackData.TextValue(ws.Range("C3").Value2)
    uSearch = modRackData.TextValue(ws.Range("C4").Value2)
    If Len(uSearch) > 0 Then requestedU = modRackData.ParseU(uSearch)

    ' Keep the same priority as Rack Lookup!E3:
    ' hostname > serial > asset tag > rack / U position.
    If Len(modRackData.TextValue(ws.Range("C6").Value2)) > 0 Then
        searchCol = hostCol
        searchValue = modRackData.TextValue(ws.Range("C6").Value2)
    ElseIf Len(modRackData.TextValue(ws.Range("C8").Value2)) > 0 Then
        searchCol = serialCol
        searchValue = modRackData.TextValue(ws.Range("C8").Value2)
    ElseIf Len(modRackData.TextValue(ws.Range("C10").Value2)) > 0 Then
        searchCol = assetCol
        searchValue = modRackData.TextValue(ws.Range("C10").Value2)
    End If

    If searchCol > 0 Then
        selectedRack = ""
        For Each entry In inventory.ListRows
            If StrComp(modRackData.RowText(entry, searchCol), _
                       searchValue, vbTextCompare) = 0 Then
                anyMatch = True
                candidate = modRackData.RowText(entry, rackCol)
                If Len(candidate) = 0 Then
                    ws.Range("K3").Value2 = "Matching server has no Rack ID."
                    GoTo CleanExit
                End If
                If Len(selectedRack) = 0 Then
                    selectedRack = candidate
                ElseIf StrComp(selectedRack, candidate, vbTextCompare) <> 0 Then
                    moreThanOneRack = True
                    Exit For
                End If
            End If
        Next entry

        If moreThanOneRack Then
            ws.Range("K3").Value2 = "Matches found in multiple racks."
            ws.Range("K4").Value2 = "Search a rack or use a unique device ID."
            GoTo CleanExit
        ElseIf Not anyMatch Then
            ws.Range("K3").Value2 = "No matching server found."
            GoTo CleanExit
        End If
    End If

    If Len(selectedRack) = 0 Then
        If Len(uSearch) > 0 Then
            ws.Range("K3").Value2 = "Enter a Rack ID in C3."
        Else
            ws.Range("K3").Value2 = "Scan a rack or search for a server."
        End If
        GoTo CleanExit
    End If

    rackUnits = GetRackUnits(selectedRack, heightNote)
    ReDim labels(1 To rackUnits)
    ReDim serials(1 To rackUnits)
    ReDim occupied(1 To rackUnits)
    ReDim highlighted(1 To rackUnits)
    ReDim conflicts(1 To rackUnits)
    ReDim slotOwner(1 To rackUnits)

    For Each entry In inventory.ListRows
        If StrComp(modRackData.RowText(entry, rackCol), _
                   selectedRack, vbTextCompare) = 0 Then
            deviceCount = deviceCount + 1
            hostname = modRackData.RowText(entry, hostCol)
            serialNumber = modRackData.RowText(entry, serialCol)
            If Len(hostname) = 0 Then hostname = "(hostname blank)"

            ' U Position identifies the bottom (lowest) rack unit of a server.
            bottomU = modRackData.ParseU(modRackData.RowText(entry, uCol))
            deviceHeight = 1
            If heightCol > 0 Then
                rawHeight = modRackData.RowText(entry, heightCol)
                If Len(rawHeight) > 0 Then
                    deviceHeight = modRackData.ParseU(rawHeight, False)
                End If
            End If

            ' Equipment extends upward from its listed bottom U.
            ' Invalid heights are never silently interpreted as 1U.
            topU = bottomU + deviceHeight - 1
            If bottomU < 1 Or bottomU > rackUnits Or _
               deviceHeight < 1 Or topU > rackUnits Then
                skipped = skipped + 1
            Else
                matchThisDevice = False
                If searchCol > 0 Then
                    matchThisDevice = (StrComp( _
                        modRackData.RowText(entry, searchCol), _
                        searchValue, vbTextCompare) = 0)
                ElseIf Len(uSearch) > 0 And requestedU > 0 Then
                    ' Search and E3 both match the BOTTOM position of a device.
                    matchThisDevice = (bottomU = requestedU)
                End If
                If matchThisDevice Then hitCount = hitCount + 1

                For u = topU To bottomU Step -1
                    If occupied(u) Then
                        ' Never merge overlapping devices into an ordinary block.
                        conflicts(u) = True
                        slotOwner(u) = 0
                        labels(u) = "U POSITION CONFLICT"
                        serials(u) = ""
                    Else
                        ' This record owns each of its U slots, not just its top U.
                        slotOwner(u) = deviceCount
                        serials(u) = serialNumber
                        If u = topU Then
                            labels(u) = hostname
                        Else
                            ' Used only when a conflict splits a device block.
                            labels(u) = hostname & " (continued)"
                        End If
                    End If

                    occupied(u) = True
                    If matchThisDevice Then highlighted(u) = True
                Next u
            End If
        End If
    Next entry

    ws.Range("K3").Value2 = selectedRack
    ws.Range("L3").Value2 = rackUnits & "U rack | " & deviceCount & " records"
    If Len(heightNote) > 0 Then ws.Range("K4").Value2 = heightNote

    If skipped > 0 Then
        ws.Range("M3").Value2 = skipped & " invalid U/height records"
    ElseIf Len(uSearch) > 0 And requestedU = 0 And searchCol = 0 Then
        ws.Range("M3").Value2 = "Invalid U search (1-42)"
    ElseIf (searchCol > 0 Or Len(uSearch) > 0) And hitCount = 0 Then
        ws.Range("M3").Value2 = "No highlighted match"
    End If

    ws.Range("K6").Value2 = "U"
    ws.Range("L6").Value2 = "Hostname"
    ws.Range("M6").Value2 = "Serial Number"
    With ws.Range("K6:M6")
        .Font.Bold = True
        .Interior.Color = RGB(40, 54, 68)
        .Font.Color = RGB(255, 255, 255)
    End With

    For u = rackUnits To 1 Step -1
        displayRow = FIRST_UNIT_ROW + rackUnits - u
        Set rowCells = ws.Range("K" & displayRow & ":M" & displayRow)
        ws.Cells(displayRow, "K").Value2 = "U" & u

        If occupied(u) Then
            rowCells.Interior.Color = RGB(220, 234, 246)
            If conflicts(u) Then
                ws.Cells(displayRow, "L").Value2 = labels(u)
            ElseIf u = rackUnits Then
                ws.Cells(displayRow, "L").Value2 = labels(u)
                ws.Cells(displayRow, "M").Value2 = serials(u)
            ElseIf slotOwner(u) <> slotOwner(u + 1) Then
                ' Show the label once per uninterrupted device block.
                ws.Cells(displayRow, "L").Value2 = labels(u)
                ws.Cells(displayRow, "M").Value2 = serials(u)
            End If
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

        ' Keep individual U markings while the server areas form one block.
        With ws.Cells(displayRow, "K").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(219, 223, 228)
        End With
        If slotOwner(u) = 0 Then
            With ws.Range("L" & displayRow & ":M" & displayRow).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Color = RGB(219, 223, 228)
            End With
        End If
    Next u

    DrawDeviceBlocks ws, slotOwner, rackUnits

CleanExit:
    Application.EnableEvents = priorEvents
    Application.ScreenUpdating = priorScreen
    Exit Sub

HandleError:
    MsgBox "Rack elevation error (" & Err.Number & "): " & _
           Err.Description, vbExclamation
    Resume CleanExit
End Sub

' Merge only the display cells belonging to the same uninterrupted device.
' U labels in column K are never merged; conflicts remain separate red rows.
Private Sub DrawDeviceBlocks(ByVal ws As Worksheet, _
                             ByRef slotOwner() As Long, _
                             ByVal rackUnits As Long)
    Dim u As Long, currentRow As Long, firstRow As Long
    Dim activeOwner As Long

    For u = rackUnits To 1 Step -1
        currentRow = FIRST_UNIT_ROW + rackUnits - u

        If slotOwner(u) <> activeOwner Then
            If activeOwner > 0 Then
                FrameDeviceBlock ws, firstRow, currentRow - 1
            End If

            activeOwner = slotOwner(u)
            If activeOwner > 0 Then firstRow = currentRow
        End If
    Next u

    If activeOwner > 0 Then
        FrameDeviceBlock ws, firstRow, FIRST_UNIT_ROW + rackUnits - 1
    End If
End Sub

Private Sub FrameDeviceBlock(ByVal ws As Worksheet, _
                             ByVal firstRow As Long, _
                             ByVal lastRow As Long)
    Dim deviceArea As Range

    ' Only the first row contains text, so merging never discards device data.
    If lastRow > firstRow Then
        ws.Range("L" & firstRow & ":L" & lastRow).Merge
        ws.Range("M" & firstRow & ":M" & lastRow).Merge
    End If

    Set deviceArea = ws.Range("L" & firstRow & ":M" & lastRow)
    With deviceArea
        .VerticalAlignment = xlVAlignCenter
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders.Color = RGB(113, 143, 172)
    End With
End Sub

Private Function GetRackUnits(ByVal rackId As String, _
                              ByRef note As String) As Long
    Dim racks As ListObject, entry As ListRow
    Dim idCol As Long, heightCol As Long
    Dim rawHeight As String
    Dim found As Boolean

    GetRackUnits = modRackData.DEFAULT_RACK_UNITS
    Set racks = modRackData.GetTable("Rack Info", "tblRacks")
    modRackData.CheckTableSize racks, modRackData.MAX_RACK_ROWS
    idCol = modRackData.ColumnIndex(racks, "Rack ID")
    heightCol = modRackData.ColumnIndex(racks, "Rack Height")

    For Each entry In racks.ListRows
        If StrComp(modRackData.RowText(entry, idCol), _
                   rackId, vbTextCompare) = 0 Then
            If found Then
                note = "Duplicate Rack ID in Rack Info; using first"
                Exit Function
            End If
            found = True
            rawHeight = modRackData.RowText(entry, heightCol)
            If Len(rawHeight) = 0 Then
                note = "Rack Height blank - defaulting to 42U"
            Else
                GetRackUnits = modRackData.RackHeight(rawHeight)
                If GetRackUnits = 0 Then
                    GetRackUnits = modRackData.DEFAULT_RACK_UNITS
                    note = "Invalid Rack Height (1-42 only) - using 42U"
                End If
            End If
        End If
    Next entry

    If Not found Then note = "Rack unregistered in Rack Info - using 42U"
End Function
