Option Explicit

' Physical serial scans are compared with tblInventory. The report also flags
' invalid U positions, heights, duplicate serials and occupied-U collisions.
' Audit output never updates or deletes live inventory records.

Public Sub PublishRackAudit(ByVal rackInput As String, _
                            ByVal scans As Object, _
                            ByRef resultMessage As String)
    Dim inventory As ListObject, item As ListRow, ws As Worksheet
    Dim canonical As String, reason As String, capacity As Long
    Dim rackCol As Long, uCol As Long, heightCol As Long
    Dim hostCol As Long, serialCol As Long
    Dim serials As Object, inventoryRows As Collection
    Dim unitCount(1 To 42) As Long
    Dim rowNumber As Long, itemNumber As Long, bottomU As Long, heightU As Long
    Dim topU As Long, u As Long, rawHeight As String, serial As String
    Dim issues As String, scanStatus As String, key As Variant
    Dim verified As Long, missing As Long, unscannable As Long
    Dim unknown As Long, badRecords As Long, expectedCount As Long

    On Error GoTo HandleError
    resultMessage = ""
    If scans Is Nothing Then
        resultMessage = "The scan collection has not been initialized."
        Exit Sub
    End If
    capacity = modServerEntry.RackCapacity(rackInput, canonical, reason)
    If capacity = 0 Then
        resultMessage = reason
        Exit Sub
    End If

    Set inventory = modRackData.GetTable("Server Inventory", "tblInventory")
    modRackData.CheckTableSize inventory, modRackData.MAX_INVENTORY_ROWS
    rackCol = modRackData.ColumnIndex(inventory, "Rack ID")
    uCol = modRackData.ColumnIndex(inventory, "U Position")
    heightCol = modRackData.ColumnIndex(inventory, "Height (U)", False)
    hostCol = modRackData.ColumnIndex(inventory, "Hostname")
    serialCol = modRackData.ColumnIndex(inventory, "Serial Number")

    Set serials = CreateObject("Scripting.Dictionary")
    serials.CompareMode = vbTextCompare
    Set inventoryRows = New Collection
    For Each item In inventory.ListRows
        If StrComp(modRackData.RowText(item, rackCol), canonical, vbTextCompare) = 0 Then
            inventoryRows.Add item
            serial = modRackData.RowText(item, serialCol)
            If Len(serial) > 0 Then
                If serials.Exists(serial) Then
                    serials(serial) = CLng(serials(serial)) + 1
                Else
                    serials.Add serial, 1
                End If
            End If
            bottomU = modRackData.ParseU(modRackData.RowText(item, uCol))
            heightU = ServerHeight(item, heightCol)
            topU = bottomU + heightU - 1
            If bottomU > 0 And heightU > 0 And topU <= capacity Then
                For u = bottomU To topU
                    unitCount(u) = unitCount(u) + 1
                Next u
            End If
        End If
    Next item

    Set ws = OwnReportSheet()
    ws.Cells.Clear
    ws.Range("J1").Value2 = "RACKSCAN_GENERATED_AUDIT"
    ws.Range("A1:F1").Merge
    ws.Range("A1").Value2 = "RACKSCAN AUDIT - " & canonical
    With ws.Range("A1:F1")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(35, 53, 74)
        .Font.Size = 15
    End With
    ws.Range("A2").Value2 = "Audit date:"
    ws.Range("B2").Value = Now
    ws.Range("B2").NumberFormat = "mm/dd/yyyy hh:mm"
    ws.Range("A3").Value2 = "Rack capacity: " & capacity & "U"
    ws.Range("A5").Value2 = "Serial Number"
    ws.Range("B5").Value2 = "Hostname"
    ws.Range("C5").Value2 = "Bottom U"
    ws.Range("D5").Value2 = "Height (U)"
    ws.Range("E5").Value2 = "Physical Scan"
    ws.Range("F5").Value2 = "Data Quality"
    With ws.Range("A5:F5")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(54, 74, 95)
    End With

    rowNumber = 6
    For itemNumber = 1 To inventoryRows.Count
        Set item = inventoryRows(itemNumber)
        serial = modRackData.RowText(item, serialCol)
        bottomU = modRackData.ParseU(modRackData.RowText(item, uCol))
        rawHeight = ""
        If heightCol > 0 Then rawHeight = modRackData.RowText(item, heightCol)
        heightU = ServerHeight(item, heightCol)
        topU = bottomU + heightU - 1
        issues = ""

        If bottomU = 0 Then issues = AppendIssue(issues, "INVALID BOTTOM U")
        If heightU = 0 Then issues = AppendIssue(issues, "INVALID HEIGHT")
        If bottomU > 0 And heightU > 0 Then
            If topU > capacity Then
                issues = AppendIssue(issues, "EXCEEDS RACK")
            Else
                For u = bottomU To topU
                    If unitCount(u) > 1 Then
                        issues = AppendIssue(issues, "U OVERLAP")
                        Exit For
                    End If
                Next u
            End If
        End If
        If Len(serial) > 0 Then
            If CLng(serials(serial)) > 1 Then issues = AppendIssue(issues, "DUPLICATE SERIAL")
        End If
        If Len(issues) = 0 Then
            issues = "OK"
        Else
            badRecords = badRecords + 1
        End If

        If Len(serial) = 0 Then
            scanStatus = "NO SERIAL / UNVERIFIABLE"
            unscannable = unscannable + 1
        ElseIf scans.Exists(serial) Then
            scanStatus = "VERIFIED"
            verified = verified + 1
        Else
            scanStatus = "MISSING"
            missing = missing + 1
        End If

        ws.Cells(rowNumber, 1).NumberFormat = "@"
        ws.Cells(rowNumber, 1).Value2 = serial
        ws.Cells(rowNumber, 2).Value2 = modRackData.RowText(item, hostCol)
        ws.Cells(rowNumber, 3).Value2 = modRackData.RowText(item, uCol)
        ws.Cells(rowNumber, 4).Value2 = IIf(Len(rawHeight) = 0, "1", rawHeight)
        ws.Cells(rowNumber, 5).Value2 = scanStatus
        ws.Cells(rowNumber, 6).Value2 = issues
        If scanStatus = "MISSING" Or issues <> "OK" Then
            ws.Range("E" & rowNumber & ":F" & rowNumber).Interior.Color = RGB(255, 225, 225)
        ElseIf scanStatus = "VERIFIED" Then
            ws.Cells(rowNumber, 5).Interior.Color = RGB(215, 239, 218)
        End If
        rowNumber = rowNumber + 1
    Next itemNumber

    For Each key In scans.Keys
        If Not serials.Exists(CStr(key)) Then
            ws.Cells(rowNumber, 1).NumberFormat = "@"
            ws.Cells(rowNumber, 1).Value2 = CStr(key)
            ws.Cells(rowNumber, 5).Value2 = "UNEXPECTED SCAN"
            ws.Cells(rowNumber, 5).Interior.Color = RGB(255, 225, 225)
            unknown = unknown + 1
            rowNumber = rowNumber + 1
        End If
    Next key

    expectedCount = inventoryRows.Count
    ws.Range("A4").Value2 = "Expected: " & expectedCount & "  |  Verified: " & verified & _
                            "  |  Missing: " & missing & "  |  Unknown: " & unknown & _
                            "  |  No serial: " & unscannable & "  |  Data issues: " & badRecords
    ws.Columns("A:F").AutoFit
    ws.Columns("J").Hidden = True
    ws.Activate
    ws.Range("A1").Select
    resultMessage = "Audit complete: " & verified & " verified, " & missing & _
                    " missing, " & unknown & " unexpected, " & badRecords & " data issues."
    Exit Sub
HandleError:
    resultMessage = "Audit failed (" & Err.Number & "): " & Err.Description
End Sub

Private Function ServerHeight(ByVal item As ListRow, ByVal heightColumn As Long) As Long
    Dim value As String
    ServerHeight = 1
    If heightColumn = 0 Then Exit Function
    value = modRackData.RowText(item, heightColumn)
    If Len(value) > 0 Then ServerHeight = modRackData.ParseU(value, False)
End Function

Private Function AppendIssue(ByVal existing As String, ByVal addition As String) As String
    If Len(existing) > 0 Then
        AppendIssue = existing & "; " & addition
    Else
        AppendIssue = addition
    End If
End Function

Private Function OwnReportSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Rack Audit Report")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "Rack Audit Report"
        ws.Range("J1").Value2 = "RACKSCAN_GENERATED_AUDIT"
    ElseIf modRackData.TextValue(ws.Range("J1").Value2) <> "RACKSCAN_GENERATED_AUDIT" Then
        Err.Raise vbObjectError + 761, "OwnReportSheet", _
                  "Existing Rack Audit Report sheet is not owned by RackScan. No sheet was cleared."
    End If
    Set OwnReportSheet = ws
End Function
