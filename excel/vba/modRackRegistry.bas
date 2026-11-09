Option Explicit

' Adds rack IDs found in tblInventory that are not already in tblRacks.
' Never deletes racks or changes an existing rack's height or barcode.
' Rack Height stays blank for newly added racks; the elevation defaults to 42U.
Public Sub SyncRackRegistry()
    Dim sourceTable As ListObject
    Dim rackTable As ListObject
    Dim sourceIdCol As Long, rackIdCol As Long, rackHeightCol As Long
    Dim sourceRow As ListRow, rackRow As ListRow
    Dim registered As Collection, existing As Variant
    Dim rackId As String, alreadyKnown As Boolean
    Dim newRow As ListRow
    Dim previousEvents As Boolean
    Dim created As Long

    previousEvents = Application.EnableEvents
    On Error GoTo HandleError
    Application.EnableEvents = False

    Set sourceTable = ThisWorkbook.Worksheets("Server Inventory").ListObjects("tblInventory")
    Set rackTable = ThisWorkbook.Worksheets("Rack Info").ListObjects("tblRacks")

    sourceIdCol = sourceTable.ListColumns("Rack ID").Index
    rackIdCol = rackTable.ListColumns("Rack ID").Index
    rackHeightCol = rackTable.ListColumns("Rack Height").Index

    ' A table accidentally covering the whole sheet can cause out-of-memory errors.
    If rackTable.ListRows.Count > 10000 Then
        Err.Raise vbObjectError + 501, "SyncRackRegistry", _
            "tblRacks has over 10,000 rows. Resize it to the actual rack records first."
    End If

    Set registered = New Collection
    For Each rackRow In rackTable.ListRows
        rackId = RackIdValue(rackRow.Range.Cells(1, rackIdCol))
        If Len(rackId) > 0 Then registered.Add rackId
    Next rackRow

    For Each sourceRow In sourceTable.ListRows
        rackId = RackIdValue(sourceRow.Range.Cells(1, sourceIdCol))
        If Len(rackId) = 0 Then GoTo NextSourceRow

        alreadyKnown = False
        For Each existing In registered
            If StrComp(CStr(existing), rackId, vbTextCompare) = 0 Then
                alreadyKnown = True
                Exit For
            End If
        Next existing

        If Not alreadyKnown Then
            ' Append only. Existing rack heights and barcodes remain untouched.
            Set newRow = rackTable.ListRows.Add
            With newRow.Range.Cells(1, rackIdCol)
                .NumberFormat = "@"
                .Value2 = rackId
            End With
            newRow.Range.Cells(1, rackHeightCol).ClearContents
            registered.Add rackId
            created = created + 1
        End If
NextSourceRow:
    Next sourceRow

CleanExit:
    Application.EnableEvents = previousEvents
    Exit Sub

HandleError:
    MsgBox "Rack registration error: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Private Function RackIdValue(ByVal cell As Range) As String
    If IsError(cell.Value2) Then Exit Function
    RackIdValue = Trim$(CStr(cell.Value2))
End Function
