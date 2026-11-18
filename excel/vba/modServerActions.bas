Option Explicit

' Edit and decommission workflows. Rack U is the BOTTOM of the server.
' These operations refuse ambiguous searches rather than modifying the wrong row.

Public Function FindServerFromLookup(ByRef issue As String) As ListRow
    Dim ws As Worksheet, inventory As ListObject, item As ListRow
    Dim rackCol As Long, uCol As Long, matchCol As Long
    Dim rackId As String, rawU As String, searchValue As String
    Dim targetU As Long, count As Long
    Dim found As ListRow

    issue = ""
    Set ws = ThisWorkbook.Worksheets("Rack Lookup")
    Set inventory = modRackData.GetTable("Server Inventory", "tblInventory")
    modRackData.CheckTableSize inventory, modRackData.MAX_INVENTORY_ROWS

    rackCol = modRackData.ColumnIndex(inventory, "Rack ID")
    uCol = modRackData.ColumnIndex(inventory, "U Position")

    If Len(modRackData.TextValue(ws.Range("C6").Value2)) > 0 Then
        matchCol = modRackData.ColumnIndex(inventory, "Hostname")
        searchValue = modRackData.TextValue(ws.Range("C6").Value2)
    ElseIf Len(modRackData.TextValue(ws.Range("C8").Value2)) > 0 Then
        matchCol = modRackData.ColumnIndex(inventory, "Serial Number")
        searchValue = modRackData.TextValue(ws.Range("C8").Value2)
    ElseIf Len(modRackData.TextValue(ws.Range("C10").Value2)) > 0 Then
        matchCol = modRackData.ColumnIndex(inventory, "Asset Tag")
        searchValue = modRackData.TextValue(ws.Range("C10").Value2)
    Else
        rackId = modRackData.TextValue(ws.Range("C3").Value2)
        rawU = modRackData.TextValue(ws.Range("C4").Value2)
        targetU = modRackData.ParseU(rawU)
        If Len(rackId) = 0 Or targetU = 0 Then
            issue = "Search by hostname, serial, asset tag, or Rack ID plus bottom U first."
            Exit Function
        End If
    End If

    For Each item In inventory.ListRows
        If matchCol > 0 Then
            If StrComp(modRackData.RowText(item, matchCol), searchValue, vbTextCompare) = 0 Then
                count = count + 1
                If count = 1 Then Set found = item
            End If
        ElseIf StrComp(modRackData.RowText(item, rackCol), rackId, vbTextCompare) = 0 Then
            If modRackData.ParseU(modRackData.RowText(item, uCol)) = targetU Then
                count = count + 1
                If count = 1 Then Set found = item
            End If
        End If
        If count > 1 Then Exit For
    Next item

    If count = 0 Then
        issue = "No matching server. Check the Rack Lookup search fields."
    ElseIf count > 1 Then
        issue = "Multiple servers match. Search with a more specific identifier."
    Else
        Set FindServerFromLookup = found
    End If
End Function

' Checks the full occupied range, excluding the row being edited.
Private Function PlacementIssue(ByVal target As ListRow, _
                                ByVal rackId As String, _
                                ByVal bottomU As Long, _
                                ByVal heightU As Long, _
                                ByVal capacity As Long) As String
    Dim inventory As ListObject, item As ListRow
    Dim rackCol As Long, uCol As Long, heightCol As Long
    Dim otherBottom As Long, otherHeight As Long, rawHeight As String
    Dim otherTop As Long, newTop As Long

    Set inventory = modRackData.GetTable("Server Inventory", "tblInventory")
    rackCol = modRackData.ColumnIndex(inventory, "Rack ID")
    uCol = modRackData.ColumnIndex(inventory, "U Position")
    heightCol = modRackData.ColumnIndex(inventory, "Height (U)")
    newTop = bottomU + heightU - 1

    For Each item In inventory.ListRows
        If item.Range.Row <> target.Range.Row Then
            If StrComp(modRackData.RowText(item, rackCol), rackId, vbTextCompare) = 0 Then
                otherBottom = modRackData.ParseU(modRackData.RowText(item, uCol))
                rawHeight = modRackData.RowText(item, heightCol)
                otherHeight = 1
                If Len(rawHeight) > 0 Then otherHeight = modRackData.ParseU(rawHeight, False)
                otherTop = otherBottom + otherHeight - 1

                If otherBottom = 0 Or otherHeight = 0 Or otherTop > capacity Then
                    PlacementIssue = "Existing equipment in this rack has invalid U/height data."
                    Exit Function
                End If
                If bottomU <= otherTop And newTop >= otherBottom Then
                    PlacementIssue = "U-position conflict with the server at U" & otherBottom & _
                                     "-U" & otherTop & "."
                    Exit Function
                End If
            End If
        End If
    Next item
End Function

Public Function TryEditServer(ByVal target As ListRow, _
                              ByVal rackInput As String, _
                              ByVal bottomText As String, _
                              ByVal heightText As String, _
                              ByVal hostname As String, _
                              ByVal serialNumber As String, _
                              ByVal assetTag As String, _
                              ByRef message As String) As Boolean
    Dim inv As ListObject, canonical As String, reason As String
    Dim capacity As Long, bottomU As Long, heightU As Long
    Dim col(1 To 6) As Long, oldValue(1 To 6) As Variant
    Dim i As Long, priorEvents As Boolean, writing As Boolean
    Dim errorNumber As Long, errorText As String

    On Error GoTo HandleError
    message = ""
    If target Is Nothing Then
        message = "No server was selected."
        Exit Function
    End If
    Set inv = modRackData.GetTable("Server Inventory", "tblInventory")
    modRackData.CheckTableSize inv, modRackData.MAX_INVENTORY_ROWS
    If Not (target.Parent Is inv) Then
        message = "The selected server is not in tblInventory."
        Exit Function
    End If

    capacity = modServerEntry.RackCapacity(rackInput, canonical, reason)
    If capacity = 0 Then
        message = reason
        Exit Function
    End If
    bottomU = modRackData.ParseU(bottomText)
    heightU = modRackData.ParseU(heightText, False)
    If bottomU = 0 Or heightU = 0 Then
        message = "Bottom U and height must be whole numbers from 1 to 42."
        Exit Function
    End If
    If bottomU + heightU - 1 > capacity Then
        message = "Equipment extends beyond the " & capacity & "U rack."
        Exit Function
    End If
    hostname = Trim$(hostname)
    serialNumber = Trim$(serialNumber)
    assetTag = Trim$(assetTag)
    If Len(hostname) > 255 Or Len(serialNumber) > 255 Or Len(assetTag) > 255 Then
        message = "Identifiers must be 255 characters or fewer."
        Exit Function
    End If

    message = PlacementIssue(target, canonical, bottomU, heightU, capacity)
    If Len(message) > 0 Then Exit Function

    col(1) = modRackData.ColumnIndex(inv, "Rack ID")
    col(2) = modRackData.ColumnIndex(inv, "U Position")
    col(3) = modRackData.ColumnIndex(inv, "Height (U)")
    col(4) = modRackData.ColumnIndex(inv, "Hostname")
    col(5) = modRackData.ColumnIndex(inv, "Serial Number")
    col(6) = modRackData.ColumnIndex(inv, "Asset Tag")
    For i = 1 To 6
        oldValue(i) = target.Range.Cells(1, col(i)).Value2
    Next i

    priorEvents = Application.EnableEvents
    Application.EnableEvents = False
    writing = True
    SetText target.Range.Cells(1, col(1)), canonical
    target.Range.Cells(1, col(2)).Value2 = bottomU
    target.Range.Cells(1, col(3)).Value2 = heightU
    SetText target.Range.Cells(1, col(4)), hostname
    SetText target.Range.Cells(1, col(5)), serialNumber
    SetText target.Range.Cells(1, col(6)), assetTag
    writing = False
    Application.EnableEvents = priorEvents

    TryEditServer = True
    message = "Server updated in " & canonical & " at U" & bottomU & "."
    Exit Function

HandleError:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    If writing Then
        For i = 1 To 6
            target.Range.Cells(1, col(i)).Value2 = oldValue(i)
        Next i
        Application.EnableEvents = priorEvents
    End If
    On Error GoTo 0
    message = "Edit error (" & errorNumber & "): " & errorText
End Function

' Archives ALL current inventory columns, including custom columns, before
' deleting the source row. Existing archives must have matching field names.
Public Function TryDecomServer(ByVal target As ListRow, _
                               ByVal reason As String, _
                               ByRef message As String) As Boolean
    Dim inv As ListObject, archive As ListObject, newRow As ListRow
    Dim column As ListColumn, destination As Long
    Dim dateCol As Long, reasonCol As Long
    Dim priorEvents As Boolean, writing As Boolean, deleted As Boolean
    Dim errorNumber As Long, errorText As String

    On Error GoTo HandleError
    message = ""
    If target Is Nothing Then
        message = "No server selected."
        Exit Function
    End If
    Set inv = modRackData.GetTable("Server Inventory", "tblInventory")
    If Not (target.Parent Is inv) Then
        message = "Selected server is no longer in tblInventory."
        Exit Function
    End If
    modRackData.CheckTableSize inv, modRackData.MAX_INVENTORY_ROWS
    If Len(reason) > 255 Then
        message = "Reason must be 255 characters or fewer."
        Exit Function
    End If
    Set archive = EnsureArchive(inv)
    If archive.ListRows.Count >= modRackData.MAX_INVENTORY_ROWS Then
        message = "The decommission archive reached its configured row limit."
        Exit Function
    End If

    ' Resolve every destination column before writing any data.
    For Each column In inv.ListColumns
        destination = modRackData.ColumnIndex(archive, column.Name)
    Next column
    dateCol = modRackData.ColumnIndex(archive, "Decom Date")
    reasonCol = modRackData.ColumnIndex(archive, "Decom Reason")

    priorEvents = Application.EnableEvents
    Application.EnableEvents = False
    writing = True
    Set newRow = archive.ListRows.Add
    For Each column In inv.ListColumns
        destination = modRackData.ColumnIndex(archive, column.Name)
        With newRow.Range.Cells(1, destination)
            .NumberFormat = target.Range.Cells(1, column.Index).NumberFormat
            .Value2 = target.Range.Cells(1, column.Index).Value2
        End With
    Next column
    newRow.Range.Cells(1, dateCol).Value = Now
    newRow.Range.Cells(1, dateCol).NumberFormat = "mm/dd/yyyy hh:mm"
    SetText newRow.Range.Cells(1, reasonCol), Trim$(reason)
    target.Delete
    deleted = True
    writing = False
    Application.EnableEvents = priorEvents
    TryDecomServer = True
    message = "Server archived to tblDecom and removed from active inventory."
    Exit Function

HandleError:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    If writing Then
        If Not deleted Then
            If Not newRow Is Nothing Then newRow.Delete
        End If
        Application.EnableEvents = priorEvents
    End If
    On Error GoTo 0
    message = "Decom error (" & errorNumber & "): " & errorText
End Function

Private Function EnsureArchive(ByVal inventory As ListObject) As ListObject
    Dim ws As Worksheet, table As ListObject, column As ListColumn
    Dim count As Long, n As Long

    If modRackData.ColumnIndex(inventory, "Decom Date", False) > 0 Or _
       modRackData.ColumnIndex(inventory, "Decom Reason", False) > 0 Then
        Err.Raise vbObjectError + 750, "EnsureArchive", _
                  "Inventory already has a Decom Date or Decom Reason column. Rename or map it explicitly first."
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Decommissioned Servers")
    On Error GoTo 0
    If ws Is Nothing Then
        ' Never overwrite a worksheet somebody else has already populated.
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "Decommissioned Servers"
        count = inventory.ListColumns.Count
        For n = 1 To count
            ws.Cells(1, n).Value2 = inventory.ListColumns(n).Name
        Next n
        ws.Cells(1, count + 1).Value2 = "Decom Date"
        ws.Cells(1, count + 2).Value2 = "Decom Reason"
        Set table = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(2, count + 2)), , xlYes)
        table.Name = "tblDecom"
        If table.ListRows.Count > 0 Then table.ListRows(1).Delete
        Set EnsureArchive = table
        Exit Function
    End If

    For Each table In ws.ListObjects
        If StrComp(table.Name, "tblDecom", vbTextCompare) = 0 Then
            For Each column In inventory.ListColumns
                n = modRackData.ColumnIndex(table, column.Name)
            Next column
            n = modRackData.ColumnIndex(table, "Decom Date")
            n = modRackData.ColumnIndex(table, "Decom Reason")
            modRackData.CheckTableSize table, modRackData.MAX_INVENTORY_ROWS
            Set EnsureArchive = table
            Exit Function
        End If
    Next table
    Err.Raise vbObjectError + 751, "EnsureArchive", _
              "Decommissioned Servers exists without tblDecom. No data was moved."
End Function

Private Sub SetText(ByVal cell As Range, ByVal value As String)
    cell.NumberFormat = "@"
    If Len(value) = 0 Then
        cell.ClearContents
    Else
        cell.Value2 = value
    End If
End Sub
