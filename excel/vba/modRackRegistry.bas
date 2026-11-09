Option Explicit

' Adds new inventory rack IDs to tblRacks without changing existing data.
' The lookup is case-insensitive and ignores leading/trailing spaces.
Public Sub SyncRackRegistry()
    Dim inventory As ListObject, racks As ListObject
    Dim sourceIdCol As Long, rackIdCol As Long, heightCol As Long
    Dim sourceRow As ListRow, rackRow As ListRow, newRow As ListRow
    Dim known As Object
    Dim rackId As String
    Dim previousEvents As Boolean

    previousEvents = Application.EnableEvents
    On Error GoTo HandleError
    Application.EnableEvents = False

    Set inventory = modRackData.GetTable("Server Inventory", "tblInventory")
    Set racks = modRackData.GetTable("Rack Info", "tblRacks")
    modRackData.CheckTableSize inventory, modRackData.MAX_INVENTORY_ROWS
    modRackData.CheckTableSize racks, modRackData.MAX_RACK_ROWS

    sourceIdCol = modRackData.ColumnIndex(inventory, "Rack ID")
    rackIdCol = modRackData.ColumnIndex(racks, "Rack ID")
    heightCol = modRackData.ColumnIndex(racks, "Rack Height")

    Set known = CreateObject("Scripting.Dictionary")
    known.CompareMode = vbTextCompare

    For Each rackRow In racks.ListRows
        rackId = modRackData.RowText(rackRow, rackIdCol)
        If Len(rackId) > 0 Then known(rackId) = True
    Next rackRow

    For Each sourceRow In inventory.ListRows
        rackId = modRackData.RowText(sourceRow, sourceIdCol)
        If Len(rackId) > 0 Then
            If Not known.Exists(rackId) Then
                If racks.ListRows.Count >= modRackData.MAX_RACK_ROWS Then
                    Err.Raise vbObjectError + 704, "SyncRackRegistry", _
                        "Rack Info table reached the configured row limit."
                End If

                Set newRow = racks.ListRows.Add
                With newRow.Range.Cells(1, rackIdCol)
                    .NumberFormat = "@"
                    .Value2 = rackId
                End With
                ' Table auto-fill can propagate formulas; leave height blank.
                newRow.Range.Cells(1, heightCol).ClearContents
                known.Add rackId, True
            End If
        End If
    Next sourceRow

CleanExit:
    Application.EnableEvents = previousEvents
    Exit Sub

HandleError:
    MsgBox "Rack registration error (" & Err.Number & "): " & _
           Err.Description, vbExclamation
    Resume CleanExit
End Sub

' Reports data-quality problems without deleting or changing any rack data.
' Called when Rack ID / Rack Height is edited, or manually from Alt+F8.
Public Sub ValidateRackRegistry()
    Dim racks As ListObject, entry As ListRow
    Dim idCol As Long, heightCol As Long
    Dim seen As Object
    Dim rackId As String, heightText As String
    Dim duplicateCount As Long, invalidHeights As Long, missingIds As Long
    Dim examples As String

    On Error GoTo HandleError
    Set racks = modRackData.GetTable("Rack Info", "tblRacks")
    modRackData.CheckTableSize racks, modRackData.MAX_RACK_ROWS
    idCol = modRackData.ColumnIndex(racks, "Rack ID")
    heightCol = modRackData.ColumnIndex(racks, "Rack Height")

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    For Each entry In racks.ListRows
        rackId = modRackData.RowText(entry, idCol)
        heightText = modRackData.RowText(entry, heightCol)

        If Len(rackId) > 0 Then
            If seen.Exists(rackId) Then
                duplicateCount = duplicateCount + 1
                If Len(examples) < 120 Then examples = examples & " " & rackId
            Else
                seen.Add rackId, True
            End If
        ElseIf Len(heightText) > 0 Then
            missingIds = missingIds + 1
        End If

        If Len(heightText) > 0 Then
            If modRackData.RackHeight(heightText) = 0 Then
                invalidHeights = invalidHeights + 1
            End If
        End If
    Next entry

    If duplicateCount > 0 Or invalidHeights > 0 Or missingIds > 0 Then
        MsgBox "Rack Info needs attention:" & vbCrLf & _
               "Duplicate Rack IDs: " & duplicateCount & vbCrLf & _
               "Invalid heights (must be 1-42): " & invalidHeights & vbCrLf & _
               "Heights without Rack IDs: " & missingIds & _
               IIf(Len(examples) > 0, vbCrLf & "Duplicate IDs:" & examples, ""), _
               vbExclamation, "Rack Info validation"
    End If
    Exit Sub

HandleError:
    MsgBox "Rack Info validation error (" & Err.Number & "): " & _
           Err.Description, vbExclamation
End Sub
