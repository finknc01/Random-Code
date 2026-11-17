Option Explicit

' Add Server workflow. U Position is always the BOTTOM occupied U.
' This module validates data and writes tblInventory; it does not own the form.

' Returns unique, nonblank Rack IDs in the order they appear in tblRacks.
Public Function RegisteredRackIds() As Collection
    Dim racks As ListObject, entry As ListRow
    Dim idCol As Long, rackId As String
    Dim seen As Object

    Set RegisteredRackIds = New Collection
    Set racks = modRackData.GetTable("Rack Info", "tblRacks")
    modRackData.CheckTableSize racks, modRackData.MAX_RACK_ROWS
    idCol = modRackData.ColumnIndex(racks, "Rack ID")

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    For Each entry In racks.ListRows
        rackId = modRackData.RowText(entry, idCol)
        If Len(rackId) > 0 Then
            If Not seen.Exists(rackId) Then
                seen.Add rackId, True
                RegisteredRackIds.Add rackId
            End If
        End If
    Next entry
End Function

' Returns 0 with a message for an unknown/duplicate rack or invalid rack height.
' A blank Rack Height means 42U; a nonblank invalid one must be corrected.
Public Function RackCapacity(ByVal rackInput As String, _
                             ByRef canonicalId As String, _
                             ByRef reason As String) As Long
    Dim racks As ListObject, entry As ListRow
    Dim idCol As Long, heightCol As Long
    Dim matchCount As Long, heightText As String
    Dim rackId As String

    canonicalId = ""
    reason = ""
    rackInput = Trim$(rackInput)
    If Len(rackInput) = 0 Then
        reason = "Select a Rack ID."
        Exit Function
    End If

    Set racks = modRackData.GetTable("Rack Info", "tblRacks")
    modRackData.CheckTableSize racks, modRackData.MAX_RACK_ROWS
    idCol = modRackData.ColumnIndex(racks, "Rack ID")
    heightCol = modRackData.ColumnIndex(racks, "Rack Height")

    For Each entry In racks.ListRows
        rackId = modRackData.RowText(entry, idCol)
        If StrComp(rackId, rackInput, vbTextCompare) = 0 Then
            matchCount = matchCount + 1
            If matchCount = 1 Then
                canonicalId = rackId
                heightText = modRackData.RowText(entry, heightCol)
            End If
        End If
    Next entry

    If matchCount = 0 Then
        reason = "Rack ID is not registered in Rack Info."
        Exit Function
    ElseIf matchCount > 1 Then
        reason = "Duplicate Rack ID in Rack Info. Correct it before adding servers."
        canonicalId = ""
        Exit Function
    End If

    If Len(heightText) = 0 Then
        RackCapacity = modRackData.DEFAULT_RACK_UNITS
    Else
        RackCapacity = modRackData.RackHeight(heightText)
        If RackCapacity = 0 Then
            reason = "Invalid Rack Height in Rack Info (must be 1-42)."
        End If
    End If
End Function

' Preselect the rack being viewed, including searches by device identifiers.
' Ambiguous/missing searches leave the dropdown unselected.
Public Function SuggestedRackId() As String
    Dim ws As Worksheet, inventory As ListObject, entry As ListRow
    Dim rackCol As Long, searchCol As Long
    Dim candidate As String, proposed As String, searchValue As String
    Dim canonical As String, reason As String

    Set ws = ThisWorkbook.Worksheets("Rack Lookup")
    If Len(modRackData.TextValue(ws.Range("C6").Value2)) > 0 Then
        searchValue = modRackData.TextValue(ws.Range("C6").Value2)
        searchCol = 1
    ElseIf Len(modRackData.TextValue(ws.Range("C8").Value2)) > 0 Then
        searchValue = modRackData.TextValue(ws.Range("C8").Value2)
        searchCol = 2
    ElseIf Len(modRackData.TextValue(ws.Range("C10").Value2)) > 0 Then
        searchValue = modRackData.TextValue(ws.Range("C10").Value2)
        searchCol = 3
    Else
        candidate = modRackData.TextValue(ws.Range("C3").Value2)
        If RackCapacity(candidate, canonical, reason) > 0 Then
            SuggestedRackId = canonical
        End If
        Exit Function
    End If

    Set inventory = modRackData.GetTable("Server Inventory", "tblInventory")
    modRackData.CheckTableSize inventory, modRackData.MAX_INVENTORY_ROWS
    rackCol = modRackData.ColumnIndex(inventory, "Rack ID")
    Select Case searchCol
        Case 1: searchCol = modRackData.ColumnIndex(inventory, "Hostname")
        Case 2: searchCol = modRackData.ColumnIndex(inventory, "Serial Number")
        Case 3: searchCol = modRackData.ColumnIndex(inventory, "Asset Tag")
    End Select

    For Each entry In inventory.ListRows
        If StrComp(modRackData.RowText(entry, searchCol), _
                   searchValue, vbTextCompare) = 0 Then
            candidate = modRackData.RowText(entry, rackCol)
            If Len(candidate) = 0 Then Exit Function

            If Len(proposed) = 0 Then
                proposed = candidate
            ElseIf StrComp(proposed, candidate, vbTextCompare) <> 0 Then
                Exit Function
            End If
        End If
    Next entry

    If Len(proposed) > 0 Then
        If RackCapacity(proposed, canonical, reason) > 0 Then
            SuggestedRackId = canonical
        End If
    End If
End Function

' Appends exactly one server row when validation succeeds.
' On a failed write, removes only the newly created row.
Public Function TryAddServer(ByVal rackInput As String, _
                             ByVal bottomText As String, _
                             ByVal heightText As String, _
                             ByVal hostname As String, _
                             ByVal serialNumber As String, _
                             ByVal assetTag As String, _
                             ByRef resultMessage As String) As Boolean
    Dim inventory As ListObject, entry As ListRow, newRow As ListRow
    Dim rackCol As Long, uCol As Long, heightCol As Long
    Dim hostCol As Long, serialCol As Long, assetCol As Long
    Dim capacity As Long, bottomU As Long, heightU As Long, topU As Long
    Dim existingBottom As Long, existingHeight As Long, existingTop As Long
    Dim existingHeightText As String, existingHost As String
    Dim invalidExisting As Long, canonical As String, reason As String
    Dim priorEvents As Boolean, writing As Boolean
    Dim errorNumber As Long, errorText As String

    On Error GoTo HandleError
    resultMessage = ""

    capacity = RackCapacity(rackInput, canonical, reason)
    If capacity = 0 Then
        resultMessage = reason
        Exit Function
    End If

    bottomU = modRackData.ParseU(bottomText)
    heightU = modRackData.ParseU(heightText, False)
    If bottomU = 0 Or heightU = 0 Then
        resultMessage = "U Position and Height (U) must be whole numbers from 1 to 42."
        Exit Function
    End If

    topU = bottomU + heightU - 1
    If topU > capacity Then
        resultMessage = "This server would extend to U" & topU & _
                        " but " & canonical & " is only " & capacity & "U tall."
        Exit Function
    End If

    hostname = Trim$(hostname)
    serialNumber = Trim$(serialNumber)
    assetTag = Trim$(assetTag)
    If Len(hostname) > 255 Or Len(serialNumber) > 255 Or Len(assetTag) > 255 Then
        resultMessage = "Hostname, Serial Number and Asset Tag must be 255 characters or less."
        Exit Function
    End If

    Set inventory = modRackData.GetTable("Server Inventory", "tblInventory")
    modRackData.CheckTableSize inventory, modRackData.MAX_INVENTORY_ROWS
    If inventory.ListRows.Count >= modRackData.MAX_INVENTORY_ROWS Then
        resultMessage = "The inventory table has reached its configured row limit."
        Exit Function
    End If

    rackCol = modRackData.ColumnIndex(inventory, "Rack ID")
    uCol = modRackData.ColumnIndex(inventory, "U Position")
    hostCol = modRackData.ColumnIndex(inventory, "Hostname")
    serialCol = modRackData.ColumnIndex(inventory, "Serial Number")
    assetCol = modRackData.ColumnIndex(inventory, "Asset Tag")
    heightCol = modRackData.ColumnIndex(inventory, "Height (U)")

    For Each entry In inventory.ListRows
        If StrComp(modRackData.RowText(entry, rackCol), _
                   canonical, vbTextCompare) = 0 Then
            existingBottom = modRackData.ParseU(modRackData.RowText(entry, uCol))
            existingHeightText = modRackData.RowText(entry, heightCol)
            existingHeight = 1
            If Len(existingHeightText) > 0 Then
                existingHeight = modRackData.ParseU(existingHeightText, False)
            End If

            If existingBottom = 0 Or existingHeight = 0 Or _
               existingBottom + existingHeight - 1 > capacity Then
                invalidExisting = invalidExisting + 1
            Else
                existingTop = existingBottom + existingHeight - 1
                If bottomU <= existingTop And topU >= existingBottom Then
                    existingHost = modRackData.RowText(entry, hostCol)
                    If Len(existingHost) = 0 Then existingHost = "(unnamed equipment)"
                    resultMessage = "U-position conflict with " & existingHost & _
                                    " (U" & existingBottom & "-U" & existingTop & ")."
                    Exit Function
                End If
            End If
        End If
    Next entry

    If invalidExisting > 0 Then
        resultMessage = "Cannot check for rack conflicts: " & invalidExisting & _
                        " existing record(s) in " & canonical & _
                        " have invalid U Position or Height (U). Correct these first."
        Exit Function
    End If

    priorEvents = Application.EnableEvents
    Application.EnableEvents = False
    writing = True

    Set newRow = inventory.ListRows.Add
    StoreText newRow.Range.Cells(1, rackCol), canonical
    newRow.Range.Cells(1, uCol).Value2 = bottomU
    newRow.Range.Cells(1, heightCol).Value2 = heightU
    StoreText newRow.Range.Cells(1, hostCol), hostname
    StoreText newRow.Range.Cells(1, serialCol), serialNumber
    StoreText newRow.Range.Cells(1, assetCol), assetTag

    Application.EnableEvents = priorEvents
    writing = False

    TryAddServer = True
    resultMessage = "Server added to " & canonical & " at U" & bottomU & _
                    " (" & heightU & "U)."
    Exit Function

HandleError:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    If writing Then
        If Not newRow Is Nothing Then newRow.Delete
        Application.EnableEvents = priorEvents
    End If
    On Error GoTo 0
    resultMessage = "Add Server error (" & errorNumber & "): " & errorText
End Function

' Updates the searches after a successful insert to highlight the new server.
Public Sub ShowAddedServer(ByVal rackId As String, ByVal bottomU As Long)
    Dim ws As Worksheet, previousEvents As Boolean
    Dim errorNumber As Long, errorText As String

    previousEvents = Application.EnableEvents
    On Error GoTo HandleError
    Set ws = ThisWorkbook.Worksheets("Rack Lookup")

    Application.EnableEvents = False
    StoreText ws.Range("C3"), rackId
    ws.Range("C4").Value2 = bottomU
    ws.Range("C6,C8,C10").ClearContents
    Application.EnableEvents = previousEvents

    ws.Activate
    ws.Range("C3").Select
    ws.Range("E3").Calculate
    modRackElevation.RenderRackElevation
    Exit Sub

HandleError:
    errorNumber = Err.Number
    errorText = Err.Description
    Application.EnableEvents = previousEvents
    Err.Raise errorNumber, "modServerEntry.ShowAddedServer", errorText
End Sub

Private Sub StoreText(ByVal cell As Range, ByVal value As String)
    With cell
        .NumberFormat = "@"
        If Len(value) = 0 Then
            .ClearContents
        Else
            .Value2 = value
        End If
    End With
End Sub
