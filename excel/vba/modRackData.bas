Option Explicit

' Shared limits, table access and numeric parsing used by the rack modules.
Public Const MAX_RACK_UNITS As Long = 42
Public Const DEFAULT_RACK_UNITS As Long = 42
Public Const MAX_INVENTORY_ROWS As Long = 50000
Public Const MAX_RACK_ROWS As Long = 10000

Public Function GetTable(ByVal sheetName As String, _
                         ByVal tableName As String) As ListObject
    Dim sheet As Worksheet
    Dim table As ListObject

    For Each sheet In ThisWorkbook.Worksheets
        If StrComp(sheet.Name, sheetName, vbTextCompare) = 0 Then
            For Each table In sheet.ListObjects
                If StrComp(table.Name, tableName, vbTextCompare) = 0 Then
                    Set GetTable = table
                    Exit Function
                End If
            Next table
            Exit For
        End If
    Next sheet

    Err.Raise vbObjectError + 701, "modRackData.GetTable", _
        "Cannot find table '" & tableName & "' on sheet '" & sheetName & "'."
End Function

Public Sub CheckTableSize(ByVal table As ListObject, _
                          ByVal maxRows As Long)
    Dim rowCount As Long

    ' Count rows before attempting to iterate over an oversized Excel Table.
    rowCount = table.Range.Rows.Count - 1
    If table.ShowTotals Then rowCount = rowCount - 1

    If rowCount > maxRows Then
        Err.Raise vbObjectError + 702, "modRackData.CheckTableSize", _
            table.Name & " covers " & rowCount & " data rows (limit " & _
            maxRows & "). Use Table Design > Resize Table to keep only real records."
    End If
End Sub

Public Function ColumnIndex(ByVal table As ListObject, _
                            ByVal heading As String, _
                            Optional ByVal required As Boolean = True) As Long
    Dim col As ListColumn

    For Each col In table.ListColumns
        If StrComp(Trim$(col.Name), heading, vbTextCompare) = 0 Then
            ColumnIndex = col.Index
            Exit Function
        End If
    Next col

    If required Then
        Err.Raise vbObjectError + 703, "modRackData.ColumnIndex", _
            "Missing column '" & heading & "' in table " & table.Name & "."
    End If
End Function

Public Function TextValue(ByVal value As Variant) As String
    If IsError(value) Then Exit Function
    If IsNull(value) Then Exit Function
    If IsEmpty(value) Then Exit Function
    TextValue = Trim$(CStr(value))
End Function

Public Function RowText(ByVal entry As ListRow, ByVal column As Long) As String
    RowText = TextValue(entry.Range.Cells(1, column).Value2)
End Function

' Accepts 1..42 or U1..U42; 0 means invalid.
' Text must contain ASCII digits only after an optional U prefix.
Public Function ParseU(ByVal rawValue As String, _
                       Optional ByVal allowPrefix As Boolean = True) As Long
    Dim value As String
    Dim i As Long
    Dim code As Long
    Dim numericValue As Double

    value = UCase$(Trim$(rawValue))
    If allowPrefix Then
        If Left$(value, 1) = "U" Then value = Mid$(value, 2)
    End If
    If Len(value) = 0 Or Len(value) > 9 Then Exit Function

    For i = 1 To Len(value)
        code = AscW(Mid$(value, i, 1))
        If code < 48 Or code > 57 Then Exit Function
    Next i

    numericValue = CDbl(value)
    If numericValue >= 1 And numericValue <= MAX_RACK_UNITS Then
        ParseU = CLng(numericValue)
    End If
End Function

Public Function RackHeight(ByVal rawValue As String) As Long
    RackHeight = ParseU(rawValue, False)
End Function
