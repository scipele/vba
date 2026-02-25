Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | FindTableValueWithRng.vba                                   |
'| EntryPoint   | GetLaborRate                                                |
'| Purpose      | Lookup Rates in Various Tables                              |
'| Inputs       | tableRange, rowValueLkpm, colmHeader                        |
'| Outputs      | Data in Table that cooresponds to the row, colm match       |
'| Ex Formula   | =FindTableValueWithRng(INDIRECT(C11&"[#All]"),D11,E11)      |
'| cell C11     | Table1                                                      |
'| cell D11     | rowValueLkp                                                 |
'| cell E11     | colValueLkp                                                 |
'| Dependencies | none                                                        |
'| By Name/Date | T.Sciple, 12/04/2024                                        |

Public Function FindTableValueWithRng(ByRef tableRange As Range, _
                                      ByVal rowValueLkp As String, _
                                      ByVal colValueLkp As String) As Variant
    Dim lookupRange As Range
    Set lookupRange = ResolveLookupRange(tableRange)

    ' Set blank and error handling conditions
    If rowValueLkp = "" Then GoTo Lbl_HandleBlankCondition

    ' Find the matching row and column using MATCH against the full resolved range.
    ' This avoids filtered/visible-only range behavior during formula evaluation.
    Dim rowIndex As Variant, colIndex As Variant
    rowIndex = Application.Match(rowValueLkp, lookupRange.Columns(1), 0)
    colIndex = Application.Match(colValueLkp, lookupRange.Rows(1), 0)

    ' Return error string if the data is not found
    If IsError(colIndex) Or IsError(rowIndex) Then
        FindTableValueWithRng = IIf(IsError(colIndex), "Column Not Found", "Row Not Found")
        Exit Function
    End If

    ' Return the cell value that matches the specified row/column
    FindTableValueWithRng = lookupRange.Cells(CLng(rowIndex), CLng(colIndex)).Value
    Exit Function
    
Lbl_HandleBlankCondition:
    FindTableValueWithRng = 0  'Return a zero if there is empty data so it doesn't throw an error in this case
    Exit Function

End Function

Private Function ResolveLookupRange(ByVal sourceRange As Range) As Range
    Dim tableObj As ListObject
    On Error Resume Next
    Set tableObj = sourceRange.ListObject
    On Error GoTo 0

    If Not tableObj Is Nothing Then
        Set ResolveLookupRange = tableObj.Range
        Exit Function
    End If

    If sourceRange.Areas.Count = 1 Then
        Set ResolveLookupRange = sourceRange
        Exit Function
    End If

    Dim area As Range
    Dim minRow As Long, minCol As Long, maxRow As Long, maxCol As Long
    minRow = sourceRange.Worksheet.Rows.Count
    minCol = sourceRange.Worksheet.Columns.Count

    For Each area In sourceRange.Areas
        If area.Row < minRow Then minRow = area.Row
        If area.Column < minCol Then minCol = area.Column
        If area.Row + area.Rows.Count - 1 > maxRow Then maxRow = area.Row + area.Rows.Count - 1
        If area.Column + area.Columns.Count - 1 > maxCol Then maxCol = area.Column + area.Columns.Count - 1
    Next area

    Set ResolveLookupRange = sourceRange.Worksheet.Range( _
        sourceRange.Worksheet.Cells(minRow, minCol), _
        sourceRange.Worksheet.Cells(maxRow, maxCol))
End Function