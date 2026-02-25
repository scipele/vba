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
'| By Name/Date | T.Sciple, 11/28/2024                                        |

Public Function FindTableValueWithRng(ByRef tableRange As Range, _
                                      ByVal rowValueLkp As String, _
                                      ByVal colValueLkp As String) As Variant
    
    'Convert the table range to a listobject
    Dim tableObj As ListObject
    Set tableObj = tableRange.ListObject
    
    Dim lookupRange As Range
    Set lookupRange = tableObj.Range
    
    ' Set blank and error handling conditions
    If rowValueLkp = "" Then GoTo Lbl_HandleBlankCondition

    ' Find the matching row and column using MATCH against the full resolved range.
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

