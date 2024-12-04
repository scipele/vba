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
    
    
    ' Set blank and error handling conditions
    If rowValueLkp = "" Then GoTo Lbl_HandleBlankCondition
    
    ' Find the matching row and column
    Dim row_match As Range, col_match As Range
    
    ' Used the .find method to match mathing row in columns(1), and then matching column from rows(1)
    Set row_match = tableRange.Columns(1).Find(What:=rowValueLkp, LookIn:=xlValues, LookAt:=xlWhole)
    Set col_match = tableRange.Rows(1).Find(What:=colValueLkp, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' Return error string if the data is not found
    If (col_match Is Nothing) Or (row_match Is Nothing) Then
        FindTableValueWithRng = IIf(col_match Is Nothing, "Column Not Found", "Row Not Found")
        Exit Function
    End If

    ' Determine the location of the table in the sheet
    Dim tbl_first_row As Long, tbl_first_col As Long
    tbl_first_row = tableRange.Rows(1).Row
    tbl_first_col = tableRange.Columns(1).Column
    
    ' Next calculate the relative position of the data in the table verses where the table is located in the sheet
    Dim tbl_row As Long, tbl_col As Long
    tbl_row = row_match.Row - tbl_first_row + 1
    tbl_col = col_match.Column - tbl_first_col + 1

    ' Return the cell value that matches the specified row/column
    FindTableValueWithRng = tableRange.Cells(tbl_row, tbl_col).Value
    Exit Function
    
Lbl_HandleBlankCondition:
    FindTableValueWithRng = 0  'Return a zero if there is empty data so it doesn't throw an error in this case
    Exit Function

End Function