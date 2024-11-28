Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | FindTableValue.vba                                          |
'| EntryPoint   | GetLaborRate                                                |
'| Purpose      | Lookup Rates in Various Tables                              |
'| Inputs       | tableRange, rowValueLkpm, colmHeader                        |
'| Outputs      | Data in Table that cooresponds to the row, colm match       |
'| Ex Formula   | =FindTableValue(INDIRECT(R8&"[#All]"),R6,R7)                |
'| cell R8      | Table1                                                      |
'| cell R6      | rowValueLkp                                                 |
'| cell R7      | colValueLkp                                                 |
'| Dependencies | none                                                        |
'| By Name/Date | T.Sciple, 11/27/2024                                        |

Public Function FindTableValue(ByRef tableRange As Range, _
                               ByVal rowValueLkp As String, _
                               ByVal colValueLkp As String) As Variant
    
    Dim rowMatch As Range
    Dim colMatch As Range
    
    On Error GoTo ErrHandler
    ' Find the matching row header
    Set rowMatch = tableRange.Columns(1).Find(What:=rowValueLkp, LookIn:=xlValues, LookAt:=xlWhole)
    If rowMatch Is Nothing Then
        FindTableValue = "Row Not Found"
        Exit Function
    End If

    ' Find the matching column header
    Set colMatch = tableRange.Rows(1).Find(What:=colValueLkp, LookIn:=xlValues, LookAt:=xlWhole)
    If colMatch Is Nothing Then
        FindTableValue = "Column Not Found"
        Exit Function
    End If

    ' Return the cell value that matches the specified row/column
    FindTableValue = tableRange.Cells(rowMatch.Row - tableRange.Rows(1).Row + 1, _
                                      colMatch.Column - tableRange.Columns(1).Column + 1).Value
    Exit Function

ErrHandler:
    FindTableValue = "Error: " & Err.Description
End Function


