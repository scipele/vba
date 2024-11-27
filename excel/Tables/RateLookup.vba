Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | rateLookup.vba                                              |
'| EntryPoint   | GetLaborRate                                                |
'| Purpose      | Lookup Rates in Various Tables                              |
'| Inputs       | rateTbl, lookupValue, fieldName                             |
'| Outputs      | labor/perdiem rate/xlErrNA                                  |
'| Dependencies | none                                                        |
'| By Name/Date | T.Sciple, 11/27/2024                                        |

Public Function GetLaborRate(rateTbl As String, lookupValue As Variant, fieldName As String) As Variant

    On Error GoTo ErrHandler
    
    'If classification is blank then return zero
    If lookupValue = "" Then
        GetLaborRate = 0
        Exit Function
    End If

    Dim tbl_index As String
    Dim tbl_ref As String
    If Not IsEmpty(rateTbl) Then
        tbl_index = Left(rateTbl, InStr(1, rateTbl, "_", vbTextCompare) - 1)
        tbl_ref = "Table" & tbl_index
    End If
    
    ' Dynamically reference the selected table
    Dim tbl As Range
    Set tbl = ThisWorkbook.Sheets(tbl_index).ListObjects(tbl_ref).Range
    
    ' Find the column number for the given field name
    Dim field_col As Integer
    field_col = 0
    
    'Search Field Name to get matching column number
    Dim headerRow As Range
    Dim found_field As Range
    Set headerRow = tbl.Rows(1)
    Set found_field = headerRow.Find(What:=fieldName, LookIn:=xlValues, LookAt:=xlWhole)
    If Not found_field Is Nothing Then
        field_col = found_field.Column - tbl.Columns(1).Column + 1
    End If
    
    'Search rows in the first column to get matching "craft designation"
    Dim first_colm As Range
    Set first_colm = tbl.Columns(1)
    Dim row_no As Integer
    Dim matching_row As Range
    Set matching_row = first_colm.Find(What:=lookupValue, LookIn:=xlValues, LookAt:=xlWhole)
    If Not matching_row Is Nothing Then
        row_no = matching_row.Row - tbl.Rows(1).Row + 1
    End If
    
    ' If the field name or row match is not found, return #N/A
    If field_col = 0 Or matching_row = 0 Then
        GetLaborRate = CVErr(xlErrNA)
        Exit Function
    End If

    ' Return the value from the specified field (column) and row
    GetLaborRate = tbl.Cells(row_no, field_col).Value
    Exit Function

ErrHandler:
    GetLaborRate = CVErr(xlErrNA) ' Return #N/A on errorEnd Function

End Function