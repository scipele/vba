Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | TblLookupByName.vba                                         |
'| EntryPoint   | FindTblValueWithName                                        |
'| Purpose      | Return Value from Table by searching for matching Row/Com   |
'| Inputs       | tableName, rowValueLkp, colValueLkp                         |
'| Outputs      | Data in Table that cooresponds to the row, colm match       |
'| Ex Formula   | =FindTblValueWithName(C12,D12,E12)                          |
'| cell c12     | Table1                                                      |
'| cell d12     | rowValueLkp                                                 |
'| cell e12     | colValueLkp                                                 |
'| Dependencies | none                                                        |
'| By Name/Date | T.Sciple, 11/28/2024                                        |


Public Function FindTblValueWithName(ByVal tableName As String, _
                                     ByVal rowValueLkp As String, _
                                     ByVal colValueLkp As String) As Variant

    If rowValueLkp = "" Then GoTo Lbl_HandleBlankCondition

    ' Find the sheet where the table is located using the function GetSheetNameWhereTableIsLocated
    Dim sht_name As String
    sht_name = GetSheetNameWhereTableIsLocated(tableName)
    
    If sht_name = "-1" Then
        FindTblValueWithName = "Error: Table not found"
        Exit Function
    End If

    Dim tableRange As Range
    Set tableRange = ThisWorkbook.Sheets(sht_name).ListObjects(tableName).Range

    Dim row_match As Range, col_match As Range
    Set row_match = tableRange.Columns(1).Find(What:=rowValueLkp, LookIn:=xlValues, LookAt:=xlWhole)
    Set col_match = tableRange.Rows(1).Find(What:=colValueLkp, LookIn:=xlValues, LookAt:=xlWhole)

    If (col_match Is Nothing) Or (row_match Is Nothing) Then
        FindTblValueWithName = IIf(col_match Is Nothing, "Column Not Found", "Row Not Found")
        Exit Function
    End If

    Dim tbl_row As Long, tbl_col As Long
    tbl_row = row_match.Row - tableRange.Rows(1).Row + 1
    tbl_col = col_match.Column - tableRange.Columns(1).Column + 1

    FindTblValueWithName = tableRange.Cells(tbl_row, tbl_col).Value
    Exit Function

Lbl_HandleBlankCondition:
    FindTblValueWithName = 0
End Function


Private Function GetSheetNameWhereTableIsLocated(tableName)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim found As Boolean
    'found = False
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set tbl = ws.ListObjects(tableName)
        On Error GoTo 0
        
        If Not tbl Is Nothing Then
            found = True
            GetSheetNameWhereTableIsLocated = ws.Name
            Exit For ' Exit loop once the table is found
        End If
    Next ws
    
    If Not found Then GetSheetNameWhereTableIsLocated = "-1"
    
End Function