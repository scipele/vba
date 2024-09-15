Option Explicit
' filename:     get_last_row_in_column.vba
' Purpose:      Finds the last non-blank cell in a single row or column uses cells function to count all rows in the column
' Usage:        row = last_row_in_colm(sht, "L")
' Dependencies: None
' By:  T.Sciple, 09/07/2024
Public Sub test()
    Dim sht As String
    sht = "Sheet1"
    Dim row As Long
    row = last_row_in_colm(sht, "L")
    Debug.Print row
End Sub


Function last_row_in_colm(ByVal sht_name As String, _
                        colm_ltr As String) _
                        As Long
    Dim colmNo As Long
    ThisWorkbook.Sheets(sht_name).Activate
    'convert the Column Letter to its numeric value
    colmNo = Range(colm_ltr & 1).Column
    
    last_row_in_colm = Cells(Rows.Count, colmNo).End(xlUp).row
End Function
