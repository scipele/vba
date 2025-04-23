Option Explicit
' filename:     colm_no_to_ltr.vba
'
' Purpose:      returns the excel column letter from long numeric input parameter
'
' Usage:        col_ltr = colm_no_to_ltr(col_no)
'
' Dependencies: None
' By:  T.Sciple, 09/07/2024
Public Sub test()
    Dim col_no As Long
    col_no = 28
    Dim col_ltr As String
    col_ltr = colm_no_to_ltr(col_no)
    Debug.Print col_ltr
End Sub

Private Function colm_no_to_ltr(ByVal colm_no As Long)
    colm_no_to_ltr = Split(Cells(1, colm_no).Address, "$")(1)
End Function