Option Explicit
' filename:     colm_ltr_to_no.vba
'
' Purpose:      returns the excel column letter converted to a number
'
' Usage:        col_no = colm_ltr_to_no("AB")
'
' Dependencies: None
' By:  T.Sciple, 09/07/2024
Public Sub test()
    Dim col_no As Integer
    col_no = colm_ltr_to_no("AB")
    Debug.Print col_no
End Sub

Public Function colm_ltr_to_no(ByVal col_ltr As String)
   colm_ltr_to_no = Range(col_ltr & 1).Column
End Function