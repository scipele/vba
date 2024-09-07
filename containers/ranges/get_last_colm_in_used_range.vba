Option Explicit
' filename:     get_last_colm_in_used_range.vba
' Purpose:      Finds the last column in the used worksheet range
' Usage:        col = get_last_colm(sht)
' Dependencies: None
' By:  T.Sciple, 09/07/2024
Public Sub test()
    Dim sht As String
    sht = "Sheet1"
    Dim col As String
    col = get_last_colm(sht)
    Debug.Print col
End Sub


Public Function get_last_colm(ByVal shtName As String) As String
    Worksheets(shtName).Activate
    Dim tmp_ary As Variant
    tmp_ary = ActiveSheet.UsedRange
    
    Dim last_colm_no As Long
    last_colm_no = UBound(tmp_ary, 2)
    Erase tmp_ary

    get_last_colm = Split(Cells(1, last_colm_no).Address, "$")(1)
End Function
