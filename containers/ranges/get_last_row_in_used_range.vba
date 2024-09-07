Option Explicit
' filename:     get_last_row_in_used_range.vba
' Purpose:      Finds the last row in the used worksheet range
' Usage:        row = get_last_row(sht)
' Dependencies: None
' By:  T.Sciple, 09/07/2024
Public Sub test()
    Dim sht As String
    sht = "Sheet1"
    Dim row As Long
    row = get_last_row(sht)
    Debug.Print row
End Sub


Public Function get_last_row(ByVal shtName As String) As String
    Worksheets(shtName).Activate
    Dim tmp_ary As Variant
    tmp_ary = ActiveSheet.UsedRange
    
    Dim last_colm_no As Long
    get_last_row = UBound(tmp_ary, 1)
    Erase tmp_ary
End Function