Option Explicit
' filename:         remove_elem_from_ary.vba
' Purpose:          removes a specified element(s) from an existing array passed to the sub by reference
' example usage:    Call rem_elem_from_ary(my_ary, elem_to_remove)
' Dependencies:     None
' By:               T.Sciple, 09/06/2024


Sub main()
    Dim ary As Variant
    ary = ThisWorkbook.Sheets("Sheet1").Range("inp_rng")
    Dim elem_to_remove As String
    elem_to_remove = "-"
    Call remove_elem_from_ary(ary, elem_to_remove)
    'send output to sheet
    Call output_1d_array_to_rng(ary)
    'cleanup
    Erase ary
End Sub


Private Sub remove_elem_from_ary(ByRef ary As Variant, ByVal elem_to_remove As String)
    'loop thru the array and count the number of items to remove
    Dim rem_count As Long
    Dim elem As Variant
    For Each elem In ary
        If elem = elem_to_remove Then
            rem_count = rem_count + 1
        End If
    Next elem
    Dim count As Long
    count = UBound(ary) - LBound(ary) - rem_count
    Dim base As Integer
    base = LBound(ary)

    'Now that you know the count them read values into the temp array
    Dim tmp_ary
    ReDim tmp_ary(base To (count + base), 1 To 1) '<- force it to a two dimensional array so that range property can by used
    Dim i As Long
    i = base
    For Each elem In ary
        If Not elem = elem_to_remove Then
            tmp_ary(i, 1) = elem
            i = i + 1
        End If
    Next elem
    'now set the contents of the original array 'ary' passed by reference to the tmp_ary
    ary = tmp_ary
    Erase tmp_ary
End Sub


Sub output_1d_array_to_rng(ByRef ary As Variant)
    Dim start_row As Integer
    start_row = 4
    Dim count As Long
    count = UBound(ary, 1) - LBound(ary, 1)
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets("Sheet1").Range("e" & start_row & ":e" & start_row + count)
    rng = ary
    'cleanup
    Set rng = Nothing
End Sub


Sub clear()
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets("Sheet1").Range("e4:e10")
    rng.ClearContents
    'cleanup
    Set rng = Nothing
End Sub