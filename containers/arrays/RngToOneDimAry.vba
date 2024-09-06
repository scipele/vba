' filename:     RngToOneDimAry.vba
'
' Purpose:      converts a range in a worksheet to a one dimensional array
'
' Usage:        my_ary = rng_to_one_dim_array("Sheet1", "A2:A5", 0)
' parameters:
'               sht_name As String
'               rng_str As String
'               base_num As Integer  ( 0 for zero Base index, 1 for 1 Base index)
'
' Dependencies: None
'
' By:  T.Sciple, 09/06/2024


Sub test()
    'test with top to bottom range
    Dim my_ary As Variant
    my_ary = rng_to_one_dim_ary("Sheet1", "A2:A5", 0) '<--  Note that the zero produces a zero based array
    
    Dim elem As Variant
    For Each elem In my_ary
        Debug.Print elem
    Next elem

    'test with left to right range
    Dim my_ary2 As Variant
    my_ary2 = rng_to_one_dim_ary("Sheet1", "e2:h2", 1) '<--  Note that the 1 produces a 1 based array
    
    For Each elem In my_ary2
        Debug.Print elem
    Next elem

    'cleanup
    Erase my_ary
    Erase my_ary2
End Sub


Private Function rng_to_one_dim_ary( _
                                    sht_name As String, _
                                    rng_str As String, _
                                    base_num As Integer) _
                                    As Variant

    Dim rng As Range
    Set rng = ThisWorkbook.Sheets(sht_name).Range(rng_str)
    
    'Dimension and resize a temporary one dimensional array to match the size of the range
    Dim tmp_ary As Variant
    ReDim tmp_ary(base_num To rng.Count + base_num - 1)

    'Read the range into the array
    Dim item As Variant
    i = base_num
    For Each item In rng
        tmp_ary(i) = item
        i = i + 1
    Next
    
    'set the function return equal to the variant temporary array 'tmp_ary'
    rng_to_one_dim_ary = tmp_ary
End Function