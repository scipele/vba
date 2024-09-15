' filename:     rng_to_ary_2d.vba
'
' Purpose:      converts a range in a worksheet to a two dimensional array
'
' Usage:        rng_to_ary_2d("Sheet1", "A2:B5", base_no)
' parameters:
'               sht_name As String
'               rng_str As String
'               base_num As Integer  ( 0 for zero Base index, 1 for 1 Base index)
'
' Dependencies: None
'
' By:  T.Sciple, 09/06/2024


Sub clear()
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets("Sheet1").Range("A12:B16")
    rng.ClearContents
    Set rng = Nothing
End Sub


Sub test()
    'test with top to bottom range
    Dim base_no As Integer
    base_no = 1
    
    Dim my_ary As Variant
    my_ary = rng_to_ary_2d("Sheet1", "A2:C5", base_no)
    
    Dim row_start As Long
    row_start = 11
    
    Dim col_ltr As String
    Dim colm As Integer
    
    Dim row_cnt As Integer
    row_cnt = UBound(my_ary, 1) - LBound(my_ary, 1) + 1
    
    Dim col_cnt As Integer
    col_cnt = UBound(my_ary, 2) - LBound(my_ary, 2) + 1
    
    Dim i As Long
    Dim j As Long
    For i = 1 To row_cnt
        For j = 1 To col_cnt
            col_ltr = colm_no_to_letter(j)
            ThisWorkbook.Sheets("Sheet1").Range(col_ltr & i + row_start).Value = my_ary(i - base_no + 1, j - base_no + 1)
        Next j
    Next i

    'cleanup
    Erase my_ary
End Sub


Private Function rng_to_ary_2d(sht_name As String, _
                               rng_str As String, _
                               base_num As Integer) _
                               As Variant
    
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets(sht_name).Range(rng_str)

    ' Get the number of rows and columns
    Dim row_count As Long
    row_count = rng.Rows.Count
    Dim col_count As Long
    col_count = rng.Columns.Count
    
    'initialize an array and size to hold the range
    Dim offset As Integer
    offset = base_num - 1
    
    Dim tmp_ary As Variant
    ReDim tmp_ary(base_num To row_count + offset, base_num To col_count + offset)
    
    ' Iterate through each row in the range, note that ranges are all base 1
    For i = 1 To row_count
        For j = 1 To col_count
            tmp_ary(i + offset, j + offset) = rng(i, j)
        Next j
    Next i
    
    'set function return value equal to the temporary array
    rng_to_ary_2d = tmp_ary
    
    'cleanup
    Erase tmp_ary
    Set rng = Nothing
End Function


Function colm_no_to_letter(ByVal n As Long) As String
    Dim result As String
    Dim ascA As Integer
    
    If n > 0 And n < 16385 Then
        ascA = 64  ' ASCII value of 'A' is 65, so subtract 1 to use it as a base
        
        Do While n > 0
            n = n - 1  ' Adjusting for 1-based index of Excel columns
            result = Chr(ascA + (n Mod 26) + 1) & result
            n = n \ 26  'integer division operator returns the integer quotient
        Loop
    Else
        result = "Error - Invalid column number"
    End If
    
    colm_no_to_letter = result
End Function
