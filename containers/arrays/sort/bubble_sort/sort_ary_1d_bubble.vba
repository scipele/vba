Option Explicit
' filename:     sort_ary_1d_bubble.vba
'
' Purpose:      sorts an array using bubble sort algorithmn which is O(n^2) Time Complexity
'                so only good for small arrays
'
' key Usage:    sort_ary_id(ary, "asc") '<- or desc
'               other items are for reading and output
'
' Dependencies: None
' By:  T.Sciple, 09/07/2024


Sub main()
    '1. read range to array
    Dim ary As Variant
    ary = ThisWorkbook.Sheets("Sheet1").Range("inp_rng")
    
    '2. convert array to 1d
    ary = rng_to_ary_1d("Sheet1", "B4:B12", 1)

    '3. sort the array
    Call sort_ary_id(ary, "asc")
    
    '4. output to sheet if not erased
    If IsArrayErased(ary) Then
        Exit Sub
    Else
        Call output_1d_array_to_rng(ary)
    End If

    'cleanup
    Erase ary
End Sub


Private Sub sort_ary_id(ByRef ary As Variant, sort_order As String)

    Dim i As Long
    Dim j As Long
    Dim Temp
    
    Dim asc_bool As Boolean
    Select Case sort_order
    
    Case "asc"
        asc_bool = True
    Case "desc"
        asc_bool = False
    Case Else
        GoTo err:
    End Select
    

    'Sort the Array A-Z
    For i = LBound(ary) To UBound(ary) - 1
        For j = i + 1 To UBound(ary)
        
            If sort_order = "asc" Then
                If UCase(ary(i)) > UCase(ary(j)) Then
                    Temp = ary(j)
                    ary(j) = ary(i)
                    ary(i) = Temp
                End If
            ElseIf sort_order = "desc" Then
                If UCase(ary(i)) < UCase(ary(j)) Then
                    Temp = ary(j)
                    ary(j) = ary(i)
                    ary(i) = Temp
                End If
            End If
        Next j
    Next i
    
    Exit Sub    'exit normally if not error
err:
    MsgBox ("invalid sort order parameter")
    Erase ary
    
End Sub


Private Function rng_to_ary_1d(sht_name As String, _
                                rng_str As String, _
                                base_num As Integer) _
                                As Variant

    Dim rng As Range
    Set rng = ThisWorkbook.Sheets(sht_name).Range(rng_str)
    
    'Dimension and resize a temporary one dimensional array to match the size of the range
    Dim tmp_ary As Variant
    ReDim tmp_ary(base_num To rng.count + base_num - 1)

    'Read the range into the array
    Dim item As Variant
    Dim i As Long
    i = base_num
    For Each item In rng
        tmp_ary(i) = item.Value
        i = i + 1
    Next
    
    'set the function return equal to the variant temporary array 'tmp_ary'
    rng_to_ary_1d = tmp_ary
End Function


Function IsArrayErased(ByVal arr As Variant) As Boolean
    On Error Resume Next ' Suppress errors
    IsArrayErased = (LBound(arr) > UBound(arr)) ' Check if the array bounds are invalid
    If err.Number <> 0 Then
        IsArrayErased = True ' If there's an error, the array is erased
    End If
    On Error GoTo 0 ' Reset error handling
End Function


Sub output_1d_array_to_rng(ByRef ary As Variant)
    'convert the array to 2d so that the range property can be set
    ary = Convert1DTo2D(ary)
    
    Dim start_row As Integer
    start_row = 4
    Dim count As Long
    count = UBound(ary, 1) - LBound(ary, 1)
    
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets("Sheet1").Range("E" & start_row & ":e" & start_row + count)
    rng = ary

    'cleanup
    Set rng = Nothing
End Sub


Function Convert1DTo2D(ByRef ary As Variant) As Variant
    Dim i As Long
    Dim tmp_ary As Variant
    ReDim tmp_ary(LBound(ary, 1) To UBound(ary, 1), 1 To 1)
    
    For i = LBound(ary) To UBound(ary)
        tmp_ary(i, 1) = ary(i)
    Next i
    Convert1DTo2D = tmp_ary
End Function


Sub clear()
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets("Sheet1").Range("e4:E12")
    rng.ClearContents
    'cleanup
    Set rng = Nothing
End Sub