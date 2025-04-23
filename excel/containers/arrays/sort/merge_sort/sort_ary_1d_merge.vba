Option Explicit
' filename:     sort_ary_1d_merge.vba
'
' Purpose:      sorts an array using merge sort algorithmn which is average time complexity of O(n log n)
'
' key Usage:    Call mergeSort(ary, left, n - 1)
'               left = 0 to start
'               n = number of items in array
'
' Dependencies: None
' By:  T.Sciple, 09/09/2024


Sub main()
    '1. read range to array
    Dim ary As Variant
    ary = ThisWorkbook.Sheets("Sheet1").Range("inp_rng")
    
    '2. convert array to 1d
    ary = rng_to_ary_1d("Sheet1", "inp_rng", 0)

    '3. sort the array
    Dim n As Long
    n = UBound(ary) - LBound(ary) + 1
    
    Dim left As Long
    left = 0
    
    Call mergeSort(ary, left, n - 1)
    
    '4. output to sheet if not erased
    If IsArrayErased(ary) Then
        Exit Sub
    Else
        Call output_1d_array_to_rng(ary)
    End If

    'cleanup
    Erase ary
End Sub


Sub mergeSort(ByRef ary As Variant, _
              ByVal left As Long, _
              ByVal right As Long)

    If (left >= right) Then
        Exit Sub
    End If
    
    Dim mid As Long
    mid = left + (right - left) \ 2
    Call mergeSort(ary, left, mid)
    Call mergeSort(ary, mid + 1, right)
    Call merge(ary, left, mid, right)
End Sub

' Merges two subvectors of vec[]
' First subarray is vec[left..mid]
' Second subarray is vec[mid+1..right]
Sub merge(ByRef ary As Variant, left As Long, mid As Long, right As Long)

    Dim n1 As Long
    n1 = mid - left + 1
    Dim n2 As Long
    n2 = right - mid

    ' Create temp vectors
    Dim L() As Variant
    ReDim L(0 To n1 - 1)
    Dim R() As Variant
    ReDim R(0 To n2 - 1)
    
    ' Copy data to temp vectors L[] and R[]
    Dim i As Long
    For i = 0 To (n1 - 1)
        L(i) = ary(left + i)
    Next i
    
    Dim j As Long
    For j = 0 To (n2 - 1)
        R(j) = ary(mid + 1 + j)
    Next j

    i = 0: j = 0
    Dim k As Long
    k = left
    
    ' Merge the temp vectors back into vec[left..right]
        
    While (i < n1 And j < n2)
        ' Following is where comparison and swap is done
        If L(i) <= R(j) Then
            ary(k) = L(i)
            i = i + 1
        Else
            ary(k) = R(j)
            j = j + 1
        End If
        k = k + 1
    Wend

    ' Copy the remaining elements of L[], if there are any
    While (i < n1)
        ary(k) = L(i)
        i = i + 1
        k = k + 1
    Wend

    ' Copy the remaining elements of R[], if there are any
    While (j < n2)
        ary(k) = R(j)
        j = j + 1
        k = k + 1
    Wend

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
    Set rng = ThisWorkbook.Sheets("Sheet1").Range("c" & start_row & ":c" & start_row + count)
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
    Set rng = ThisWorkbook.Sheets("Sheet1").Range("c4:c508")
    rng.ClearContents
    'cleanup
    Set rng = Nothing
End Sub