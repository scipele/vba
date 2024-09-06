Option Explicit
' filename:     ary_to_unique_ary.vba
' Purpose:      converts an existing array to a unique array
' Usage:        uniq_ary = ary_to_unique_ary(orig_ary)
' Dependencies: None
' By:  T.Sciple, 09/06/2024


Sub main()
    Dim my_ary As Variant
    my_ary = ThisWorkbook.Sheets("Sheet1").Range("inp_rng")
    my_ary = ary_to_unique_ary(my_ary)
    Call output_1d_array_to_rng(my_ary)

    'cleanup
    erase my_ary
End Sub


Function ary_to_unique_ary(ByRef my_ary As Variant) As Variant
    'Create a dictionary object using early binding
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim elem As Variant
    For Each elem In my_ary
        If Not dict.exists(elem) Then
            dict.Add elem, Nothing  '<- Only 'key' part of the dict object is needed, therefore 'item' is set to nothing
        End If
    Next elem
    
    ary_to_unique_ary = dict.keys

    'cleanup
    set dict = nothing
End Function


Function Convert1DTo2D(ByRef ary As Variant) As Variant
    Dim i As Long
    Dim newArr() As Variant
    ReDim newArr(LBound(ary) To UBound(ary), 1 To 1)
    
    For i = LBound(ary) To UBound(ary)
        newArr(i, 1) = ary(i)
    Next i
    Convert1DTo2D = newArr
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
    set rng = nothing
End Sub


Sub clear()
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets("Sheet1").Range("e4:E10")
    rng.ClearContents
    'cleanup
    set rng = nothing
End Sub