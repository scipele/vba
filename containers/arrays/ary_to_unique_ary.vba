Option Explicit
' filename:     ary_to_unique_ary.vba
' Purpose:      converts an existing array to a unique array
' Usage:        uniq_ary = ary_to_unique_ary(orig_ary)
' Dependencies: None
' By:  T.Sciple, 09/06/2024

Sub test()
    Dim my_ary As Variant
    my_ary = Array(1, 1, 2, 10, 11, 15, 3, 5, 10)   'Include duplicates in the original array that will be removed
    my_ary = ary_to_unique_ary(my_ary)
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
End Function