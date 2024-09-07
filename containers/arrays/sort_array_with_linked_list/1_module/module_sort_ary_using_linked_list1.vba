Option Explicit
' filename:     sort_ary_using_linked_list.vba
' Purpose:      sorts an array using a LinkedListSort and Node Classes
' Usage:        ary = sort_ary_using_linked_list(ary)
'
' Dependencies - Class Modules:
'
'               Class Module - LinkedListSort
'                   Public Sub - InsertInOrder method takes each value and creates a sorted object from head, and assigns nextNode(s) in order
'                   Public Function - return_sorted_ary
'
'               Class Node - This class is used in the LinkedListSort class in order to store the head and order of the list values
'
' By:  T.Sciple, 09/06/2024


Sub main()
    Dim ary As Variant
    ary = ThisWorkbook.Sheets("Sheet1").Range("inp_rng")
    ary = sort_ary_using_linked_list(ary)
    
    Call output_1d_array_to_rng(ary)
    'cleanup
    Erase ary
End Sub


Private Function sort_ary_using_linked_list(ByRef ary As Variant)
    Dim link_list As LinkedListSort
    Set link_list = New LinkedListSort
    
    ' Insert items into the linked list in sorted order
    Dim i As Long
    For i = LBound(ary, 1) To UBound(ary, 1)
        link_list.InsertInOrder ary(i, 1)
    Next i

    'reset the array equal to temp array created within the LinkedListSort Class
    sort_ary_using_linked_list = link_list.return_sorted_ary
End Function


Sub output_1d_array_to_rng(ByRef ary As Variant)
    'convert the array to 2d so that the range property can be set
    ary = ary_1d_to_2d(ary)
    
    Dim start_row As Integer
    start_row = 5
    Dim count As Long
    count = UBound(ary, 1) - LBound(ary, 1)
    
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets("Sheet1").Range("d" & start_row & ":d" & start_row + count)
    rng = ary

    'cleanup
    Set rng = Nothing
End Sub


Function ary_1d_to_2d(ByRef ary As Variant) As Variant
    Dim i As Long
    Dim tmp_ary As Variant
    ReDim tmp_ary(LBound(ary) To UBound(ary), 1 To 1)
    
    For i = LBound(ary) To UBound(ary)
        tmp_ary(i, 1) = ary(i)
    Next i
    ary_1d_to_2d = tmp_ary
End Function


Sub clear()
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets("Sheet1").Range("d5:d1006")
    rng.ClearContents
    'cleanup
    Set rng = Nothing
End Sub