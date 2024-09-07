' filename:     scripting_dictionary.vba
'
' Purpose:      This is an example of how to use a scripting dictionary object
'
' Dependencies: Library - Microsoft Scripting Runtime required when early binding method is used
'               which is more efficient.  Early binding means that the compiler will
'               recognize the object's properties and methods at compile time.
'
' By:  T.Sciple, 09/06/2024

Option Explicit


Sub main()
    Dim dict As scripting.Dictionary 'Early Binding Method when New Keyword is used, requires reference library Microsoft Scripting Runtime
    Set dict = New scripting.Dictionary
    
    'Alternate Late Binding Syntax
    'Dim dict As Object
    'Set dict = CreateObject("Scripting.Dictionary")
    
    Call read_in_dict(dict)
    Call retrieve_item_fr_dict(dict)
    Call output_keys_and_items(dict)
    Call output_keys_only(dict)
    Call output_items_only(dict)
    ThisWorkbook.Sheets("Sheet1").Range("j37").Value = dict.Count
    Call check_if_key_exists(dict)
    Set dict = Nothing
End Sub


Sub read_in_dict(ByRef dict As Dictionary)
    Dim ary As Variant
    ary = ThisWorkbook.Sheets("Sheet1").Range("data")
    
    Dim i As Integer
    For i = LBound(ary, 1) To UBound(ary, 1)
        'syntax-> dict.add key, item
        dict.Add ary(i, 1), ary(i, 2)
    Next i
End Sub


Sub retrieve_item_fr_dict(ByRef dict As Dictionary)
    dict.compareMode = comparemethod.binarycompare
    Dim key_to_get As Variant
    key_to_get = ThisWorkbook.Sheets("Sheet1").Range("k3").Value
    ThisWorkbook.Sheets("Sheet1").Range("k4").Value = dict(key_to_get)
End Sub


Sub output_keys_and_items(ByRef dict As Dictionary)
    Dim start_row As Integer
    start_row = 7
    Dim key As Variant
    Dim i As Integer
    For Each key In dict.Keys
        ThisWorkbook.Sheets("Sheet1").Range("j" & start_row + i).Value = key
        ThisWorkbook.Sheets("Sheet1").Range("k" & start_row + i).Value = dict(key)
        i = i + 1
    Next key
End Sub


Sub output_keys_only(ByRef dict As Dictionary)
    Dim start_row As Integer
    start_row = 17
    Dim key As Variant
    Dim i As Integer
    For Each key In dict.Keys
        ThisWorkbook.Sheets("Sheet1").Range("j" & start_row + i).Value = key
        i = i + 1
    Next key
End Sub


Sub output_items_only(ByRef dict As Dictionary)
    Dim start_row As Integer
    start_row = 27
    Dim item As Variant
    Dim i As Integer
    For Each item In dict.Items
        ThisWorkbook.Sheets("Sheet1").Range("j" & start_row + i).Value = item
        i = i + 1
    Next item
End Sub


Sub check_if_key_exists(ByRef dict As Dictionary)
    Dim chk_item As Variant
    chk_item = ThisWorkbook.Sheets("Sheet1").Range("k40").Value
    ThisWorkbook.Sheets("Sheet1").Range("k41").Value = dict.Exists(chk_item)
End Sub


Sub Clear()
    ThisWorkbook.Sheets("Sheet1").Range("k4").ClearContents
    ThisWorkbook.Sheets("Sheet1").Range("j7:k14").ClearContents
    ThisWorkbook.Sheets("Sheet1").Range("j17:j24").ClearContents
    ThisWorkbook.Sheets("Sheet1").Range("j27:j34").ClearContents
    ThisWorkbook.Sheets("Sheet1").Range("j37").ClearContents
    ThisWorkbook.Sheets("Sheet1").Range("k41").ClearContents
End Sub