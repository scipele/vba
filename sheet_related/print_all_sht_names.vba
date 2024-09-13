' Sub:     print_all_sht_names()
'
' Purpose:  1. This code loops thru each sheet and prints the name of each sheet to the immediate window
'           2. Also it concatenates all the sheet names together with quotes and comma between them
'
'
' Dependencies:  None
'
' By:  T. Sciple, 8/8/2024

Sub print_all_sht_names()
    ' Loop thru each sheet
    dim str as String
    Dim ws As Variant
    For Each ws In ThisWorkbook.Worksheets
        Debug.Print ws.Name
        str = str & """" & ws.Name & """" & ","
    Next ws

    ' Remove last comma from the concatenated string
    str = left(str, len(str) - 1)
    Debug.Print ws.Name
End Sub