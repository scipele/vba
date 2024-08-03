Option Explicit

'  filename:    compare_lists.xlsm
'  Purpose:     Document Numbering Comparison Tool
'
'  Purpose:
'           Read two different lists (Table1 & Table2) and indicate the comparison of each list in Table 3
'           Table1: base list of items
'                     col A    col B
'                    ---------------------
'                    | id_1 |  list1     |
'                    ---------------------
'           Table2: second list of items to compare to the first list
'                     col D    col E
'                    ---------------------
'                    | id_2 |  list2     |
'                    ---------------------
'           Table3: table where the output of the comparison is shown
'                     col G    col H          col I     col J      col K
'                    --------------------------------------------------------
'                    | id_3 |  both_lists  |  list 1  |  list2  |  remarks  |
'                    --------------------------------------------------------
'  Dependencies: Uses microsoft sripting runtime library reference (dictionary objects early binding)
'
'  by jas, 8/2/2024


Public Sub Main()
    Dim sht_name As String
    sht_name = "compare"
    
    Dim tbl_a As String
    Dim tbl_b As String
    Dim tbl_c As String
    tbl_a = "Table1"
    tbl_b = "Table2"
    tbl_c = "Table3"
    
    'get List 1 and 2
    Dim list1_ary As Variant
    Dim list2_ary As Variant
    list1_ary = ReadTableFieldToArray(sht_name, tbl_a, "list1")
    list2_ary = ReadTableFieldToArray(sht_name, tbl_b, "list2")
    
    'Replace the arrays of list 1 and 2 with unique values using the sub 'get_unique_array'
    Call GetUniqueAry(list1_ary)
    Call GetUniqueAry(list2_ary)
    
    'Create the 'both_ary' array by calling the following sub where the arrays arguments are passed by reference
    Dim both_ary As Variant
    Call CombineUniqueItemsInArrays(list1_ary, list2_ary, both_ary)
    
    'Next sort the both array
    Call SortAryAtoZ(both_ary)
    
    'Next create a two dimensional array to hold the comparison of the list1/2
    Dim compare_ary As Variant
    Call GetComparisonArray(list1_ary, list2_ary, both_ary, compare_ary)
    
    'Cleanup sheet
    Call CleanupSht(sht_name, tbl_c)
    
    'Output the Array to comparison sheet
    Dim output_colm As String
    Dim output_top_row As Integer
    Dim output_format_code As String
    
    'Set the argument values for that are used in the Sub Call Below
    output_colm = "H"
    output_top_row = 2
    output_format_code = "TTTT"
    
    Call OutputAryToSheet(sht_name, output_colm, output_top_row, compare_ary, output_format_code)
End Sub


'Get List1 and make it unique
Function ReadTableFieldToArray(ByVal sht_name As String, _
                            ByVal tableName As String, _
                            ByVal fieldName As String) _
                            As Variant
    
    ' Declare a Range object to hold the column range
    Dim columnRange As Range
    
    ' Get the column range from the table
    With ThisWorkbook.Sheets(sht_name).ListObjects(tableName)
        Set columnRange = .ListColumns(fieldName).DataBodyRange
    End With
    
    ' Read the column values into the array
    ReadTableFieldToArray = columnRange.Value
End Function


Sub GetUniqueAry(ByRef orig_array As Variant)
'This function receives an array passed by reference

    Dim dict As Scripting.Dictionary
    Dim unique_ary As Variant
    Dim i As Long
    
    'Initialize the dictionary
    Set dict = New Scripting.Dictionary
    
    'Populate the dictionary with values from the array
    For i = LBound(orig_array, 1) To UBound(orig_array, 1)
        If Not dict.Exists(orig_array(i, 1)) Then
            dict.Add orig_array(i, 1), Nothing
        End If
    Next i
    
    'Transfer unique keys to an array
    unique_ary = dict.Keys
    
    'Next reset the original array to a unique one
    orig_array = unique_ary
    
    'Clear the dictionary
    dict.RemoveAll
    
    'Set the dictionary to Nothing
    Set dict = Nothing
    Erase unique_ary
    
End Sub


'Create the 'both_ary' array by calling the following sub where the arrays arguments are passeed by reference
Sub CombineUniqueItemsInArrays(ByRef list1_ary As Variant, _
                               ByRef list2_ary As Variant, _
                               ByRef combined_ary As Variant)

    Dim dict As Scripting.Dictionary
    Dim i As Long
    
    ' Initialize the dictionary
    Set dict = New Scripting.Dictionary
    
    ' Populate the dictionary with unique values from the array1
    For i = LBound(list1_ary) To UBound(list1_ary)
        If Not dict.Exists(list1_ary(i)) Then
            dict.Add list1_ary(i), Nothing
        End If
    Next i

    ' Populate the dictionary with unique values from the array2
    For i = LBound(list2_ary) To UBound(list2_ary)
        If Not dict.Exists(list2_ary(i)) Then
            dict.Add list2_ary(i), Nothing
        End If
    Next i

    ' Transfer unique keys to the array reference passed to the sub
    combined_ary = dict.Keys
    
    ' Clear the dictionary
    dict.RemoveAll
    
    ' Set the dictionary to Nothing
    Set dict = Nothing

End Sub


Private Sub SortAryAtoZ(ByRef both_ary As Variant)
    ' This is a basic bubble sort which is straight forward this is good enough relatively for small arrays,
    ' but consider more efficient sort algorithm for larger arrays
    Dim i As Long
    Dim j As Long
    Dim temp As Variant

    'Sort the Array A-Z
    For i = LBound(both_ary) To UBound(both_ary) - 1
        For j = i + 1 To UBound(both_ary)
            If UCase(both_ary(i)) > UCase(both_ary(j)) Then
                temp = both_ary(j)
                both_ary(j) = both_ary(i)
                both_ary(i) = temp
            End If
        Next j
    Next i
End Sub


'Next create a two dimensional array to hold the comparison of the list1/2
Public Sub GetComparisonArray(ByRef list1_ary As Variant, _
                         ByRef list2_ary As Variant, _
                         ByRef combined_ary As Variant, _
                         ByRef compare_ary As Variant)

    Dim dict1 As Scripting.Dictionary
    Dim dict2 As Scripting.Dictionary
    Dim i As Long
    
    ' Initialize the dictionary
    Set dict1 = New Scripting.Dictionary
    Set dict2 = New Scripting.Dictionary
    
    ' Populate the three dictionary objects with unique values from the arrays
    For i = LBound(list1_ary) To UBound(list1_ary)
        dict1.Add list1_ary(i), Nothing
    Next i

    For i = LBound(list2_ary) To UBound(list2_ary)
        dict2.Add list2_ary(i), Nothing
    Next i

    ' Setup the comparison array 'compare_ary' to store the values matched up from the different lists
    ReDim compare_ary(LBound(combined_ary) To UBound(combined_ary), 0 To 3)

    For i = LBound(compare_ary, 1) To UBound(compare_ary, 1)
        compare_ary(i, 0) = combined_ary(i)
        
        'set the 2nd array dimension if there is a matching value in the list1
        If dict1.Exists(combined_ary(i)) Then
            compare_ary(i, 1) = combined_ary(i)
        End If
        
        'set the 3rd array dimension if there is a matching value in the list1
        If dict2.Exists(combined_ary(i)) Then
            compare_ary(i, 2) = combined_ary(i)
        End If
        
        'set the remarks or 4th element in the array equal to 'same' 'list1' 'list2'
        If compare_ary(i, 1) = compare_ary(i, 2) Then
            compare_ary(i, 3) = "same"
        Else
            If compare_ary(i, 1) = "" Then
                compare_ary(i, 3) = "list2"
            End If
            
            If compare_ary(i, 2) = "" Then
                compare_ary(i, 3) = "list1"
            End If
        End If
    Next i

    ' Clear the dictionary
    dict1.RemoveAll
    dict2.RemoveAll
    
    ' Set the dictionary to Nothing
    Set dict1 = Nothing
    Set dict2 = Nothing
End Sub


'Cleanup sheet
Sub CleanupSht(ByVal sht_name As String, _
                  ByVal tbl_name As String)

    Dim tbl As ListObject
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sht_name) ' Adjust the sheet name as necessary
    Set tbl = ws.ListObjects(tbl_name)

    ' Check if the table has more than one data row
    If tbl.ListRows.count > 1 Then
        ' Delete all rows except the first data row
        tbl.DataBodyRange.Offset(1, 0).Resize(tbl.ListRows.count - 1).Rows.Delete
    End If
End Sub


Public Sub OutputAryToSheet(ByVal sht_name As String, _
                            ByVal colm_ltr As String, _
                            ByVal row_no_top As Integer, _
                            ByVal tmp_ary As Variant, _
                            ByVal colm_format As String)
    
    'Places values from two dimensional array to a worksheet
    'Usage-> Call OutputAryToSheet("sht1", "D", "2", dwgList)
    Dim ary_colms, ary_btm_row, col_start_no, col_end_no As Integer
    Dim col_end_ltr As String
    
    'determine the number of columns and rows based on the array dimensions
    ary_btm_row = row_no_top + UBound(tmp_ary, 1) - LBound(tmp_ary, 1)
    ary_colms = UBound(tmp_ary, 2) - LBound(tmp_ary, 2)
    
    'determine the top, bottom row numbers and the start and ending column letters
    col_start_no = Range(colm_ltr & 1).Column
    col_end_no = col_start_no + ary_colms
    col_end_ltr = Split(Cells(1, col_end_no).Address, "$")(1)
    Dim out_rng_str As String
    out_rng_str = colm_ltr & row_no_top & ":" & col_end_ltr & ary_btm_row
    
    'call sub to format columns
    Call SetColFormat(sht_name, col_start_no, col_end_no, row_no_top, ary_btm_row, colm_format)
    
    Dim rng_target As Range
    Set rng_target = ThisWorkbook.Worksheets(sht_name).Range(out_rng_str)
    rng_target = tmp_ary
End Sub


Sub SetColFormat(ByVal sht_name, _
                   ByVal col_start_no As Integer, _
                   ByVal col_end_no As Integer, _
                   ByVal row_no_top As Integer, _
                   ByVal ary_btm_row As Integer, _
                   ByVal colm_format As String)
    
    'dim variables before looping
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sht_name)
    Dim my_rng As Range
    Dim cur_col_str As String       'T = Text, N = Number, D = Date Time, d - Date Only
    Dim cur_col_code As String
    Dim cur_col_ltr As String
    Dim i As Integer
    Dim index As Long
    index = 1                       'index used for each char in the 'colm_format' string
    
    For i = col_start_no To col_end_no
        cur_col_str = Mid(colm_format, index, 1)
        index = index + 1
        cur_col_code = Switch(cur_col_str = "T", "@", cur_col_str = "N", "0", cur_col_str = "D", "mm/dd/yyyy hh:mm:ss", cur_col_str = "d", "mm/dd/yyyy")
        cur_col_ltr = Split(Cells(1, i).Address, "$")(1)
        Set my_rng = ws.Range(cur_col_ltr & row_no_top & ":" & cur_col_ltr & ary_btm_row)
        my_rng.NumberFormat = cur_col_code
    Next i
End Sub


Public Sub Cleanup(ByVal numStr As String)

    Dim sht_name As String
    sht_name = "compare"
    
    Dim tbl_a As String
    tbl_a = "Table" & numStr

    Call CleanupSht(sht_name, tbl_a)
    
    Dim rng_str As String
    'set the column to clear depending on the call 'numStr'
    rng_str = Switch(numStr = "1", "B2", numStr = "2", "E2", numStr = "3", "h2:k2")
    
    ThisWorkbook.Sheets(sht_name).Range(rng_str).Value = ""
End Sub
