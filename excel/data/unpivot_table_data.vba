Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | unpivot_table_data.vba                                      |
'| EntryPoint   | unpivotTableData                                            |
'| Purpose      | unpivot paired columns in an excel table                    |
'| Inputs       | Reads Worksheet Object Table1                               |
'| Outputs      | unpivoted data placed on sheet named 'output'               |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 8/12/2025                                         |


Sub unpivotTableData()
    'set the worksheet and table (change "sheet1" and "table1" to your sheet and table names)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Set ws = ThisWorkbook.Worksheets("entry")
    Set tbl = ws.ListObjects("Table1")
    
    'get the header row
    Dim headerRow As Range
    Set headerRow = tbl.HeaderRowRange
    
    'indicate the column numbers pairs to unpivot
    Dim pair_ary As Variant
    pair_ary = Array(2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
    
    'loop through the table rows and populate the data array
    Dim row As ListRow
    Dim cell_value As Variant
    Dim val_a As Variant
    Dim val_b As Variant
    Dim indx As Long
    
    'first run thru to count the size needed for the array
    For Each row In tbl.ListRows
       For j = 0 To UBound(pair_ary) Step 2
           pair_a_indx = pair_ary(j)
           pair_b_indx = pair_ary(j + 1)
           
           'keep the id for reference
           val_a = tbl.DataBodyRange.Cells(row.Index, pair_a_indx).Value
           val_b = tbl.DataBodyRange.Cells(row.Index, pair_b_indx).Value
           
           If val_a <> "" Or val_b <> "" Then
               indx = indx + 1
           End If
       Next j
    Next row
    
    'now read in the data
    Dim data As Variant
    ReDim data(1 To indx, 0 To 2)
    
    'reset the indx
    indx = 0
    For Each row In tbl.ListRows
        For j = 0 To UBound(pair_ary) Step 2
            pair_a_indx = pair_ary(j)
            pair_b_indx = pair_ary(j + 1)
            
            'keep the id for reference
            val_a = tbl.DataBodyRange.Cells(row.Index, pair_a_indx).Value
            val_b = tbl.DataBodyRange.Cells(row.Index, pair_b_indx).Value
            
            If val_a <> "" Or val_b <> "" Then
                indx = indx + 1
                data(indx, 0) = tbl.DataBodyRange.Cells(row.Index, 1).Value
                data(indx, 1) = tbl.DataBodyRange.Cells(row.Index, pair_a_indx).Value
                data(indx, 2) = tbl.DataBodyRange.Cells(row.Index, pair_b_indx).Value
            End If
        Next j
    Next row
    
    Call output_ary_to_sht("output", "A", 2, data, "NNn")
    
    'cleanup
    Erase data
    Set ws = Nothing
    Set tbl = Nothing
    
End Sub


Private Sub output_ary_to_sht( _
    ByVal sht_name As String, _
    ByVal colm_loc As String, _
    ByVal row_top_loc As Integer, _
    ByRef tmp_ary As Variant, _
    ByVal colm_format As String)
    
    'Places values from two dimensional array to a worksheet
    'Usage-> Call output_ary_to_sht("sht1", "D", "2", dwgList, )

    Dim rng_target As Range
    Dim ary_colms, ary_btm_row, start_colm_num, end_colm_num As Integer
    Dim end_colm_ltr As String
    
    ary_btm_row = row_top_loc + UBound(tmp_ary, 1) - LBound(tmp_ary, 1)
    ary_colms = UBound(tmp_ary, 2) - LBound(tmp_ary, 2)
    
    start_colm_num = colm_ltr_to_no(colm_loc)
    end_colm_num = start_colm_num + ary_colms
    end_colm_ltr = colm_no_to_ltr(end_colm_num)
    
    'call sub to format columns
    Call set_colm_format(sht_name, start_colm_num, end_colm_num, row_top_loc, ary_btm_row, colm_format)
    
    'Next Set the range values from the 2D Array
    Set rng_target = ActiveWorkbook.Worksheets(sht_name).Range(colm_loc & row_top_loc & ":" & end_colm_ltr & ary_btm_row)
    rng_target = tmp_ary
End Sub


Private Sub set_colm_format( _
    ByVal sht_name, _
    ByVal start_colm_num As Integer, _
    ByVal end_colm_num As Integer, _
    ByVal row_top_loc As Integer, _
    ByVal ary_btm_row As Integer, _
    ByVal colm_format As String)
    
    Dim myRng As Range
    Dim i As Integer
    Dim cur_colm_str As String    ' T = Text, N = Number, D = Date Time, d - Date Only
    Dim cur_colm_code As String
    Dim cur_colm_ltr As String
    
    ' Set the worksheet where your data is located
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sht_name)
    
    For i = start_colm_num To end_colm_num
        cur_colm_str = Mid(colm_format, i, 1)
        ' set Format codes as follows T = @, N = "0", "D" = "mm/dd/yyyy hh:mm:ss", "d" = "mm/dd/yyyy"
        cur_colm_code = Switch(cur_colm_str = "T", "@", cur_colm_str = "N", "0", cur_colm_str = "n", "0.000", cur_colm_str = "D", "mm/dd/yyyy hh:mm:ss", cur_colm_str = "d", "mm/dd/yyyy")
        cur_colm_ltr = colm_no_to_ltr(i)
        
        Set myRng = ws.Range(cur_colm_ltr & row_top_loc & ":" & cur_colm_ltr & ary_btm_row)
        myRng.NumberFormat = cur_colm_code
    Next i
End Sub


Private Function colm_ltr_to_no(ByVal ColLtr As String)
    colm_ltr_to_no = Range(ColLtr & 1).Column
End Function


Private Function colm_no_to_ltr(ByVal colmNo As Long)
    colm_no_to_ltr = Split(Cells(1, colmNo).Address, "$")(1)
End Function