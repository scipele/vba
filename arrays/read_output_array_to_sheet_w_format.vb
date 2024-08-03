' filename - read_output_array_to_sheet_w_format
'
'
' 1.  This sub provides a method to read a named range 'input_rng' into an array
' 2.  Next titles are added to the output location based on an array of strings
' 3.  Finally the OutputAryToSheet is called which formats the cells and then sets the target range equal to the contents of the array
'
' By Tony Sciple, 8/2/2024 scipele@yahoo.com


Sub main()
    
    'Call Sub to read data from range into array
    Dim tmp_ary  As Variant
    Dim input_rng_str As String
    input_rng_str = "input_rng"
    Call GetArrayData(tmp_ary, input_rng_str)
    
   'add titles
    Dim title_ary As Variant
    title_ary = Array("Filename", "Folder Name", "size", "Type", "Date", "fileSHA1Hash", "Count")
    Call Transpose1DArrayto2D(title_ary)
    Call OutputAryToSheet("output", "A", 5, title_ary, "TTTTTTT")
    
    'Output Array to Output Sheet
    Call OutputAryToSheet("output", "A", 6, tmp_ary, "TTNTDTN")
    
    'cleanup
    Erase tmp_ary
    ThisWorkbook.Sheets("output").Columns("A:G").AutoFit

End Sub


Sub GetArrayData(ByRef tmp_ary As Variant, _
                 ByVal input_rng_str As String)

    ' Set the named range
    Dim input_rng As Range
    Set input_rng = ThisWorkbook.Names(input_rng_str).RefersToRange
    
    'Read the Range Object into the variant Array that was passed by reference
    tmp_ary = input_rng
    
End Sub


Private Sub Transpose1DArrayto2D(ByRef title_ary As Variant)
    Dim tmp_ary As Variant
    ReDim tmp_ary(0 To 0, LBound(title_ary) To UBound(title_ary))
    
    Dim i As Long
    For i = LBound(title_ary) To UBound(title_ary)
        tmp_ary(0, i) = title_ary(i)
    Next i
    
    'reset the title_ary equal to the transposed tmp_ary
    title_ary = tmp_ary
    
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
    Call setcolm_format(sht_name, col_start_no, col_end_no, row_no_top, ary_btm_row, colm_format)
    
    Dim rng_target As Range
    Set rng_target = ThisWorkbook.Worksheets(sht_name).Range(out_rng_str)
    rng_target = tmp_ary

End Sub


Sub set_col_format(ByVal sht_name, _
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
