' filename:     rng_to_ary_and_output_to_rng_formatted.vba
'
' Purpose:      1. Reads range into a 2d array
'               2. Output the array data onto the sheet, with format codes
'               3. Output title data onto the sheet, with format codes
'
'               Note that Format Characters passed to the used in the T = Text, N = Number, D = Date Time, d - Date Only
'               passed as an argument to the 'colm_format' parameter in the sub 'set_colm_format'
'
' Dependencies: Noine
'
' By:  T.Sciple, 09/06/2024

Option Explicit


Sub main()
    '1. Read data from range into array
    Dim ary As Variant
    ary = ThisWorkbook.Sheets("data").Range("A2:G28")
    
    '2. Output Array to Output Sheet
    Call output_ary_to_sht("output", "A", 6, ary, "TTNTDTN")
    
   '3. add titles
    Dim title_ary As Variant
    title_ary = Split("Filename|Folder Name|size|Type|Date|fileSHA1Hash|Count", "|")
    title_ary = transpose_1D_ary_to_2D(title_ary)
    Call output_ary_to_sht("output", "A", 5, title_ary, "TTTTTTT")
    
    '4 cleanup column width
    ThisWorkbook.Sheets("output").Range("A:G").Columns.AutoFit
    
    Erase ary
End Sub


Public Sub clear()
    ThisWorkbook.Sheets("output").Range("A5:G32").ClearContents
End Sub


Private Function transpose_1D_ary_to_2D(ByVal ary As Variant) As Variant
    Dim tmp_ary As Variant
    ReDim tmp_ary(0 To 0, LBound(ary) To UBound(ary))
    Dim i As Long
    For i = LBound(ary) To UBound(ary)
        tmp_ary(0, i) = ary(i)
    Next i
    
    transpose_1D_ary_to_2D = tmp_ary
End Function


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
        cur_colm_code = Switch(cur_colm_str = "T", "@", cur_colm_str = "N", "0", cur_colm_str = "D", "mm/dd/yyyy hh:mm:ss", cur_colm_str = "d", "mm/dd/yyyy")
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

