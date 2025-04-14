Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | highlightArrowsColors.vba                                   |
'| EntryPoint   | Function call from spreadsheet                              |
'| Purpose      | creates and up/down arrow and color it Red, Green, Black    |
'| Inputs       | two ranges from excel sheet                                 |
'| Outputs      | symbol, font color of cells                                 |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 4/14/2025                                         |


Public Function get_arrow_and_num(ByRef mo_rng As Range, _
                                  ByVal base_qty As Double) _
                                  As Variant
   
    ' Loop thru each cell in the 'mo_rng' month quantity range until blank
    Dim mo As Variant
    Dim cur_mo_qty As Double
    Dim month_data As Long
    For Each mo In mo_rng
        If mo <> "" Then
            cur_mo_qty = mo.Value
            month_data = month_data + 1
        End If
    Next mo
    
    ' Compute the fractional change of the latest quantity verses the base qty
    Dim frac_qty_chg As Double
    
    If base_qty = 0 Then
        frac_qty_chg = 1    'Just set to 100% if there was no previous qty
    Else
        frac_qty_chg = (cur_mo_qty - base_qty) / base_qty
    End If
    
    ' Call the private function below to get the string range for
    ' the cells that will be colored either black or colored
    Dim rng_str As Variant
    rng_str = GetColorRangeStrings(mo_rng, month_data)
    
    ' Set const long integer values for the colors
    ' note that RGB decimal equivalents are hard coded
    ' Long Color = Red + (Green×256) + (Blue×65536))
    Const RED_COLOR As Long = 255       ' Long equivalent of RGB(255, 0, 0)
    Const GREEN_COLOR As Long = 5287936 ' Long equivalent of  RGB(0, 176, 80)
    Const BLACK_COLOR As Long = 0       ' Long equivalent of  RGB(0, 0, 0)
    
    'set other older months to black color
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Client Summary") ' Change to your sheet name
    ws.Range(rng_str(0)).Font.Color = BLACK_COLOR
        
    'set uparrow, downarrow, or none
    Dim symb As String
    
    Select Case True
        Case frac_qty_chg > 0
            symb = ChrW(&H2191)
            ws.Range(rng_str(1)).Font.Color = RED_COLOR
            
        Case frac_qty_chg < 0
            symb = ChrW(&H2193)
            ws.Range(rng_str(1)).Font.Color = GREEN_COLOR
        
        Case frac_qty_chg = 0
            symb = ""
            ws.Range(rng_str(1)).Font.Color = BLACK_COLOR
    End Select
    
    get_arrow_and_num = symb & Format(frac_qty_chg * 100, "0") & "%"
    
End Function


Private Function GetColorRangeStrings(ByVal mo_rng As Range, _
                                      ByVal month_data As Long) _
                                      As Variant

    ' Determine the row number bsaed on the range passed which is used to
    ' change the font color
    Dim cur_row As Long
    cur_row = mo_rng.Rows(1).Row
    
    'get range to highlight in colors, and previous months to make black
    Dim colm_no(0 To 2) As Long
    colm_no(0) = mo_rng.Column  'start of the months range
    colm_no(1) = mo_rng.Column + month_data - 1
    If colm_no(1) < colm_no(0) Then
        colm_no(1) = colm_no(0)
    End If
    colm_no(2) = mo_rng.Column + mo_rng.Count
    
    ' convert the column numbers above to letters using the cellls.address
    Dim colm_ltr(0 To 2) As String
    colm_ltr(0) = Split(Cells(1, colm_no(0)).Address, "$")(1)
    colm_ltr(1) = Split(Cells(1, colm_no(1)).Address, "$")(1)
    colm_ltr(2) = Split(Cells(1, colm_no(2)).Address, "$")(1)
    
    Dim rng_str As Variant
    ReDim rng_str(0 To 1)
    rng_str(0) = colm_ltr(0) & cur_row & ":" & colm_ltr(1) & cur_row   'Black
    rng_str(1) = colm_ltr(1) & cur_row & ":" & colm_ltr(2) & cur_row   'Color
    
    GetColorRangeStrings = rng_str
End Function