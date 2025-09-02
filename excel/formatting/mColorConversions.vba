Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | setCellColor.vba                                            |
'| EntryPoint   | various                                                     |
'| Purpose      | sets interior cell colors and fonts                         |
'| Inputs       | color values read from named ranges                         |
'| Outputs      | sets the background and font color                          |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 8/25/2025                                         |


Sub SetColorsAlt()
    
    Dim target_rng As Range
    Set target_rng = Range("a1:p255")

    Dim r As Integer
    Dim b As Integer
    Dim g As Integer
    
    Dim row As Long
    Dim col As Long
    
    row = 1
    col = 1
    
    Dim cell As Range
    Dim font_clr_lng As Long
    For r = 0 To 255 Step 17
        For b = 0 To 255 Step 17
            For g = 0 To 255 Step 17
                Set cell = target_rng(row, col)
                cell.Interior.Color = RGB(r, g, b)
                
                font_clr_lng = GetContrastingColor(r, g, b)
                cell.Font.Color = font_clr_lng
                cell.Value = rgbToHex(r, g, b)
                
                If col = 16 Then
                    col = 1
                    row = row + 1
                Else: col = col + 1
                End If
            Next g
        Next b
    Next r
End Sub
    
    
Private Function rgbToHex(ByVal r As Byte, _
                            ByVal g As Byte, _
                            ByVal b As Byte) _
                            As String

    ' Format as #RRGGBB
    rgbToHex = "#" & _
                Right$("0" & Hex(r), 2) & _
                Right$("0" & Hex(g), 2) & _
                Right$("0" & Hex(b), 2)
End Function

' Returns either vbBlack or vbWhite depending on best contrast with the background
Public Function GetContrastingColor(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte) As Long
    Dim rNorm As Double, gNorm As Double, bNorm As Double
    Dim rLin As Double, gLin As Double, bLin As Double
    Dim Y As Double
    
    ' Normalize to [0,1]
    rNorm = r / 255#
    gNorm = g / 255#
    bNorm = b / 255#
    
    ' Gamma correction (sRGB)
    rLin = IIf(rNorm <= 0.04045, rNorm / 12.92, ((rNorm + 0.055) / 1.055) ^ 2.4)
    gLin = IIf(gNorm <= 0.04045, gNorm / 12.92, ((gNorm + 0.055) / 1.055) ^ 2.4)
    bLin = IIf(bNorm <= 0.04045, bNorm / 12.92, ((bNorm + 0.055) / 1.055) ^ 2.4)
    
    ' Relative luminance (W3C formula)
    Y = 0.2126 * rLin + 0.7152 * gLin + 0.0722 * bLin
    
    ' Decide best contrast: if background is light, return black; else white
    If Y > 0.5 Then
        GetContrastingColor = vbBlack
    Else
        GetContrastingColor = vbWhite
    End If
End Function


Sub SetColors()
    
    Dim r() As Byte
    Dim g() As Byte
    Dim b() As Byte
    r = GetAryFromRange("rRng", "c")
    g = GetAryFromRange("gRng", "c")
    b = GetAryFromRange("bRng", "r")

    Dim target_rng As Range
    Set target_rng = Range("targetRng")

    Dim i As Long
    Dim j As Long
    Dim cell As Range
    
    For i = LBound(r) To UBound(r)
        For j = LBound(b) To UBound(b)
            Set cell = target_rng(i, j)
            cell.Interior.Color = RGB(r(i), g(i), b(j))
        Next j
    Next i
End Sub


Private Function GetAryFromRange(ByVal rngName As String, _
                                 ByVal r_c As String) _
                                 As Byte()
    Dim tmp_ary As Variant
    tmp_ary = Range(rngName)
    
    Dim byte_ary() As Byte
    
    Dim elem As Variant
    Dim i As Long
    i = LBound(tmp_ary)
    
    If r_c = "c" Then
        ReDim byte_ary(LBound(tmp_ary) To UBound(tmp_ary))
        
        For i = LBound(tmp_ary) To UBound(tmp_ary)
            byte_ary(i) = tmp_ary(i, 1)
        Next i
    Else
        ReDim byte_ary(LBound(tmp_ary, 2) To UBound(tmp_ary, 2))
        For i = LBound(tmp_ary, 2) To UBound(tmp_ary, 2)
            byte_ary(i) = tmp_ary(1, i)
        Next i
    End If

    GetAryFromRange = byte_ary
End Function


Public Sub SetColorsByRGB()
    Dim cell As Range
    Set cell = Range("targetRng1")
    
    Dim color_ary As Variant
    color_ary = Range("rgbColorRng")
    
    Dim r_back As Byte, g_back As Byte, b_back As Byte
    r_back = color_ary(1, 1)
    g_back = color_ary(2, 1)
    b_back = color_ary(3, 1)
    
    Dim r_fore As Byte, g_fore As Byte, b_fore As Byte
    r_fore = color_ary(1, 2)
    g_fore = color_ary(2, 2)
    b_fore = color_ary(3, 2)
    
    Dim clr_back As Long
    clr_back = RGB(r_back, g_back, b_back)
    cell.Interior.Color = clr_back
    
    Dim clr_fore As Long
    clr_fore = RGB(r_fore, g_fore, b_fore)
    cell.Font.Color = clr_fore

End Sub


Public Sub SetColorsByHex()
    Dim cell As Range
    Set cell = Range("targetRng2")
    
    Dim color_ary As Variant
    color_ary = Range("hexColorRng")
    
    Dim clr_back As Long
    clr_back = GetHexColorToLong(color_ary(1, 1))
    cell.Interior.Color = clr_back
    
    Dim clr_fore As Long
    clr_fore = GetHexColorToLong(color_ary(1, 2))
    cell.Font.Color = clr_fore

End Sub


Function GetHexColorToLong(s As Variant) As String
    Dim i As Integer
    Dim validHex As String
    Dim cleanString As String
    validHex = "0123456789ABCDEFabcdef"
    
    ' Remove # prefix if present
    If Left(s, 1) = "#" Then
        cleanString = Mid(s, 2)
    Else
        cleanString = s
    End If
    
    ' Check if string is exactly 6 characters
    If Len(cleanString) <> 6 Then
        GetHexColorToLong = ""
        Exit Function
    End If
    
    ' Check each character
    For i = 1 To Len(cleanString)
        If InStr(1, validHex, Mid(cleanString, i, 1), vbTextCompare) = 0 Then
            GetHexColorToLong = ""
            Exit Function
        End If
    Next i
    
    ' Convert RGB to BGR by rearranging the hex string
    Dim bgr_hex As String
    bgr_hex = Mid(cleanString, 5, 2) & Mid(cleanString, 3, 2) & Mid(cleanString, 1, 2)
    
    ' Convert BGR hex to Long
    GetHexColorToLong = CLng("&H" & bgr_hex)
End Function


Public Sub SetInteriorColorByLongVal(ByVal scrollValue As Long)
    Dim cell As Range
    Set cell = Range("targetRng3")
    
    Dim colorIndex As Long
    colorIndex = (scrollValue * 16777215) \ 32767 ' Integer division for efficiency
    
    cell.Interior.Color = colorIndex
End Sub