Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | setCellColor.vba                                            |
'| EntryPoint   | Sub call from spreadsheet button                            |
'| Purpose      | sets interior cell colors and fonts                         |
'| Inputs       | color values read from named ranges                         |
'| Outputs      | sets the background and font color                          |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 8/25/2025                                         |


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

