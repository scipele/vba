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


'Sheet 1 Code

Option Explicit
Const scale_fact As Double = 16777216 / 32767


Private Sub cmdSetBackgroundColor_Click()
    mColorConversions.SetColorsByRGB
End Sub


Private Sub cmdSetColorsByHex_Click()
    mColorConversions.SetColorsByHex
End Sub


Private Sub CommandButton1_Click()
    mColorConversions.SetColorsByLongVal
End Sub


Private Sub CommandButton2_Click()
    UserForm1.Show
End Sub


Private Sub ScrollBar1_Change()
    
    Dim cell As Range
    Set cell = Range("targetRng3")
    
    Application.ScreenUpdating = False
    
    ' Get scrollbar value (0 to 32767)
    Dim scrollValue As Long
    scrollValue = Sheet1.ScrollBar1.Value
    
    ' Scale to color index (0 to approximately 16777215) without overflow
    Dim color_lng As Long
    color_lng = Int(scrollValue * scale_fact)
    
    ' Set the cell interior color
    cell.Interior.Color = color_lng
    
    ' Set cell value that represents the long color value
    Dim long_clr_cell As Range
    Set long_clr_cell = Range("long_back_clr")
    long_clr_cell.Value = color_lng
    
    ' Re-enable screen updates
    Application.ScreenUpdating = True
End Sub


Private Sub ScrollBar2_Change()
    
    Dim cell As Range
    Set cell = Range("targetRng3")
    
    Application.ScreenUpdating = False
    
    ' Get scrollbar value (0 to 32767)
    Dim scrollValue As Long
    scrollValue = Sheet1.ScrollBar2.Value
    
    ' Scale to color index (0 to approximately 16777215) without overflow
    Dim color_lng As Long
    color_lng = Int(scrollValue * scale_fact)
    
    ' Set the cell interior color
    cell.Font.Color = color_lng
    
    ' Set cell value that represents the long color value
    Dim long_font_clr_rng As Range
    Set long_font_clr_rng = Range("long_font_clr")
    long_font_clr_rng.Value = color_lng
    
    ' Re-enable screen updates
    Application.ScreenUpdating = True
End Sub


Private Sub ScrollBarRed1_Change()
    Dim cell As Range
    Set cell = Range("redBackRng")
    cell.Value = ScrollBarRed1.Value
    mColorConversions.SetColorsByRGB

End Sub
Private Sub ScrollBarGrn1_Change()
    Dim cell As Range
    Set cell = Range("grnBackRng")
    cell.Value = ScrollBarGrn1.Value
    mColorConversions.SetColorsByRGB
End Sub

Private Sub ScrollBarBlu1_Change()
    Dim cell As Range
    Set cell = Range("bluBackRng")
    cell.Value = ScrollBarBlu1.Value
    mColorConversions.SetColorsByRGB
End Sub


Private Sub ScrollBarRed2_Change()
    Dim cell As Range
    Set cell = Range("redFontRng")
    cell.Value = ScrollBarRed2.Value
    mColorConversions.SetColorsByRGB
End Sub

Private Sub ScrollBarGrn2_Change()
    Dim cell As Range
    Set cell = Range("grnFontRng")
    cell.Value = ScrollBarGrn2.Value
    mColorConversions.SetColorsByRGB
End Sub

Private Sub ScrollBarBlu2_Change()
    Dim cell As Range
    Set cell = Range("bluFontRng")
    cell.Value = ScrollBarBlu2.Value
    mColorConversions.SetColorsByRGB
End Sub


'userform code
' ==== API Declarations ====
#If VBA7 Then
    Private Declare PtrSafe Function CreateDIBSection Lib "gdi32" (ByVal hdc As LongPtr, _
        pBitmapInfo As Any, ByVal un As Long, ByRef lplpVoid As LongPtr, _
        ByVal handle As LongPtr, ByVal dw As Long) As LongPtr
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As Any, RefIID As Any, _
        ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
'#Else
'    Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, _
'        pBitmapInfo As Any, ByVal un As Long, ByRef lplpVoid As Long, _
'        ByVal handle As Long, ByVal dw As Long) As Long
'    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'    Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As Any, RefIID As Any, _
'        ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
#End If

#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByVal Destination As LongPtr, _
        ByRef Source As Any, _
        ByVal Length As LongPtr)
'#Else
'    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
'        ByVal Destination As Long, _
'        ByRef Source As Any, _
'        ByVal Length As Long)
#End If


Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 0) As RGBQUAD
End Type

Private Type PicBmp
    Size As Long
    Type As Long
    hBmp As LongPtr
    hPal As LongPtr
    Reserved As LongPtr
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Const BI_RGB As Long = 0
Private Const DIB_RGB_COLORS As Long = 0
Private Const vbPicTypeBitmap = 1

' === Draw Gradient directly into Image1.Picture ===
Private Sub DrawRGBGradient()

    Const W As Long = 1024
    Const H As Long = 1024

    Dim bmi As BITMAPINFO
    Dim hBmp As LongPtr, pBits As LongPtr
    Dim x As Long, Y As Long
    Dim r As Byte, g As Byte, b As Byte
    
    ' Initialize BITMAPINFO
    With bmi.bmiHeader
        .biSize = Len(bmi.bmiHeader)
        .biWidth = W
        .biHeight = -H          ' negative = top-down DIB
        .biPlanes = 1
        .biBitCount = 32        ' 32-bit (RGBQUAD)
        .biCompression = BI_RGB
    End With
    
    ' Create DIBSection (returns pointer to pixel buffer)
    hBmp = CreateDIBSection(0, bmi, DIB_RGB_COLORS, pBits, 0, 0)
    If hBmp = 0 Then Exit Sub
    
    ' Fill pixels directly in memory
    Dim pixels() As Long
    ReDim pixels(0 To W * H - 1)

    For bVal = 0 To 255 Step 8
        For Y = 0 To H - 1
            For x = 0 To W - 1
                r = CByte((x * 255) \ (W - 1))
                g = CByte((Y * 255) \ (H - 1))
                b = bVal
                pixels(Y * W + x) = b Or (g * &H100&) Or (r * &H10000)
            Next x
        Next Y
        ' CopyMemory into DIB and refresh Image1.Picture
        ' Sleep or DoEvents to see animation
    Next bVal


    ' Copy pixel data into DIB memory
    CopyMemory ByVal pBits, pixels(0), W * H * 4
    
    ' Wrap the HBITMAP into StdPicture
    Dim IID_IDispatch As GUID
    With IID_IDispatch
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    
    Dim pic As PicBmp
    Dim IPic As IPicture
    
    pic.Size = Len(pic)
    pic.Type = vbPicTypeBitmap
    pic.hBmp = hBmp
    
    OleCreatePictureIndirect pic, IID_IDispatch, True, IPic
    
    ' Assign to Image1
    Set Image1.Picture = IPic
End Sub


Private Sub CommandButton1_Click()
    Call DrawRGBGradient
End Sub
