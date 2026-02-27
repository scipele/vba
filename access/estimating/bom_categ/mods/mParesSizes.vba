Option Compare Database

Sub test()
    Dim strg As String
    strg = "8"" S/STD BORE A106B SMLS CS PIPE"
    Dim col_marker As Integer
    Dim s As String
    
    s = getSize1(strg, col_marker)
    
    RemoveNonPrintableASCII (strg)
    Debug.Print strg
End Sub




' Subroutine: ParseAndUpdateBomSizes
' Reads all records from 'd_bom_raw', parses size 1 and 2 from desc_raw,
' and updates sz_1 and sz_2 fields in the same table.
Public Sub ParseAndUpdateBomSizes()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sz1 As String, sz2 As String
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM d_bom_raw", dbOpenDynaset)
    Dim col_marker As Integer
    Dim desc_str As String

    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do While Not rs.EOF
            col_marker = 0
            desc_str = Nz(rs!desc_w_size, "")
            desc_str = RemoveNonPrintableASCII(desc_str)
            sz1 = getSize1(desc_str, col_marker)
            sz2 = getSize2(desc_str, col_marker)
            rs.Edit
            rs!sz_1 = sz1
            rs!sz_2 = sz2
            If col_marker > 0 Then
                rs!desc = Right(desc_str, Len(desc_str) - col_marker - 1)
            Else
                rs!desc = desc_str
            End If
            rs.Update
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    MsgBox "Size parsing and update complete.", vbInformation
End Sub


Public Function getSize1(ByVal strg As String, _
                         ByRef col_marker As Integer) _
                         As String
    'shorten the string
    If Len(strg) > 7 Then strg = Left(strg, 7)
    'handle flat size

    If InStr(1, strg, "flat", vbTextCompare) > 0 Then
        col_marker = InStr(1, strg, "flat", vbTextCompare)
        getSize1 = "FLAT"
        Exit Function
    End If
    'check for inch mark
    Dim inchLoc1 As Integer
    inchLoc1 = InStr(1, strg, """", vbTextCompare)
    
    If inchLoc1 > 0 Then
        col_marker = inchLoc1
        getSize1 = convFtInToDecIn(Trim(Left(strg, inchLoc1)))
    Else
        getSize1 = ""
    End If
End Function


Public Function getSize2(ByVal strg As String, _
                         ByRef col_marker As Integer) _
                         As String
                         
    Dim inchLoc1, inchLoc2, locX, LenLoc As Integer
    Dim tmpSize2 As String
    
    inchLoc1 = InStr(1, strg, """", vbTextCompare)
    inchLoc2 = InStr(inchLoc1 + 1, strg, """", vbTextCompare)
    
    'Make Sure that Size 2 is not actually a length
        LenLoc = InStr(inchLoc1 + 1, LCase(strg), """ long", vbTextCompare)
    
        If LenLoc = 0 Then
            LenLoc = InStr(inchLoc1 + 1, LCase(strg), """ lg", vbTextCompare)
        End If
    
    If inchLoc2 = LenLoc Or inchLoc2 > 13 Then
        getSize2 = ""
        Exit Function
    End If
  
    If Not inchLoc2 = 0 Then
        'make sure that the character to left of inch mark is numeric
        Dim prev_char As String
        prev_char = Mid(strg, inchLoc2 - 1, 1)
        If Not IsNumeric(prev_char) Then
            getSize2 = ""
            Exit Function
        End If
        
        tmpSize2 = Mid(strg, inchLoc1, inchLoc2 - inchLoc1 + 1)
        locX = InStr(1, LCase(tmpSize2), "x", vbTextCompare)
        tmpSize2 = Right(tmpSize2, Len(tmpSize2) - locX)
        
        col_marker = inchLoc2
        getSize2 = convFtInToDecIn(tmpSize2)
    End If
End Function


Public Function convFtInToDecIn(measStr As String)
    Dim mInch, mFeet As Double
    If measStr = "" Then
        convFtInToDecFt = ""
        Exit Function
    End If
    mFeet = GetFeetPart(measStr)
    mInch = GetInchPart(measStr)
    convFtInToDecIn = mFeet * 12 + mInch
End Function


Private Function GetFeetPart(measStr As String) As Double
    Dim i, delimLoc As Integer
    Dim curDelimeter() As String
    curDelimeter = Split("'-,',feet,ft.,ft,f", ",")
    delimLoc = GetStartDelimLocation(measStr, curDelimeter)
    If delimLoc > 0 Then
        GetFeetPart = Trim(Left(measStr, delimLoc - 1))
    Else
        GetFeetPart = 0
    End If
End Function


Private Function GetInchPart(measStr As String) As Double
    Dim i, delimEndLocFt, delimStartLocIn, strLnInch As Integer
    Dim curDelimeter() As String
    Dim InchPartStr As String
    
    curDelimeter = Split("'-,',feet,ft.,ft,f", ",")
    delimEndLocFt = GetEndOfDelimLocation(measStr, curDelimeter)
    
    curDelimeter = Split(""",inches,in.,in,i", ",")
    delimStartLocIn = GetStartDelimLocation(measStr, curDelimeter)
    
    strLnInch = delimStartLocIn - delimEndLocFt
    InchPartStr = Trim(Mid(measStr, delimEndLocFt, strLnInch))
    GetInchPart = InchStrToDec(InchPartStr)
End Function

Private Function GetStartDelimLocation(measStr As String, curDelimeter)
    Dim i As Integer
    For i = 0 To UBound(curDelimeter)
        If InStr(measStr, curDelimeter(i)) > 0 Then
            GetStartDelimLocation = InStr(1, measStr, curDelimeter(i), vbTextCompare)
            Exit For
        End If
    Next i
End Function


Private Function InchStrToDec(InchPartStr As String)
    Dim InchPartAry() As String
    Dim InchFormat As String
    
    InchPartStr = Replace(InchPartStr, " ", "-", 1, -1, vbTextCompare)
    InchPartStr = Replace(InchPartStr, "-", ",", 1, -1, vbTextCompare)
    InchPartStr = Replace(InchPartStr, "/", ",", 1, -1, vbTextCompare)
    
    InchPartAry = Split(InchPartStr, ",")

    Select Case UBound(InchPartAry)
        Case 0 'Whole number only when array is only one element
            InchStrToDec = Val(InchPartStr)
        Case 1 'Fraction only when array contains only two elements
            InchStrToDec = Val(InchPartAry(0) / InchPartAry(1))
        Case 2 'Whole number and Fraction when array contains three elements (i.e. (0),(1),(2)
            InchStrToDec = Val(InchPartAry(0) + InchPartAry(1) / InchPartAry(2))
    End Select
End Function


Private Function GetEndOfDelimLocation(measStr As String, curDelimeter)
    Dim i As Integer
    For i = 0 To UBound(curDelimeter)
        If InStr(measStr, curDelimeter(i)) > 0 Then
            GetEndOfDelimLocation = InStr(measStr, curDelimeter(i)) + Len(curDelimeter(i))
            Exit For
        End If
    Next i
    If GetEndOfDelimLocation = "" Then GetEndOfDelimLocation = 1
End Function


Private Function RemoveNonPrintableASCII(ByVal desc_str As String)
    Dim i As Integer
    Dim result As String
    result = ""
    For i = 1 To Len(desc_str)
        Dim charCode As Integer
        charCode = AscW(Mid(desc_str, i, 1))
        If charCode >= 32 And charCode <= 126 Then
            result = result & Mid(desc_str, i, 1)
        End If
    Next i
    RemoveNonPrintableASCII = result
End Function
