Public Sub fracFeetInToDecIn()
    Dim text As String
    Dim DecInches As Double
    For Each cell In Selection
        text = cell.Value
        cell.Value = convFtInToDecIn(text)
    Next cell
End Sub


Public Sub fracftInToDecFt()
    Dim text As String
    Dim DecFt As Double
    For Each cell In Selection
        text = cell.Value
        cell.Value = convFtInToDecFt(text)
    Next cell
End Sub


Public Function convFtInToDecIn(measStr As String)
    convFtInToDecIn = convFtInToDecFt(measStr) * 12
End Function


Public Function convFtInToDecFt(measStr As String)
    Dim mInch, mFeet As Double
    If measStr = "" Then
        convFtInToDecFt = ""
        Exit Function
    End If
    mFeet = GetFeetPart(measStr)
    mInch = GetInchPart(measStr)
    convFtInToDecFt = mFeet + mInch / 12
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