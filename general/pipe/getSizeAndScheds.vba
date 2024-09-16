Option Explicit
' filename:         getSizeAndScheds.vba
'
' purpose:          split portions of strings for size1, size2, 
'                   sched1, sched2
'
' usage:            sz1 = getSize1(line_string)
'
' dependencies:     none
'
' By:               T.Sciple, 09/16/2024

Public Function getSize1(ByVal strg As String) _
                         As String
    
    Dim inchLoc1 As Integer
    inchLoc1 = InStr(1, strg, """", vbTextCompare)
    getSize1 = convFtInToDecIn(Left(strg, inchLoc1))
End Function


Public Function getSize2(ByVal strg As String) _
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
    
    If inchLoc2 = LenLoc Then
        inchLoc2 = 0
    End If
        
    If inchLoc2 = 0 Then
        getSize2 = ""
    Else
        tmpSize2 = Mid(strg, inchLoc1, inchLoc2 - inchLoc1 + 1)
        locX = InStr(1, LCase(tmpSize2), "x", vbTextCompare)
        tmpSize2 = Right(tmpSize2, Len(tmpSize2) - locX)
        getSize2 = convFtInToDecIn(tmpSize2)
    End If
End Function


Public Function get_sch_1(ByVal strg As String) _
                          As String
    Dim locX As Integer
    locX = InStr(1, strg, "x", vbBinaryCompare)
    If locX > 0 Then
        get_sch_1 = Left(strg, locX - 2)
    Else
        get_sch_1 = strg
    End If
End Function


Public Function get_sch_2(ByVal strg As String) _
                          As String
    Dim locX As Integer
    locX = InStr(1, strg, "x", vbBinaryCompare)
    
    If locX > 0 Then
        get_sch_2 = Right(strg, Len(strg) - locX - 1)
    Else
        get_sch_2 = ""
    End If
End Function