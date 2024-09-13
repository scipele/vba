Public Function getTextLftOf(ByVal text As String, _
                      ByVal matchTxt As String, _
                      ByVal startLoc As Integer, _
                      ByVal noDelimeters As Integer) _
                      As String
    Dim i
    Dim Location As Long
    For i = 1 To noDelimeters
        Location = InStr(startLoc, text, matchTxt, vbTextCompare)
        startLoc = Location + 1
    Next i

    getTextLftOf = Trim(Left(text, Location - 1))
End Function


Public Function getTextRghtOf(ByVal text As String, _
                       ByVal matchTxt As String) _
                       As String
    Dim Location As Long
    Location = InStr(1, text, matchTxt, vbTextCompare)
    
    Dim strLength As Long
    strLength = Len(text)
    getTextRghtOf = Trim(Right(text, strLength - Location))
End Function


Public Function getTextBetween(ByVal text As String, _
                        ByVal param1 As String, _
                        ByVal param2 As String, _
                        ByVal Seq1 As String, _
                        ByVal Seq2 As String) _
                        As String
    Dim i As Integer
    Dim Location1 As Long
    Dim startLoc As Long
    startLoc = 1
    For i = 1 To Seq1
        Location1 = InStr(startLoc, text, param1, vbTextCompare)
        startLoc = Location1 + 1
    Next i
    startLoc = 1
    
    Dim Location2 As Long
    For i = 1 To Seq2
        Location2 = InStr(startLoc, text, param2, vbTextCompare)
        startLoc = Location2 + 1
    Next i
    
    Dim strLen As Long
    strLen = Location2 - Location1 - 1
    getTextBetween = Mid(text, Location1 + 1, strLen)
End Function


Public Function lookForSimilarText(ByVal lookupText As String, _
                            ByVal lookupRng As Range) _
                            As Variant

    Dim i As Long
    For i = 0 To UBound(lookupRng())
        Location = InStr(1, UCase(lookupRng(i)), UCase(lookupText), vbTextCompare)
        If Location > 0 Then Exit For
    Next i
    
    If Location > 0 Then
        lookForSimilarText = lookupRng(i).Value2
    Else
        lookForSimilarText = "Not Found"
    End If

End Function


Public Function lookForSimilarTextRow(ByVal lookupText As String, _
                               ByVal lookupRng As Range) _
                               As Variant

    'returns the row number were text is found in a given range
    Dim i As Long
    For i = 0 To UBound(lookupRng())
        Location = InStr(1, UCase(lookupRng(i)), UCase(lookupText), vbTextCompare)
        If Location > 0 Then Exit For
    Next i
    
    If Location > 0 Then
        lookForSimilarTextRow = i
    Else
        lookForSimilarTextRow = "Not Found"
    End If
End Function


Public Function get_vlv_oper_type(str As String) As String
    Dim tmp As Variant  'used variant to handle the null if not found
    tmp = Switch(InStr(1, str, ", HW", vbTextCompare) > 0, "HW", _
        InStr(1, str, "Gear", vbTextCompare) > 0, "Gear")
    If IsNull(tmp) Then tmp = ""
    get_vlv_oper_type = tmp
End Function


Public Function get_vlv_wedge_type(str As String) As String
    Dim tmp As Variant  'used variant to handle the null if not found
    tmp = Switch(InStr(1, str, "Flex Wdg", vbTextCompare) > 0, "Flex Wdg", _
        InStr(1, str, "Flexible Wedge", vbTextCompare) > 0, "Flex Wdg", _
        InStr(1, str, "Sol Wdg", vbTextCompare) > 0, "Sol Wdg", _
        InStr(1, str, "Solid Wdg", vbTextCompare) > 0, "Solid Wdg")
    If IsNull(tmp) Then tmp = ""
    get_vlv_wedge_type = tmp
End Function


Public Function get_vlv_type(str As String) As String
    Dim tmp As Variant  'used variant to handle the null if not found
    tmp = Switch(InStr(1, str, "GAT", vbTextCompare) > 0, "GATE", _
        InStr(1, str, "BALL", vbTextCompare) > 0, "BALL", _
        InStr(1, str, "CHECK", vbTextCompare) > 0, "CHECK", _
        InStr(1, str, "CHK", vbTextCompare) > 0, "CHECK", _
        InStr(1, str, "NEEDLE", vbTextCompare) > 0, "NEEDLE", _
        InStr(1, str, "GLOBE", vbTextCompare) > 0, "GLOBE")
    If IsNull(tmp) Then tmp = "not found"
    get_vlv_type = tmp
End Function


Public Function get_vlv_body_matl(str As String) As String
    Dim tmp As Variant  'used variant to handle the null if not found
    tmp = Switch(InStr(1, str, "A216 WCB", vbTextCompare) > 0, "A216-WCB", _
        InStr(1, str, "A352 LCC", vbTextCompare) > 0, "A352-LCC", _
        InStr(1, str, "A105", vbTextCompare) > 0, "A105", _
        InStr(1, str, "A216 WCB", vbTextCompare) > 0, "A216-WCB", _
        InStr(1, str, "A350 LF2 CI 1", vbTextCompare) > 0, "A350-LF2-CL1", _
        InStr(1, str, "A350 LF2 CI 1", vbTextCompare) > 0, "A350-LF2-CL1", _
        InStr(1, str, "A350", vbTextCompare) > 0, "A350-LF2-CL1")  ' address typo missing info
    If IsNull(tmp) Then tmp = "not found"
    get_vlv_body_matl = tmp
End Function


Public Function get_vlv_api_trim(str As String) As String
    Dim tmp As Variant  'used variant to handle the null if not found
    tmp = Switch(InStr(1, str, "#6", vbTextCompare) > 0, "#6", _
        InStr(1, str, "#8", vbTextCompare) > 0, "#8", _
        InStr(1, str, "#10", vbTextCompare) > 0, "#10", _
        InStr(1, str, "#10", vbTextCompare) > 0, "#10", _
        InStr(1, str, "#12", vbTextCompare) > 0, "#12", _
        InStr(1, str, "SS316 TRIM", vbTextCompare) > 0, "#10")
    If IsNull(tmp) Then tmp = "not found"
    get_vlv_api_trim = tmp
End Function


Public Function get_rtg_lb_or_sb(str As String) As String
    Dim tmp As Variant  'used variant to handle the null if not found
    tmp = Switch(InStr(1, str, "1500", vbTextCompare) > 0, "1500", _
                 InStr(1, str, "300", vbTextCompare) > 0, "300", _
                 InStr(1, str, "600", vbTextCompare) > 0, "600", _
                 InStr(1, str, "800", vbTextCompare) > 0, "800", _
                 InStr(1, str, "900", vbTextCompare) > 0, "900", _
                 InStr(1, str, "150", vbTextCompare) > 0, "150", _
                 InStr(1, str, "2500", vbTextCompare) > 0, "2500")
    If IsNull(tmp) Then tmp = "not found"
    get_rtg_lb_or_sb = tmp
End Function


Public Function get_rtg_lb(str As String) As String
    Dim tmp As Variant  'used variant to handle the null if not found
    tmp = Switch(InStr(1, str, "1500", vbTextCompare) > 0, "1500", _
                 InStr(1, str, "300", vbTextCompare) > 0, "300", _
                 InStr(1, str, "600", vbTextCompare) > 0, "600", _
                 InStr(1, str, "900", vbTextCompare) > 0, "900", _
                 InStr(1, str, "150", vbTextCompare) > 0, "150", _
                 InStr(1, str, "2500", vbTextCompare) > 0, "2500")
    If IsNull(tmp) Then tmp = "not found"
    get_rtg_lb = tmp
End Function


Function get_rtg_sb(str As String) As String
    Dim tmp As Variant  'used variant to handle the null if not found
    tmp = Switch(InStr(1, str, "1500", vbTextCompare) > 0, "1500", _
        InStr(1, str, "800", vbTextCompare) > 0, "800")
    If IsNull(tmp) Then tmp = "not found"
    get_rtg_sb = tmp
End Function


Function get_weld_sch1(str As String) As String
    Dim tmp As Variant  'used variant to handle the null if not found
    tmp = Switch(InStr(1, str, "3000", vbTextCompare) > 0, "XS", _
        InStr(1, str, "S-0.625", vbTextCompare) > 0, "0.625", _
        InStr(1, str, "6000", vbTextCompare) > 0, "160", _
        InStr(1, str, "S-40S", vbTextCompare) > 0, "40S", _
        InStr(1, str, "S-80S", vbTextCompare) > 0, "80S", _
        InStr(1, str, "S-10S", vbTextCompare) > 0, "10S", _
        InStr(1, str, "S-STD", vbTextCompare) > 0, "STD", _
        InStr(1, str, "S-XS", vbTextCompare) > 0, "XS", _
        InStr(1, str, "S-XXS", vbTextCompare) > 0, "XXS", _
        InStr(1, str, "S-20", vbTextCompare) > 0, "20", _
        InStr(1, str, "S-30", vbTextCompare) > 0, "30", _
        InStr(1, str, "S-60", vbTextCompare) > 0, "60", _
        InStr(1, str, "S-80", vbTextCompare) > 0, "80", _
        InStr(1, str, "S-40", vbTextCompare) > 0, "40", _
        InStr(1, str, "S-100", vbTextCompare) > 0, "100", _
        InStr(1, str, "S-120", vbTextCompare) > 0, "120", _
        InStr(1, str, "S-160", vbTextCompare) > 0, "160" _
        )
    If IsNull(tmp) Then tmp = "not found"
    get_weld_sch1 = tmp
End Function


Sub SplitCellsByDelimiter()
    Dim cell As Range
    Dim delimiter As Variant
    Dim splitValues As Variant
    Dim i As Integer
    
    ' Prompt the user for the delimiter
    delimiter = InputBox("Enter the delimiter:", "Delimiter")
    
    If delimiter = "vbLf" Then delimiter = vbLf
    If delimiter = "vbCr" Then delimiter = vbCr
    If delimiter = "vbCrLf" Then delimiter = vbCrLf
    
    ' Check if the delimiter is not empty
    If delimiter = "" Then
        MsgBox "Please enter a valid delimiter."
        Exit Sub
    End If
    
    ' Iterate through each selected cell
    For Each cell In Selection
        ' Split the cell value by the delimiter
        splitValues = Split(cell.Value, delimiter)
        
        ' Clear the current cell
        cell.Value = ""
        
        ' Place the split values in the cell and adjacent cells to the right
        For i = LBound(splitValues) To UBound(splitValues)
            cell.Offset(0, i).Value = splitValues(i)
        Next i
    Next cell
End Sub


Public Sub merge_contents_w_cr()
    Dim i As Long
    For Each cell In Selection
        i = i + 1
        ActiveCell.Value = Switch(i = 1, cell.Value & vbCrLf, _
                        i < Selection.Count, ActiveCell.Value & cell.Value & vbCrLf, _
                        i = Selection.Count, ActiveCell.Value & cell.Value)
        If i > 1 Then cell.Value = ""
    Next cell
End Sub


Public Sub merge_contents_w0_cr()
    Dim i As Long
    For Each cell In Selection
        i = i + 1
        ActiveCell.Value = Switch(i = 1, cell.Value, _
                        i <= Selection.Count, ActiveCell.Value & cell.Value)
        If i > 1 Then cell.Value = ""
    Next cell
End Sub


Public Sub merge_contents_w_comma()
    Dim i As Long
    For Each cell In Selection
        i = i + 1
        ActiveCell.Value = Switch(i = 1, cell.Value & ",", _
                        i < Selection.Count, ActiveCell.Value & cell.Value & ",", _
                        i = Selection.Count, ActiveCell.Value & cell.Value)
        If i > 1 Then cell.Value = ""
    Next cell
End Sub


Public Sub UpperC()
    For Each cell In Selection
    cell.Value = UCase(cell.Value)
    Next cell
End Sub


Public Sub lowerC()
    For Each cell In Selection
    cell.Value = LCase(cell.Value)
    Next cell
End Sub


Public Sub properC()
    For Each cell In Selection
    cell.Value = WorksheetFunction.Proper(cell.Value)
    Next cell
End Sub


Public Sub addApst()
    For Each cell In Selection
    cell.Value = "'" & cell.Value
    Next cell
End Sub


Public Function RemoveMultSpaces(inp As String)
    Do
        inp = Replace(inp, "  ", " ")
    Loop Until InStr(inp, "  ") = 0
    RemoveMultSpaces = inp
End Function


Public Sub PasteDn()
'Selected Range Cells are copied down to blank cells below with identical values
    Dim lastcellVal, currentText As Variant
    Dim cell As Variant
    
    Call macrSettings.Speedup
    
    lastcellVal = ""
    For Each cell In Selection
        currentText = cell.Value
        If currentText = "" Then
            cell.Value = lastcellVal
        End If
        lastcellVal = cell.Value
    Next cell
    
    Call macrSettings.Restore

End Sub

Public Sub CreateHyperLink()
'Selected Range Cells are copied down to blank cells below with identical values
    Dim LinkAry As Variant
    Dim i, j, curRow As Long
    Dim curtxt As String
    
    'set the starting location for your hyperlinks
    startCol = Selection.Column + 2
    startRow = Selection.row
    startColLtr = Split(Cells(startRow, startCol).Address, "$")(1)
    
    LinkAry = Selection
    
    curRow = startRow
    
    For i = LBound(LinkAry, 1) To UBound(LinkAry, 1)
        If IsNumeric(LinkAry(i, 1)) = True Then
            curtxt = CStr(LinkAry(i, 1))
        Else:
            curtxt = LinkAry(i, 1)
        End If
        
        ActiveSheet.Hyperlinks.Add Range(startColLtr & curRow), Address:=LinkAry(i, 2), TextToDisplay:=curtxt
        curRow = curRow + 1
    Next i

End Sub


Public Function OddNo(x As Long)
'PURPOSE: Test whether an number is odd or even

    If x Mod 2 = 0 Then
        OddNo = False
    Else
        OddNo = True
    End If

End Function


Public Function PadTrailSpaces(inp As String, numChars As Integer)
    'This function pads a string input with trailing zeros to get to the number of characters specified
    'Example
    '   inputNo     = 'ball'
    '   numChars    = 7
    '   Results     = 'ball   '

    If Len(inp) >= numChars Then
        PadTrailSpaces = "error input longer than specified no of chars"
        Exit Function
    End If
    
    Dim build_return As String

    'start with it equal to the input
    build_return = inp

    While Len(build_return) < numChars
        build_return = build_return & " "
    Wend
    
    PadTrailSpaces = build_return

End Function


Public Function PadLeadZeros(inputNo As Variant, numChars As Integer)
    'This function pads a numeric input with zeros to get to the number of characters specified
    'Example
    '   inputNo     = 25
    '   numChars    =  5
    '   Results     = 00025
    Dim inputNoStr As String
    If IsNumeric(inputNo) = True Then
        inputNoStr = Trim(str(inputNo))
    Else
        inputNoStr = Trim(inputNo)
    End If
    
    While Len(inputNoStr) < numChars
        inputNoStr = "0" & inputNoStr
    Wend
    
    If Len(inputNoStr) > numChars Then
        inputNoStr = inputNoStr & " Warning - Number input contains more characters than specified"
    End If
    PadLeadZeros = inputNoStr
End Function


Public Function findInReverse(stringToSearch, SoughtItem)
    Dim foundLoc As Integer
    foundLoc = InStrRev(stringToSearch, SoughtItem, -1, vbTextCompare)
    findInReverse = foundLoc
End Function


Public Function asciien(ByVal s As String) As String
' Returns the string to its respective ascii numbers
   Dim i As Integer
   For i = 1 To Len(s)
      asciien = asciien & CStr(Asc(Mid(s, i, 1))) & Chr(13)
   Next i
   
   Debug.Print asciien
End Function