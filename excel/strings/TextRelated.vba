'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | Strings.vba                                                 |
'| Purpose      | various subs for string manipulation                        |
'| By Name,Date | T.Sciple, 12/14/2024                                        |
'|--------------|-------------------------------------------------------------|
'| Listing of Functions in thie module:                                       |
'|     Function GetRatingLargeOrSmallBore                                     |
'|     Function GetRatingSmallBore                                            |
'|     Function GetValveApiTrim                                               |
'|     Function GetValveBodyMatl                                              |
'|     Function GetValveOpererator                                            |
'|     Function GetValveType                                                  |
'|     Function GetValveWedgeType                                             |
'|     Function GetWeldSchedule                                               |
'|     Function OddNo                                                         |
'|     Function PadLeadZeros                                                  |
'|     Function PadTrailSpaces                                                |
'|     Function RemoveMultSpaces                                              |
'|     Function asciien                                                       |
'|     Function findInReverse                                                 |
'|     Function getTextBetween                                                |
'|     Function getTextLftOf                                                  |
'|     Function getTextRghtOf                                                 |
'|     Sub CreateHyperLink                                                    |
'|     Sub MergeContentsWithCarriageReturn                                    |
'|     Sub MergeContentsWithComma                                             |
'|     Sub MergeContentsWithSpace                                             |
'|     Sub MergeContentsWithoutCarriageReturn                                 |
'|     Sub PasteDn                                                            |
'|     Sub RemoveAnyLineFeedChar                                              |
'|     Sub SplitCellsByDelimiter                                              |
'|     Sub TrimLeadingAndTrailingSpaces                                       |
'|     Sub UpperC                                                             |
'|     Sub addApst                                                            |
'|     Sub lowerC                                                             |
'|     Sub properC                                                            |


'**************************************************************
'************************* TEXT PARTS ************************* 
'**************************************************************

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


'***************************************************************************
'************************* SPLIT AND MERGE CELLS *************************** 
'***************************************************************************
Sub SplitCellsByDelimiter()
    ' Prompt the user for the delimiter
    Dim delimiter As Variant
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
    Dim cell As Range
    Dim splitValues As Variant
    For Each cell In Selection
        ' Split the cell value by the delimiter
        splitValues = Split(cell.value, delimiter)
        ' Clear the current cell
        cell.value = ""
        ' Place the split values in the cell and adjacent cells to the right
        Dim i As Integer
        For i = LBound(splitValues) To UBound(splitValues)
            cell.Offset(0, i).value = splitValues(i)
        Next i
    Next cell
End Sub


Public Sub MergeContentsWithCarriageReturn()
    Dim i As Long
    Dim cell as Range
    For Each cell In Selection
        i = i + 1
        ActiveCell.value = Switch(i = 1, cell.value & vbCrLf, _
                                  i < Selection.Count, ActiveCell.value & cell.value & vbCrLf, _
                                  i = Selection.Count, ActiveCell.value & cell.value)
        If i > 1 Then cell.value = ""
    Next cell
End Sub


Public Sub MergeContentsWithSpace()
    Dim i As Long
    Dim cell As Range
    For Each cell In Selection
        i = i + 1
        ActiveCell.value = Switch(i = 1, cell.value & " ", _
                                  i < Selection.Count, ActiveCell.value & cell.value & " ", _
                                  i = Selection.Count, ActiveCell.value & cell.value)
        If i > 1 Then cell.value = ""
    Next cell
End Sub


Public Sub MergeContentsWithoutCarriageReturn()
    Dim i As Long
    Dim cell as Range
    For Each cell In Selection
        i = i + 1
        ActiveCell.value = Switch(i = 1, cell.value, _
                        i <= Selection.Count, ActiveCell.value & cell.value)
        If i > 1 Then cell.value = ""
    Next cell
End Sub


Public Sub MergeContentsWithComma()
    Dim i As Long
    Dim cell as Range
    For Each cell In Selection
        i = i + 1
        ActiveCell.value = Switch(i = 1, cell.value & ",", _
                        i < Selection.Count, ActiveCell.value & cell.value & ",", _
                        i = Selection.Count, ActiveCell.value & cell.value)
        If i > 1 Then cell.value = ""
    Next cell
End Sub


'**************************************************************************************
'************************* RETURN RATINGS FOR PIPE COMPONENTS ************************* 
'**************************************************************************************
Public Function GetValveOpereratorType(str As String) As String
    Dim tmp As Variant  'used variant to handle the null if not found
    tmp = Switch(InStr(1, str, ", HW", vbTextCompare) > 0, "HW", _
        InStr(1, str, "Gear", vbTextCompare) > 0, "Gear")
    If IsNull(tmp) Then tmp = ""
    GetValveOpereratorType = tmp
End Function


Public Function GetValveWedgeType(str As String) As String
    Dim tmp As Variant  'used variant to handle the null if not found
    tmp = Switch(InStr(1, str, "Flex Wdg", vbTextCompare) > 0, "Flex Wdg", _
                 InStr(1, str, "Flexible Wedge", vbTextCompare) > 0, "Flex Wdg", _
                 InStr(1, str, "Sol Wdg", vbTextCompare) > 0, "Sol Wdg", _
                 InStr(1, str, "Solid Wdg", vbTextCompare) > 0, "Solid Wdg")
    If IsNull(tmp) Then tmp = ""
    GetValveWedgeType = tmp
End Function


Public Function GetValveType(str As String) As String
    Dim tmp As Variant  'used variant to handle the null if not found
    tmp = Switch(InStr(1, str, "GAT", vbTextCompare) > 0, "GATE", _
                 InStr(1, str, "BALL", vbTextCompare) > 0, "BALL", _
                 InStr(1, str, "CHECK", vbTextCompare) > 0, "CHECK", _
                 InStr(1, str, "CHK", vbTextCompare) > 0, "CHECK", _
                 InStr(1, str, "NEEDLE", vbTextCompare) > 0, "NEEDLE", _
                 InStr(1, str, "GLOBE", vbTextCompare) > 0, "GLOBE")
    If IsNull(tmp) Then tmp = "not found"
    GetValveType = tmp
End Function


Public Function GetValveBodyMatl(str As String) As String
    Dim tmp As Variant  'used variant to handle the null if not found
    tmp = Switch(InStr(1, str, "A216 WCB", vbTextCompare) > 0, "A216-WCB", _
                 InStr(1, str, "A352 LCC", vbTextCompare) > 0, "A352-LCC", _
                 InStr(1, str, "A105", vbTextCompare) > 0, "A105", _
                 InStr(1, str, "A216 WCB", vbTextCompare) > 0, "A216-WCB", _
                 InStr(1, str, "A350 LF2 CI 1", vbTextCompare) > 0, "A350-LF2-CL1", _
                 InStr(1, str, "A350 LF2 CI 1", vbTextCompare) > 0, "A350-LF2-CL1", _
                 InStr(1, str, "A350", vbTextCompare) > 0, "A350-LF2-CL1")  ' address typo missing info
    If IsNull(tmp) Then tmp = "not found"
    GetValveBodyMatl = tmp
End Function


Public Function GetValveApiTrim(str As String) As String
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


Public Function GetRatingLargeOrSmallBore(str As String) As String
    Dim tmp As Variant  'used variant to handle the null if not found
    tmp = Switch(InStr(1, str, "1500", vbTextCompare) > 0, "1500", _
                 InStr(1, str, "300", vbTextCompare) > 0, "300", _
                 InStr(1, str, "600", vbTextCompare) > 0, "600", _
                 InStr(1, str, "800", vbTextCompare) > 0, "800", _
                 InStr(1, str, "900", vbTextCompare) > 0, "900", _
                 InStr(1, str, "150", vbTextCompare) > 0, "150", _
                 InStr(1, str, "2500", vbTextCompare) > 0, "2500")
    If IsNull(tmp) Then tmp = "not found"
    GetRatingLargeOrSmallBore = tmp
End Function


Public Function GetRatingLargeBore(str As String) As String
    Dim tmp As Variant  'used variant to handle the null if not found
    tmp = Switch(InStr(1, str, "1500", vbTextCompare) > 0, "1500", _
                 InStr(1, str, "300", vbTextCompare) > 0, "300", _
                 InStr(1, str, "600", vbTextCompare) > 0, "600", _
                 InStr(1, str, "900", vbTextCompare) > 0, "900", _
                 InStr(1, str, "150", vbTextCompare) > 0, "150", _
                 InStr(1, str, "2500", vbTextCompare) > 0, "2500")
    If IsNull(tmp) Then tmp = "not found"
    GetRatingLargeBore = tmp
End Function


Function GetRatingSmallBore(str As String) As String
    Dim tmp As Variant  'used variant to handle the null if not found
    tmp = Switch(InStr(1, str, "1500", vbTextCompare) > 0, "1500", _
                 InStr(1, str, "800", vbTextCompare) > 0, "800")
    If IsNull(tmp) Then tmp = "not found"
    GetRatingSmallBore = tmp
End Function


Function GetWeldSchedule(str As String) As String
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
    GetWeldSchedule = tmp
End Function

'****************************************************************
'************************* STRING CASES ************************* 
'****************************************************************
Public Sub UpperC()
    Dim cell as Range
    For Each cell In Selection
        cell.value = UCase(cell.value)
    Next cell
End Sub


Public Sub lowerC()
    Dim cell as Range
    For Each cell In Selection
        cell.value = LCase(cell.value)
    Next cell
End Sub


Public Sub properC()
    Dim words As Variant
    Dim word As Variant
    Dim cell As Range
    Dim prop_str As String
    
    For Each cell In Selection
        prop_str = ""
        
        words = Split(LCase(cell.value))
        For Each word In words
            word = UCase(Left(word, 1)) & Right(word, Len(word) - 1)
            If prop_str = "" Then
                prop_str = word
            Else
                prop_str = prop_str & " " & word
            End If
        Next word
        cell.value = prop_str
    Next cell
End Sub

'****************************************************************
'************************* STRING MODS ************************** 
'****************************************************************
Public Sub addApst()
    For Each cell In Selection
        cell.Value = "'" & cell.Value
    Next cell
End Sub


Public Sub RemoveAnyLineFeedChar()
    Dim cell As Range
    Dim str As String
    ' Iterate through each selected cell
    For Each cell In Selection
        str = cell.value
        ' Replace all occurrences of vbCrLf, vbCR, and vbLF
        str = Replace(str, vbCrLf, " ") ' Replace combined carriage return + line feed
        str = Replace(str, vbCr, " ")   ' Replace carriage return
        str = Replace(str, vbLf, " ")   ' Replace line feed
        cell.value = str
    Next cell
End Sub


Public Sub TrimLeadingAndTrailingSpaces()
    Dim cell As Range
    For Each cell In Selection
        cell.value = Trim(cell.value)
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


Public Function OddNo(ByVal x As Long)
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

'****************************************************************
'************************* HYPERLINKS *************************** 
'****************************************************************
Public Sub CreateHyperLink()
    
    'Selected Range Cells are copied down to blank cells below with identical values
    'Example Usage: with selection on A1:B2
    '+-----+------------------+--------------------+----------------------------------+
    '|     |         A        |         B          |                  C               |
    '+-----+------------------+--------------------+----------------------------------+
    '|  1  |  my_link_name1   | c:\t\test.my_file1 |    my_link_name1<-placed here    |
    '|  2  |  my_link_name2   | c:\t\test.my_file2 |    my_link_name2<-placed here    |
    '+-----+------------------+--------------------+----------------------------------+
    
    'set the starting location for your hyperlinks (2 columns over from the staring spot)
    Dim startCol As Long
    startCol = Selection.Column + 2
    
    Dim startRow As Long
    startRow = Selection.row
    
    Dim startColLtr As String
    startColLtr = Split(Cells(startRow, startCol).Address, "$")(1)
    
    Dim LinkAry As Variant
    LinkAry = Selection
    
    Dim curRow As Long
    curRow = startRow
    
    'loop thru each item in the selection and create the hyperlink from each
    Dim i As Long
    Dim curtxt As String
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