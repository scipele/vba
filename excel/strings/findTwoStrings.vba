Option Explicit
'| Item	        | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | findTwoStrings.vba                                          |
'| EntryPoint   | TestFindTwoSubstringMatches                                 |
'| Purpose      | Search for up to two instances of specified substrings      |
'| Inputs       | str - the string to search,                                 |
'|              | srchItemsStr - the substrings to search for,                |
'|              | splitDelim - the delimiter for substrings                   |
'| Outputs      | An array of MatchData w/ found substrings & positions       |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 03/20/2026                                        |


Public Type MatchData
    FoundTxt As String
    NextSearchPosition As Long
End Type


Sub TestFindTwoSubstringMatches()
    Dim str As String
    str = "Concentric Reducer s/80 x s/40"
    
    Dim results() As MatchData
    results = findTwoSubstringMatches(str, "S/40|S/80|S/160|S/60", "|")
    
    Dim i As Long
    For i = LBound(results) To UBound(results)
        Debug.Print "Match " & i + 1 & ": " & _
            "FoundTxt=" & results(i).FoundTxt & ", " & _
            "NextSearchPosition=" & results(i).NextSearchPosition
    Next i
End Sub


Function findTwoSubstringMatches(ByVal str As String, _
                                 ByVal srchItemsStr As Variant, _
                                 ByVal splitDelim As String) _
                                 As MatchData()
    Dim md(1) As MatchData
    
    Dim srchItems As Variant
    srchItems = Split(srchItemsStr, splitDelim)
    
    Dim start_pos As Long
    start_pos = 1
    Dim i As Long
    Dim earliest_pos As Long
    Dim item As Variant
    Dim cur_pos As Long
    For i = 0 To 1
        earliest_pos = 0
        
        For Each item In srchItems
            cur_pos = InStr(start_pos, str, item, vbTextCompare)
            If cur_pos > 0 Then
                If earliest_pos = 0 Or cur_pos < earliest_pos Then
                    md(i).FoundTxt = item
                    md(i).NextSearchPosition = cur_pos + Len(item)
                    earliest_pos = cur_pos
                End If
            End If
        Next item
        
        If md(i).NextSearchPosition = 0 Or md(i).NextSearchPosition > Len(str) Then Exit For
        start_pos = md(i).NextSearchPosition
    Next i
    
    findTwoSubstringMatches = md
End Function