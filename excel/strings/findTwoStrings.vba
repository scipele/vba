Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | findTwoMatches.vba                                          |
'| EntryPoint   | TestFindTwoSubstringMatches                                 |
'| Purpose      | Search for up to two instances of specified substrings      |
'| Inputs       | str - the string to search,                                 |
'|              | srchItemsStr - the substrings to search for,                |
'|              | splitDelim - the delimiter for substrings                   |
'| Outputs      | An array of MatchData w/ found substrings & positions       |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 03/24/2026                                        |


Public Type MatchData
    MatchingId As Long
    NextSearchPosition As Integer
End Type


Sub FindTwoSubstringMatches()
    
    Dim rules As Variant
    rules = Split("S/40|S/60|S/80|S/120|S/160", "|")
    
    Dim test_strings As Variant
    test_strings = Split("Concentric Reducer s/80 x s/40|Elbow s/120", "|")
    
    Dim md As MatchData
    Dim str As Variant
    
    For Each str In test_strings
        
        Debug.Print "string searched", str
        
        ' === First match (earliest position) ===
        md = mcMatcher.FindEarliestPositionMatch(CStr(str), rules)
        Debug.Print vbTab & "Earliest match : " & mcMatcher.GetMatchedRule(md.MatchingId, rules) & _
                    "  (next search position " & md.NextSearchPosition & ")"
        
        ' === Second match after first one ===
        Dim secondId As Long
        secondId = mcMatcher.FindIndexOfAnyMatch(CStr(str), md.NextSearchPosition, rules)
        
        Debug.Print vbTab & "Second match   : " & mcMatcher.GetMatchedRule(secondId, rules)
        Debug.Print String(60, "-")
    Next str
    
End Sub


Sub FindAnySubstringMatch()
    
    Dim rules As Variant
    rules = Split("S/40|S/60|S/80|S/120|S/160", "|")
    
    Dim test_strings As Variant
    test_strings = Split("Concentric Reducer s/80 x s/40|Elbow s/120", "|")
    
    Dim md As MatchData
    Dim str As Variant
    
    For Each str In test_strings
        
        Debug.Print "string searched", str
        ' === Single Match
        Dim secondId As Long
        secondId = FindIndexOfAnyMatch(CStr(str), 1, rules)
        
        Debug.Print vbTab & "Single match   : " & mcMatcher.GetMatchedRule(secondId, rules)
        Debug.Print String(60, "-")
    Next str
    
End Sub


Public Function FindEarliestPositionMatch(ByVal searched_str As String, _
                                           ByRef rules As Variant) _
                                           As MatchData

    Dim cur_pos As Integer
    Dim earliest_pos As Integer
    Dim earliest_match_index As Long
    
    Dim td As MatchData
    
    td.MatchingId = -1
    td.NextSearchPosition = 0
    
    Dim i As Long
    For i = LBound(rules) To UBound(rules)
        cur_pos = InStr(1, searched_str, rules(i), vbTextCompare)
        If cur_pos > 0 Then
            If earliest_pos = 0 Or cur_pos < earliest_pos Then
                earliest_pos = cur_pos
                earliest_match_index = i
                td.MatchingId = i
                td.NextSearchPosition = cur_pos + Len(rules(i))
            End If
        End If
    Next i
    
    FindEarliestPositionMatch = td
End Function


Public Function FindIndexOfAnyMatch(ByVal searched_str As String, _
                                    ByVal start_pos As Integer, _
                                    ByRef rules As Variant) _
                                    As Long

    start_pos = IIf(start_pos = 0, 1, start_pos)
    
    Dim cur_pos As Long
    Dim i As Long
    For i = LBound(rules) To UBound(rules)
        cur_pos = InStr(start_pos, searched_str, rules(i), vbTextCompare)
        If cur_pos > 0 Then
            FindIndexOfAnyMatch = i
            Exit Function
        End If
    Next i
    
    FindIndexOfAnyMatch = -1
End Function


Public Function GetMatchedRule(ByVal id As Long, _
                        ByRef rules As Variant) _
                        As String
    If id = -1 Then
        GetMatchedRule = "Not Found"
    Else
        GetMatchedRule = rules(id)
    End If
End Function