Option Explicit

Public Type MatchData
    FoundTxt As String
    NextPosition As Long
End Type


Sub TestFindTwoSubstringMatches()
    Dim str As String
    str = "Concentric Reducer s/80 x s/40"
    
    Dim srchItems As Variant
    srchItems = Split("S/40|S/80|S/160|S/60", "|")
    
    Dim results() As MatchData
    results = findTwoSubstringMatches(str, srchItems)
    
    Dim i As Long
    For i = 0 To UBound(results)
        Debug.Print "Match " & i + 1 & ": " & _
            "FoundTxt=" & results(i).FoundTxt & ", " & _
            "NextPosition=" & results(i).NextPosition
    Next i
End Sub



Function findTwoSubstringMatches(ByVal str As String, ByVal srchItems As Variant) As MatchData()
    Dim md(1) As MatchData
    Dim start_pos As Long
    Dim cur_pos As Long
    Dim earliest_pos As Long
    Dim i As Long
    Dim item As Variant
    
    start_pos = 1
    For i = 0 To 1
        earliest_pos = 0
        
        For Each item In srchItems
            cur_pos = InStr(start_pos, str, item, vbTextCompare)
            If cur_pos > 0 Then
                If earliest_pos = 0 Or cur_pos < earliest_pos Then
                    md(i).FoundTxt = item
                    md(i).NextPosition = cur_pos + Len(item)
                    earliest_pos = cur_pos
                End If
            End If
        Next item
        
        If md(i).NextPosition = 0 Or md(i).NextPosition > Len(str) Then Exit For
        start_pos = md(i).NextPosition
    Next i
    
    findTwoSubstringMatches = md
End Function

