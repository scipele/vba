Option Explicit


Function lookForSimilarText1(lookupText, lookupRng As Range)
    Dim i As Integer
    Dim location As Integer
    For i = 1 To UBound(lookupRng())
        location = InStr(1, UCase(lookupRng(i)), UCase(lookupText), vbTextCompare)
        If location > 0 Then Exit For
    Next i
    If location > 0 Then
        lookForSimilarText1 = lookupRng(i).Value2
    Else
        lookForSimilarText1 = "Not Found"
    End If
End Function


Function lookForSimilarTextRow1(lookupText, lookupRng As Range)
    Dim i  As Integer
    Dim location As Integer
    For i = 1 To UBound(lookupRng())
        location = InStr(1, UCase(lookupRng(i)), UCase(lookupText), vbTextCompare)
        If location > 0 Then Exit For
    Next i
    If location > 0 Then
        lookForSimilarTextRow1 = i
    Else
        lookForSimilarTextRow1 = "Not Found"
    End If
End Function


Sub FindBestMatch()
    Dim JobData(1 To 1826, 1 To 5)
    Dim estimData(1 To 2784, 1 To 6)
    
    Dim i, j, k As Integer
    Dim MatchScore As Integer '1 to 4
    
    Workbooks("Merged Job List (For Comparison Only).xlsx").Worksheets("ParentJbList").Activate
    For i = 1 To UBound(JobData, 1)
        JobData(i, 1) = Range("A" & i + 1).Value 'Yr
        JobData(i, 2) = Range("B" & i + 1).Value 'Parent JobNo
        JobData(i, 3) = Range("C" & i + 1).Value 'Client
        JobData(i, 4) = Range("D" & i + 1).Value 'Job Desc
        JobData(i, 5) = Range("E" & i + 1).Value 'Total of Parent Job Cost
    Next i
    
    Workbooks("Estimate Log R1.xlsx").Worksheets("Log").Activate
    For i = 1 To UBound(estimData, 1)
        estimData(i, 1) = Left(Range("A" & i + 1).Value, 2) 'Yr taken as first two digits
        estimData(i, 2) = Range("A" & i + 1).Value 'Estim No
        estimData(i, 3) = Range("B" & i + 1).Value 'Client
        estimData(i, 4) = Range("D" & i + 1).Value 'Estim Desc
        estimData(i, 5) = Range("M" & i + 1).Value 'Estim Value
        estimData(i, 6) = Range("G" & i + 1).Value 'Status i.e. Pending
    Next i
    
    'start comparing the Pending Estimate Data and adjust match score to find the best possible match
    Workbooks("Merged Job List (For Comparison Only).xlsx").Worksheets("PotentialMatches").Activate
    k = 2
    For i = 1 To UBound(estimData, 1)
        MatchScore = 0
        If estimData(i, 6) = "Pending" Then
            'First Check to See if Company Matches
            For j = 1 To UBound(JobData, 1)
                'Make sure that Job Year is same or one year following Estimate Number
                If Val(JobData(j, 1)) >= Val(estimData(i, 1)) And Val(JobData(j, 1)) - Val(estimData(i, 1)) < 2 Then
                    'See if client names are similar
                    If Similarity(estimData(i, 3), JobData(j, 3)) > 0.35 Then
                        MatchScore = 1
                        'Now see if descriptions are similar
                        If Similarity(estimData(i, 4), JobData(j, 4)) > 0.15 Then
                            MatchScore = MatchScore + 1
                            'Now see if values are similar
                            If JobData(j, 5) > 0 And IsNumeric(estimData(i, 5)) And IsNumeric(JobData(j, 5)) Then
                                If estimData(i, 5) / JobData(j, 5) > 0.99 And estimData(i, 5) / JobData(j, 5) < 1.01 Then
                                    MatchScore = MatchScore + 1
                                End If
                            End If
                        
                        End If
                        If MatchScore >= 3 Then
                            k = k + 1
                            Range("A" & k).Value = JobData(j, 1) & "-" & JobData(j, 2) 'Parent Yr and Job No
                            Range("B" & k).Value = JobData(j, 3) 'Parent Job Client
                            Range("C" & k).Value = JobData(j, 4) 'Parent Job Desc
                            Range("D" & k).Value = JobData(j, 5) 'Parent Job Amt
                            
                            Range("E" & k).Value = estimData(i, 2) 'Parent Estim No
                            Range("F" & k).Value = estimData(i, 3) 'Parent Estim Client
                            Range("G" & k).Value = estimData(i, 4) 'Parent Estim Desc
                            Range("H" & k).Value = estimData(i, 5) 'Parent Estim Amt
                        End If
                    End If
                End If
            Next j
        End If
    Next i
    
    Erase JobData
    Erase estimData
End Sub


Public Function Similarity(ByVal String1 As String, _
    ByVal String2 As String, _
    Optional ByRef RetMatch As String, _
    Optional min_match = 1) As Single
    Dim b1() As Byte, b2() As Byte
    Dim lngLen1 As Long, lngLen2 As Long
    Dim lngResult As Long

    If UCase(String1) = UCase(String2) Then
        Similarity = 1
    Else:
        lngLen1 = Len(String1)
        lngLen2 = Len(String2)
        If (lngLen1 = 0) Or (lngLen2 = 0) Then
            Similarity = 0
        Else:
            b1() = StrConv(UCase(String1), vbFromUnicode)
            b2() = StrConv(UCase(String2), vbFromUnicode)
            lngResult = Similarity_sub(0, lngLen1 - 1, _
            0, lngLen2 - 1, _
            b1, b2, _
            String1, _
            RetMatch, _
            min_match)
            Erase b1
            Erase b2
            If lngLen1 >= lngLen2 Then
                Similarity = lngResult / lngLen1
            Else
                Similarity = lngResult / lngLen2
            End If
        End If
    End If
End Function


Private Function Similarity_sub(ByVal start1 As Long, ByVal end1 As Long, _
                                ByVal start2 As Long, ByVal end2 As Long, _
                                ByRef b1() As Byte, ByRef b2() As Byte, _
                                ByVal FirstString As String, _
                                ByRef RetMatch As String, _
                                ByVal min_match As Long, _
                                Optional recur_level As Integer = 0) As Long
    '* CALLED BY: Similarity *(RECURSIVE)
    Dim lngCurr1 As Long, lngCurr2 As Long
    Dim lngMatchAt1 As Long, lngMatchAt2 As Long
    Dim i As Long
    Dim lngLongestMatch As Long, lngLocalLongestMatch As Long
    Dim strRetMatch1 As String, strRetMatch2 As String
    
    If (start1 > end1) Or (start1 < 0) Or (end1 - start1 + 1 < min_match) _
    Or (start2 > end2) Or (start2 < 0) Or (end2 - start2 + 1 < min_match) Then
        Exit Function '(exit if start/end is out of string, or length is too short)
    End If
    
    For lngCurr1 = start1 To end1
        For lngCurr2 = start2 To end2
            i = 0
            Do Until b1(lngCurr1 + i) <> b2(lngCurr2 + i)
                i = i + 1
                If i > lngLongestMatch Then
                    lngMatchAt1 = lngCurr1
                    lngMatchAt2 = lngCurr2
                    lngLongestMatch = i
                End If
                If (lngCurr1 + i) > end1 Or (lngCurr2 + i) > end2 Then Exit Do
            Loop
        Next lngCurr2
    Next lngCurr1
    
    If lngLongestMatch < min_match Then Exit Function
    
    lngLocalLongestMatch = lngLongestMatch
    RetMatch = ""
    
    lngLongestMatch = lngLongestMatch _
                    + Similarity_sub(start1, lngMatchAt1 - 1, _
                      start2, lngMatchAt2 - 1, _
                      b1, b2, _
                      FirstString, _
                      strRetMatch1, _
                      min_match, _
                      recur_level + 1)
    If strRetMatch1 <> "" Then
        RetMatch = RetMatch & strRetMatch1 & "*"
    Else
        RetMatch = RetMatch & IIf(recur_level = 0 _
        And lngLocalLongestMatch > 0 _
        And (lngMatchAt1 > 1 Or lngMatchAt2 > 1) _
        , "*", "")
    End If
    
    RetMatch = RetMatch & Mid$(FirstString, lngMatchAt1 + 1, lngLocalLongestMatch)
    
    lngLongestMatch = lngLongestMatch _
    + Similarity_sub(lngMatchAt1 + lngLocalLongestMatch, end1, _
    lngMatchAt2 + lngLocalLongestMatch, end2, _
    b1, b2, _
    FirstString, _
    strRetMatch2, _
    min_match, _
    recur_level + 1)
    
    If strRetMatch2 <> "" Then
        RetMatch = RetMatch & "*" & strRetMatch2
    Else
        RetMatch = RetMatch & IIf(recur_level = 0 _
        And lngLocalLongestMatch > 0 _
        And ((lngMatchAt1 + lngLocalLongestMatch < end1) _
        Or (lngMatchAt2 + lngLocalLongestMatch < end2)) _
        , "*", "")
    End If
    
    Similarity_sub = lngLongestMatch
End Function