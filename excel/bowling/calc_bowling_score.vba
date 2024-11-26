Option Explicit

' filename:     CalcBowlingScore.vba
' EntryPoint:   CalcBowlingScore
' Purpose:      custom formula for computing bowling score in excel, sample layout below
'                       A         B       C       D       E       F       G       H       I       J       K
'                   +---------+-------+-------+-------+-------+-------+-------+-------+-------+-------+-------+
'                1  | Frame-> |   1   |   2   |   3   |   4   |   5   |   6   |   7   |   8   |   9   |  10   |
'                   +---------+-------+-------+-------+-------+-------+-------+-------+-------+-------+-------+
'                2  |         |  7 2  |  3 /  |  3 6  |  8 1  |  6 /  |  8 /  |   X   |  9 -  |   X   |  8 -   |
'                   +  SEAN   +-------+-------+-------+-------+-------+-------+-------+-------+-------+-------+
'                3  |         |   9   |  22   |  31   |  40   |  58   |  78   |  97   |  106  |  124  |  132   |
'                   +---------+-------+-------+-------+-------+-------+-------+-------+-------+-------+-------+
'                 Example Score Formula for Cell E3 '=CalcBowlingScore(D3,E$1,E2,F2,G2)'
'
' Inputs:       prevScore, frameNo, frameA, frameB, frameC
' Outputs:      calculated Score for each frame
' Dependencies: None
' By:  T.Sciple, 11/26/2024

Private Enum scoreType
    stPrevScoreEmptyExceptFrameOne
    stFrameDataEmpty
    stSpare
    stStrike
    stOther
End Enum


Public Function CalcBowlingScore(ByVal prevScore As Variant, _
                            ByVal frameNo As Long, _
                            ByVal frameA As String, _
                            ByVal frameB As String, _
                            ByVal frameC As String) _
                            As Variant  'using variant to return a "" empty value
    
    
    Dim current_score_type As Integer
    current_score_type = GetScoreType(prevScore, frameNo, frameA)
    
    Dim frame_score As Integer
    Dim num_rolls_to_get As Integer
    
    Select Case current_score_type
    
        Case stPrevScoreEmptyExceptFrameOne, stFrameDataEmpty
            GoTo LblReportNoScore
        
        Case stSpare
            num_rolls_to_get = 1
            frame_score = GetNextRollOrRolls(frameNo, frameA, frameB, frameC, num_rolls_to_get)
            If frame_score = -1 Then
                GoTo LblReportNoScore
            Else
                frame_score = frame_score + 10
            End If
        
        Case stStrike
            num_rolls_to_get = 2    'default
            Dim dbl_strike_in_tenth_frame As Boolean
            Dim one_strike_in_tenth_frame As Boolean
            
            If frameNo = 10 And Left(frameA, 3) = "X X" Then
                num_rolls_to_get = 1
                dbl_strike_in_tenth_frame = True
            ElseIf frameNo = 10 And Left(frameA, 1) = "X" Then
                one_strike_in_tenth_frame = True
            End If
            
            frame_score = GetNextRollOrRolls(frameNo, frameA, frameB, frameC, num_rolls_to_get)
            
            'calculate score depending on the various cases
            If frame_score = -1 Then GoTo LblReportNoScore
            
            frame_score = Switch( _
                          dbl_strike_in_tenth_frame, frame_score + 10 + 10, _
                          one_strike_in_tenth_frame, 2 * frame_score + 10, _
                          True, frame_score + 10)
        
        Case stOther
            frame_score = GetFrameScore(frameA)
    
    End Select
    
    CalcBowlingScore = CInt(prevScore) + frame_score
    'exit if score was computed
    Exit Function
    
LblReportNoScore:
    CalcBowlingScore = ""

End Function


Private Function GetScoreType(ByVal prevScore As Variant, _
                        ByVal frameNo As Long, _
                        ByVal frameA As String _
                        ) As Integer

    GetScoreType = Switch( _
                            prevScore = "" And frameNo <> 1, stPrevScoreEmptyExceptFrameOne, _
                            frameA = "", stFrameDataEmpty, _
                            Mid(frameA, 3, 1) = "/", stSpare, _
                            Left(frameA, 1) = "X", stStrike, _
                            True, stOther _
                            )
End Function


Private Function GetFrameScore(ByRef frame As String) As Variant

    Dim rolls As Variant
    rolls = Split(frame, " ")
    
    'convert any dashes to zeros
    If rolls(0) = "-" Then rolls(0) = 0
    If rolls(1) = "-" Then rolls(1) = 0

    GetFrameScore = CInt(rolls(0)) + CInt(rolls(1))

End Function


Private Function GetNextRollOrRolls( _
                        ByVal frameNo As Long, _
                        ByVal frameA As String, _
                        ByVal frameB As String, _
                        ByVal frameC As String, _
                        ByVal numRollsToAdd As Integer) As Integer
    
    Dim str As String
    
    'check to see if you are on the tenth frame where there are three possible scores)
    If frameNo = 10 And Len(frameA) > 3 Then
        If numRollsToAdd = 1 Then
            str = Right(frameA, 1) & " "
        Else
            str = Right(frameA, Len(frameA) - 2) & " "
        End If
    Else
        str = frameB & " " & frameC
    End If
    
    If Len(str) <= numRollsToAdd Then
        GetNextRollOrRolls = -1  'return -1 if not enough data to score
        Exit Function
    End If
    
    Dim rolls As Variant
    rolls = Split(str, " ")
    Dim roll As Variant
    Dim score As Integer
    score = 0
    Dim indx As Integer
    indx = 0
    For Each roll In rolls
        indx = indx + 1
        If indx > numRollsToAdd Then
            Exit For
        End If
        
        If roll = "X" Then
            score = score + 10
        ElseIf roll = "/" Then
            score = 10
        ElseIf roll = "-" Then
            score = score + 0
        Else
            score = score + roll
        End If
    Next roll
    
    GetNextRollOrRolls = score
End Function