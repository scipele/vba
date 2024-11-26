Option Explicit

' filename:     calc_bowling_score.vba
'
' Purpose:      custom formula for computing bowling score in excel, sample layout below
'                     A         B       C       D       E       F       G       H       I       J       K
'                 +---------+-------+-------+-------+-------+-------+-------+-------+-------+-------+-------+
'              1  | Frame-> |   1   |   2   |   3   |   4   |   5   |   6   |   7   |   8   |   9   |  10   |
'                 +---------+-------+-------+-------+-------+-------+-------+-------+-------+-------+-------+
'              2  |         |  7 2  |  3 /  |  3 6  |  8 1  |  6 /  |  8 /  |   X   |  9 -  |   X   |  8 -   |
'                 +  SEAN   +-------+-------+-------+-------+-------+-------+-------+-------+-------+-------+
'              3  |         |   9   |  22   |  31   |  40   |  58   |  78   |  97   |  106  |  124  |  132   |
'                 +---------+-------+-------+-------+-------+-------+-------+-------+-------+-------+-------+
'               Example Score Formula for Cell E3 '=calc_bowling_score(D3,E$1,E2,F2,G2)'
'
' Dependencies: None
'
' By:  T.Sciple, 11/26/2024


Private Enum scoreType
    stPrevScoreEmptyExceptFrameOne
    stFrameDataEmpty
    stSpare
    stStrike
    stOther
End Enum


Function calc_bowling_score(ByVal prev_score As Variant, _
                            ByVal frame_no As Long, _
                            ByVal frame_a As String, _
                            ByVal frame_b As String, _
                            ByVal frame_c As String) _
                            As Variant  'using variant to return a "" empty value
    
    Dim current_score_type As Integer
    current_score_type = get_score_type(prev_score, frame_no, frame_a)
    
    Dim frame_score As Integer
    Dim num_rolls_to_get As Integer
    
    Select Case current_score_type
    
        Case stPrevScoreEmptyExceptFrameOne, stFrameDataEmpty
            GoTo ReportNoScoreLabel
        
        Case stSpare
            num_rolls_to_get = 1
            frame_score = get_next_roll(frame_no, frame_a, frame_b, frame_c, num_rolls_to_get)
            If frame_score = -1 Then
                GoTo ReportNoScoreLabel
            Else
                frame_score = frame_score + 10
            End If
        
        Case stStrike
            num_rolls_to_get = 2    'default
            Dim dblStrikeInTenthFrame As Boolean
            Dim oneStrikeInTenthFrame As Boolean
            
            If frame_no = 10 And Left(frame_a, 3) = "X X" Then
                num_rolls_to_get = 1
                dblStrikeInTenthFrame = True
            ElseIf frame_no = 10 And Left(frame_a, 1) = "X" Then
                oneStrikeInTenthFrame = True
            End If
            
            frame_score = get_next_roll(frame_no, frame_a, frame_b, frame_c, num_rolls_to_get)
            
            'calculate score depending on the various cases
            If frame_score = -1 Then GoTo ReportNoScoreLabel
            
            frame_score = Switch( _
                          dblStrikeInTenthFrame, frame_score + 10 + 10, _
                          oneStrikeInTenthFrame, 2 * frame_score + 10, _
                          True, frame_score + 10)
        
        Case stOther
            frame_score = get_frame_score(frame_a)
    
    End Select
    
    calc_bowling_score = CInt(prev_score) + frame_score
    'exit if score was computed
    Exit Function
    
ReportNoScoreLabel:
    calc_bowling_score = ""

End Function


Function get_score_type(ByVal prev_score As Variant, _
                        ByVal frame_no As Long, _
                        ByVal frame_a As String _
                        ) As Integer

    get_score_type = Switch( _
                            prev_score = "" And frame_no <> 1, stPrevScoreEmptyExceptFrameOne, _
                            frame_a = "", stFrameDataEmpty, _
                            Mid(frame_a, 3, 1) = "/", stSpare, _
                            Left(frame_a, 1) = "X", stStrike, _
                            True, stOther _
                            )

End Function


Function get_frame_score(ByRef frame As String) As Variant

    Dim roll As Variant
    roll = Split(frame, " ")
    
    'convert any dashes to zeros
    If roll(0) = "-" Then roll(0) = 0
    If roll(1) = "-" Then roll(1) = 0

    get_frame_score = CInt(roll(0)) + CInt(roll(1))

End Function


Function get_next_roll( _
                        ByVal frame_no As Long, _
                        ByVal frame_a As String, _
                        ByVal frame_b As String, _
                        ByVal frame_c As String, _
                        ByVal num_rolls_to_add As Integer) As Integer
    
    Dim str As String
    
    'check to see if you are on the tenth frame where there are three possible scores)
    If frame_no = 10 And Len(frame_a) > 3 Then
        If num_rolls_to_add = 1 Then
            str = Right(frame_a, 1) & " "
        Else
            str = Right(frame_a, Len(frame_a) - 2) & " "
        End If
    Else
        str = frame_b & " " & frame_c
    End If
    
    If Len(str) <= num_rolls_to_add Then
        get_next_roll = -1  'return -1 if not enough data to score
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
        If indx > num_rolls_to_add Then
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
    
    get_next_roll = score
End Function