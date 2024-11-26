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
' By:  T.Sciple, 11/25/2024


'Defined Global Constants for scoring types or Enums
Const SCORE_NO_PREV As Integer = 0      'No previous score unless on frame 1
Const SCORE_NO_FRAME_DATA As Integer = 1
Const SCORE_SPARE As Integer = 2
Const SCORE_STRIKE As Integer = 3
Const SCORE_OTHER As Integer = 4


Function calc_bowling_score(ByVal prev_score As Variant, _
                            ByVal frame_no As Long, _
                            ByVal frame_a As String, _
                            ByVal frame_b As String, _
                            ByVal frame_c As String) _
                            As Variant  'using variant to return a "" empty value
    
    Dim score_type As Integer
    score_type = get_score_type(prev_score, frame_no, frame_a)
    
    Dim tmp As Integer
    Dim num_rolls_to_get As Integer
    
    Select Case score_type
    
        Case SCORE_NO_PREV, SCORE_NO_FRAME_DATA
            GoTo report_no_score
        
        Case SCORE_SPARE
            num_rolls_to_get = 1
            tmp = get_next_roll(frame_no, frame_a, frame_b, frame_c, num_rolls_to_get)
            If tmp = -1 Then
                GoTo report_no_score
            Else
                tmp = tmp + 10
            End If
        
        Case SCORE_STRIKE
            num_rolls_to_get = 2    'default
            Dim dblStrikeInTenthFrame As Boolean
            Dim oneStrikeInTenthFrame As Boolean
            
            If frame_no = 10 And Left(frame_a, 3) = "X X" Then
                num_rolls_to_get = 1
                dblStrikeInTenthFrame = True
            ElseIf frame_no = 10 And Left(frame_a, 1) = "X" Then
                oneStrikeInTenthFrame = True
            End If
            
            tmp = get_next_roll(frame_no, frame_a, frame_b, frame_c, num_rolls_to_get)
            
            'calculate score depending on the various cases
            If tmp = -1 Then GoTo report_no_score
            
            tmp = Switch( _
                          dblStrikeInTenthFrame, tmp + 10 + 10, _
                          oneStrikeInTenthFrame, tmp + tmp + 10, _
                          True, tmp + 10)
        
        Case SCORE_OTHER
            tmp = get_frame_score(frame_a)
    
    End Select
    
    calc_bowling_score = CInt(prev_score) + tmp
    'exit if score was computed
    Exit Function
    
report_no_score:
    calc_bowling_score = ""

End Function


Function get_score_type(ByVal prev_score As Variant, _
                        ByVal frame_no As Long, _
                        ByVal frame_a As String _
                        ) As Integer

    get_score_type = Switch( _
                            prev_score = "" And frame_no <> 1, SCORE_NO_PREV, _
                            frame_a = "", SCORE_NO_FRAME_DATA, _
                            Mid(frame_a, 3, 1) = "/", SCORE_SPARE, _
                            Left(frame_a, 1) = "X", SCORE_STRIKE, _
                            True, SCORE_OTHER _
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