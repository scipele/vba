Option Explicit

' filename:     calc_bowling_score.vba
'
' Purpose:      custom formula for computing bowling score in excel
'
'         A         B       C       D       E       F       G       H       I       J       K
'      +---------+-------+-------+-------+-------+-------+-------+-------+-------+-------+-------+
'   1  | Frame-> |   1   |   2   |   3   |   4   |   5   |   6   |   7   |   8   |   9   |  10   |
'      +---------+-------+-------+-------+-------+-------+-------+-------+-------+-------+-------+
'   2  |         |  7 2  |  3 /  |  3 6  |  8 1  |  6 /  |  8 /  |   X   |  9 -  |   X   |  8 -   |
'      +  SEAN   +-------+-------+-------+-------+-------+-------+-------+-------+-------+-------+
'   3  |         |   9   |  22   |  31   |  40   |  58   |  78   |  97   |  106  |  124  |  132   |
'      +---------+-------+-------+-------+-------+-------+-------+-------+-------+-------+-------+
'
' Example Score Formula for Cell E3 '=calc_bowling_score(D3,E$1,E2,F2,G2)'
'
' Dependencies: None
'
' By:  T.Sciple, 11/22/2024


Function calc_bowling_score(ByVal prev_score As Variant, _
                            ByVal frame_no As Long, _
                            ByVal frame_a As String, _
                            ByVal frame_b As String, _
                            ByVal frame_c As String) _
                            As Variant  'using variant to return a "" empty value
    
    Dim score_type As Integer
    score_type = get_score_type(prev_score, frame_no, frame_a)
    
    Dim tmp As Integer
    Dim num_rolls_to_get
    
    Select Case score_type
    
        Case 0  'No previous score unless on frame 1
            calc_bowling_score = ""
            Exit Function
    
        Case 1  'No frame data
            calc_bowling_score = ""
            Exit Function
        
        Case 2  'for strike
            num_rolls_to_get = 2
            tmp = get_next_roll(frame_no, frame_a, frame_b, frame_c, num_rolls_to_get)
            If tmp = -1 Then
                calc_bowling_score = ""
                Exit Function
            Else
                tmp = tmp + 10
            End If
        
        Case 3  'for spare
            num_rolls_to_get = 1
            tmp = get_next_roll(frame_no, frame_a, frame_b, frame_c, num_rolls_to_get)
            If tmp = -1 Then
                calc_bowling_score = ""
                Exit Function
            Else
                tmp = tmp + 10
            End If
        
        Case 4  'strike on 10th frame roll 1 and 2
            num_rolls_to_get = 1
            tmp = get_next_roll(frame_no, frame_a, frame_b, frame_c, num_rolls_to_get)
            If tmp = -1 Then
                calc_bowling_score = ""
                Exit Function
            Else
                tmp = tmp + 10 + 10
            End If
        
        Case 5  'strike on 10th frame roll 1
            num_rolls_to_get = 2
            tmp = get_next_roll(frame_no, frame_a, frame_b, frame_c, num_rolls_to_get)
            If tmp = -1 Then
                calc_bowling_score = ""
                Exit Function
            Else
                tmp = 10 + tmp + tmp
            End If
        
        Case 6  'for other
            tmp = get_frame_score(frame_a)
    
    End Select
    
    calc_bowling_score = CInt(prev_score) + tmp

End Function


Function get_score_type(ByVal prev_score As Variant, _
                        ByVal frame_no As Long, _
                        ByVal frame_a As String _
                        ) As Integer

    'return values
    '0  No prev_score unless on frame 1
    '1  No current frame data
    '2  for strike
    '3  for spare
    '4  strike on 10th frame roll 1 and 2
    '5  strike on 10th frame roll 1
    '6  for other normal scores
    
    get_score_type = Switch( _
                            prev_score = "" And frame_no <> 1, 0, _
                            frame_a = "", 1, _
                            frame_a = "X", 2, _
                            Mid(frame_a, 3, 1) = "/", 3, _
                            frame_no = 10 And Left(frame_a, 3) = "X X", 4, _
                            frame_no = 10 And Left(frame_a, 1) = "X", 5, _
                            True, 6)

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
                        ByVal num_rolls_to_add As Integer) _
                        As Integer
    
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