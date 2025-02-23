Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | bowling_lane.vba                                            |
'| EntryPoint   | Function calls from spreadsheet                             |
'| Purpose      | compute bowling impact momentum calcs, and aim points       |
'| Inputs       | from function parameters                                    |
'| Outputs      | varies                                                      |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 02/23/2025                                        |


Private Enum returnType
    rtVb2           '0
    rtVp2           '1
    rtball_angle2   '2
End Enum


Sub test()
    
    Const mb As Double = 16
    Const mp As Double = 3.5
    Const e As Double = 0.9
    Const impactAngle As Double = -45
    Const vb1 As Double = 19
    Const vp1 As Double = 0#
    Const ball_angle1 As Double = 90

    Const return_type As Integer = rtVb2

    Dim result As Variant
    result = solveImpactEquation(mb, mp, vb1, e, impactAngle, vp1, ball_angle1, return_type)

End Sub


Public Function solveImpactEquation(ByVal mb As Double, _
                                      ByVal mp As Double, _
                                      ByVal vb1 As Double, _
                                      ByVal e As Double, _
                                      ByVal impactAngle As Double, _
                                      ByVal vp1 As Double, _
                                      ByVal ballAngle1 As Double, _
                                      ByVal returnTypeParam As Integer) _
                                      As Variant

    ' calculate the ball and pin speed after the collision, and aalso the angle of the ball and pin after the impact
    ' the x axis is drawn for convenience along the line of impact since that is the only direction of force that
    ' affects the change in momentum for the ball and pin
    '
    ' Break the initial velocities into x and y components
    Dim vbx1 As Double, vby1 As Double, vpx1 As Double, vpy1 As Double
    
    vbx1 = vb1 * Cos(radians(impactAngle))
    vby1 = vb1 * Sin(radians(impactAngle))
    vpx1 = 0
    vpy1 = 0
    
    ' Equation 1:  conservation of momentum in the x direction gives equation
    ' mb (vbx2) + mp (vpx2) = mb (vbx1) + mp (vpx1)
    ' which can be reduced to a standard linear equation in form as follows
    ' mb (vbx2) + mp (vpx2) = calc knowns
    
    ' a1 (x) + b1 (y) = c1
     
     Dim a1 As Double, b1 As Double, c1 As Double, a2 As Double, b2 As Double, c2 As Double
     a1 = mb
     b1 = mp
     c1 = mb * vbx1 + mp * vpx1
     
    ' Equation 2: coefficient of restitution formula is
    '         vpx2 - vbx2
    ' e =    -------------
    '         vbx1 - vpx1
    '
    ' rearranging this in standard form consistant with equation 1 gives
    ' 1 * (vbx2) + -1 (vpx2) = ( e * (vbx1 - vpx1) )
    ' or
    ' a2 (x) + b2 (y) = c2
    
    a2 = -1
    b2 = 1
    c2 = e * (vbx1 - vpx1)
    
    Dim result As Variant
    result = solveTwoLinearEquations(a1, b1, c1, a2, b2, c2)
    
    Dim vbx2 As Double, vpx2 As Double
    vbx2 = result(0)
    vpx2 = result(1)
    
    ' Equation 3: conservation of momentum in the y direction
    ' mb (vby1) = mb (vpy2)
    ' mp (vby1) = mp (vby2)
    ' hence
    Dim vby2 As Double, vpy2 As Double
    vby2 = vby1
    vpy2 = vpy1
    
    ' Compute resultant velocitys
    Dim vb2 As Double, vp2 As Double
    vb2 = (vbx2 ^ 2 + vby2 ^ 2) ^ 0.5
    vp2 = (vpx2 ^ 2 + vpy2 ^ 2) ^ 0.5
        
   
    Dim ball_angle2 As Double
    ball_angle2 = ballAngle1 - (impactAngle - degree(Atn(vby2 / vbx2)))
    
    'Debug.Print "vbx2....... = ", vbx2
    'Debug.Print "vby2....... = ", vby2
    'Debug.Print "vpx2....... = ", vpx2
    'Debug.Print "vpy2....... = ", vpy2
    'Debug.Print "vb2........ = ", vb2
    'Debug.Print "vp2........ = ", vp2
    'Debug.Print "ball_angle2 = ", ball_angle2
    
    Select Case returnTypeParam
        Case rtVb2
            solveImpactEquation = vb2
        
        Case rtVp2
            solveImpactEquation = vp2
        
        Case rtball_angle2
            solveImpactEquation = ball_angle2
        
        Case Else
            solveImpactEquation = -1

    End Select

End Function


Private Function solveTwoLinearEquations(ByVal a1 As Double, _
                                         ByVal b1 As Double, _
                                         ByVal c1 As Double, _
                                         ByVal a2 As Double, _
                                         ByVal b2 As Double, _
                                         ByVal c2 As Double) _
                                         As Variant

    ' Coefficients of the equations in lear form
    ' a1*x + b1*y = c1: "
    ' a2*x + b2*y = c2: "
    
    Dim D As Double
    D = a1 * b2 - a2 * b1

    ' Check if the determinant is zero (no unique solution)
    If D = 0 Then
        solveTwoLinearEquations = Array("No Unique Solution")
        Exit Function
    End If
    
    ' Solve for x and y
    Dim results As Variant
    ReDim results(0 To 1)
    results(0) = (b2 * c1 - b1 * c2) / D    'solve for x
    results(1) = (a1 * c2 - a2 * c1) / D    'solve for y
    
    solveTwoLinearEquations = results
    
End Function


Public Function getAimBoardAtArrows(ByVal rt_foot_board As Double, _
                                    ByVal target_x As Double, _
                                    ByVal target_y As Double, _
                                    ByVal start_line_x_in As Double) _
                                    As Double
                                    
    Const right_foot_offset_brds As Double = 6
    Const nom_16ft_arrow_x As Double = 336  'Tip of Center Arrow is 16' from the foul line

    'perform basic x/y ratio delta calculations to obtain the 'y' value at the 16' line
    Dim start_y_in As Double, target_x_delta As Double, target_y_delta As Double, nom_16_y As Double, target_line_angle_deg As Double, start_shot_board As Double
    
    start_shot_board = rt_foot_board - right_foot_offset_brds
    start_y_in = (start_shot_board - 0.5) * 41.5 / 39
    
    target_x_delta = target_x - start_line_x_in
    target_y_delta = start_y_in - target_y
    nom_16_y = start_y_in - (nom_16ft_arrow_x / target_x_delta * target_y_delta)
    target_line_angle_deg = degree(Atn(target_x_delta / target_y_delta))
    
    ' next determine the slight offset based on the triangle that is made at the tips of the arrows
    ' and the target line using the law of sines
    ' b and c are sides of the triangle
    ' angA, angC are the angles that are opposite of the respective sides
    
    Dim b As Double, c As Double, g As Double, angC As Double, angG As Double
    Const angB As Double = 48.43492346  ' Angle formed by the arrow tips ( DEGREE(ATAN(15.96"/18") )
    

    Const mid_board_x_in As Double = 20.75
    
    c = nom_16_y - mid_board_x_in
    angG = Abs(90 - target_line_angle_deg)
    angC = 180 - ((angG + 90) + angB)
    ' compute side b with law of sins
    b = c / Sin(radians(angC)) * Sin(radians(angB))
    
    ' use normal trig to solve this triangle
    g = b * Sin(radians(angG))
    
    If (nom_16_y <= 20.75) Then g = -1 * g
    Dim aim_y_at_arrows_in As Double
    aim_y_at_arrows_in = nom_16_y + g
    getAimBoardAtArrows = aim_y_at_arrows_in * 39 / 41.5 + 0.5

    'Debug.Print "rt_foot_board........ = ", rt_foot_board
    'Debug.Print "start_shot_board..... = ", start_shot_board
    'Debug.Print "target_x............. = ", target_x
    'Debug.Print "target_y............. = ", target_y
    'Debug.Print "start_line_x_in...... = ", start_line_x_in
    'Debug.Print "start_y_in........... = ", start_y_in
    'Debug.Print "target_x_delta....... = ", target_x_delta
    'Debug.Print "target_y_delta....... = ", target_y_delta
    'Debug.Print "nom_16_y............. = ", nom_16_y
    'Debug.Print "target_line_angle_deg = ", target_line_angle_deg
    'Debug.Print "angC................. = ", angC
    'Debug.Print "angG................. = ", angG
    'Debug.Print "b.................... = ", b
    'Debug.Print "c.................... = ", c
    'Debug.Print "g.................... = ", g
    'Debug.Print "aim_y_at_arrows_in... = ", aim_y_at_arrows_in
    'Debug.Print "getAimBoardAtArrows.. = ", getAimBoardAtArrows
    'Debug.Print
End Function


Private Function degree(ByVal angle_radians As Double) As Double
    Const pi As Double = 3.14159265358979
    degree = angle_radians * 180 / pi
End Function
                                    

Private Function radians(ByVal angle_degrees As Double) As Double
    Const pi As Double = 3.14159265358979
    radians = angle_degrees * pi / 180
End Function