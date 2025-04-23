Option Explicit
' filename:         GetweldWt.vba
'
' purpose:          return the weld weight for a given nomSize, sched, optional thick
'
' usage:            weld_wt = buttweld_wt(6, "XS", "")
'                   not that optional thickness is not used unless there is an
'                   odd thickness for a non standard schedule i.e. 1.25" thick pipe
'
' dependencies:     none
'
' By:               T.Sciple, 09/16/YYYY2024

Public Function buttweld_wt(nomSize As Double, sched As String, optional_thk As String)
    Dim PI As Double
    PI = 3.14159265358979
    Dim root_face As Double
    root_face = 0.0625   'reference weldbend catalog page 107
    Dim root_gap As Double
    If nomSize <= 6 Then
        root_gap = 1 / 8   'Root Gap Tolerance is an non-essential variable and shall; comply with the WPS. Normal Root Gap used is 2.4mm +/- 0.8mm and it depends on factors like Thickness, size of Filler wire used for root pass, welder's skill and no restriction as per ASME B 31.3.
    Else
        root_gap = 5 / 32
    End If
    Dim half_bevel_angle As Double
    half_bevel_angle = 75 / 2
    Dim pipe_rad As Double
    pipe_rad = pipe_od(nomSize) / 2
    Dim thk As Double
    thk = pipe_thk(nomSize, sched)

    'compute areas and volumes
    '1. Area of rectangle from root gap thru the thickness of the pipe
    Dim area1 As Double
    area1 = root_gap * thk
    Dim rad1 As Double
    rad1 = pipe_rad - thk / 2
    Dim vol1 As Double
    vol1 = 2 * PI * rad1 * area1

    '2. Area of both triangles from root face od of pipe at the bevel angle
    Dim leg2a As Double
    Dim leg2b As Double
    leg2a = thk - root_face
    leg2b = leg2a * Tan((half_bevel_angle) * PI / 180)
    Dim area2 As Double
    area2 = leg2a * leg2b
    Dim rad2 As Double
    rad2 = pipe_rad - leg2a / 3
    Dim vol2 As Double
    vol2 = 2 * PI * rad2 * area2
    
    '3. Compute the area of the outside weld cap by assuming it is equal to the area of a circular chord segment
    Dim sagitta_len As Double   'this is the math term for distance from chord line to outer radius,  height of weld
    If nomSize <= 8 Then
        sagitta_len = 0.0625
    Else
        sagitta_len = 0.125
    End If
    
    'calculate the estimated length of weld cap
    Dim chord_len As Double
    chord_len = 2 * leg2b + root_gap + 2 * sagitta_len
    Dim chord_radius As Double
    chord_radius = chord_len ^ 2 / (8 * sagitta_len) + sagitta_len / 2
    Dim chord_ang_radians As Double
    chord_ang_radians = 2 * Math.Arcsin(chord_len / (2 * chord_radius))
    Dim area3 As Double
    area3 = 0.5 * chord_radius ^ 2 * (chord_ang_radians - Sin(chord_ang_radians))
    Dim rad3 As Double
    rad3 = pipe_rad + sagitta_len / 2    'this is the mid point which is slightly conservative
    Dim vol3 As Double
    vol3 = 2 * PI * rad3 * area3
    
    'calculate the weight of the weld area
    Dim dens_steel As Double    'density of steel in lbs per cubic inch
    dens_steel = 0.2836
    buttweld_wt = (vol1 + vol2 + vol3) * dens_steel ' density of steel
End Function