'Define functions that retrieve/calculate data within the pipeData Class Module
Public Function pipe_od(size As Double) As Double
    Dim pip As pipeData 'dimension a variable of type of custom class 'pipeData'
    Set pip = New pipeData
    pipe_od = pip.od(size)
End Function


Public Function pipe_id(size As Double, schedule As String) As Double
    Dim pip As pipeData 'dimension a variable of type of custom class 'pipeData'
    Set pip = New pipeData
    pipe_id = pip.id(size, schedule)
End Function


Public Function pipe_thk(size As Double, schedule As String) As Double
    Dim pip As pipeData 'dimension a variable of type of custom class 'pipeData'
    Set pip = New pipeData
    pipe_thk = pip.thk(size, schedule)
End Function


Public Function pipe_area_metal_sqin(size As Double, schedule As String) As Double
    Dim pip As pipeData 'dimension a variable of type of custom class 'pipeData'
    Set pip = New pipeData
    pipe_area_metal_sqin = pip.area_metal_sqin(size, schedule)
End Function


Public Function pipe_wt_empty_lbs_per_ft(size As Double, schedule As String) As Double
    Dim pip As pipeData 'dimension a variable of type of custom class 'pipeData'
    Set pip = New pipeData
    pipe_wt_empty_lbs_per_ft = pip.wt_empty_lbs_per_ft(size, schedule)
End Function


Public Function pipe_wt_full_lbs_per_ft(size As Double, schedule As String) As Double
    Dim pip As pipeData 'dimension a variable of type of custom class 'pipeData'
    Set pip = New pipeData
    pipe_wt_full_lbs_per_ft = pip.wt_full_h20_lbs_per_ft(size, schedule)
End Function


Public Function pipe_moment_of_inertia_in4(size As Double, schedule As String) As Double
    Dim pip As pipeData 'dimension a variable of type of custom class 'pipeData'
    Set pip = New pipeData
    pipe_moment_of_inertia_in4 = pip.moment_of_inertia_in4(size, schedule)
End Function


Public Function pipe_section_modulus_in3(size As Double, schedule As String) As Double
    Dim pip As pipeData 'dimension a variable of type of custom class 'pipeData'
    Set pip = New pipeData
    pipe_section_modulus_in3 = pip.section_modulus_in3(size, schedule)
End Function


Public Function pipe_paint_area_sf_per_ft(size As Double) As Double
    Dim pip As pipeData 'dimension a variable of type of custom class 'pipeData'
    Set pip = New pipeData
    
    pipe_paint_area_sf_per_ft = pip.paint_area_sf_per_ft(size)
End Function


Sub SetFunctionDescription()
    Application.MacroOptions Macro:="pipedata", Description:="Returns pipe data given size and schedule" & vbCrLf & _
    "Parameters:" & vbCrLf _
    & "   nps (Nominal Pipe Size)" & vbCrLf _
    & "   sch (Schedule format xs, 40, 80)  " & vbCrLf _
    & "   returnType (Type of data to return:" & vbCrLf _
    & "      'thk' for thickness" & vbCrLf _
    & "      'id' for inside diameter" & vbCrLf _
    & "      'od' for outside diameter)."
End Sub


Public Function pipedata2(ByVal nps As Double, _
                ByVal sch As String, _
                ByVal returnType As String _
                ) As Double

    'returns pipe data given size and schedule
    'returnType - 'thk' returns thickness
    'returnType - 'id' returns id in inches
    'returnType - 'od' returns od in inches
    'errors:    -1 schedule not found
    '           -2 nps not found
    '           -3 invalid return type specified
    
    Dim i As Integer
    Dim j As Integer
    Dim pip_ary() As Double
    Dim sch_ary() As String
    Dim sch_dic As Object   'dictionary object to hold the pipe schedules
    Dim nps_dic As Object   'dictionary object to hold the nominal pipe sizes
    Dim err_code As Integer
    Dim thk_override As Double
    
    sch = LCase(sch)    'convert schedule to lcase
    returnType = LCase(returnType)    'convert schedule to lcase
    thk_override = 0    'initialize to zero
    
    'read the data table for pipe sizes and schedules
    Call get_pipe_ary(pip_ary, sch_ary)
    
    'setup a dictionary for nps as keys and counting index
    Set nps_dic = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(pip_ary, 1) + 1
        nps_dic(pip_ary(i - 1, 0)) = i
    Next i
    
    'print dictionary for ref to immediate window
    'Call print_dict(nps_dic)  (only used for troubleshooting)
    
    'setup a dictionary for schedules with schedule as keys and index
    Set sch_dic = CreateObject("Scripting.Dictionary")
    For j = 1 To UBound(sch_ary) + 1
        sch_dic(sch_ary(j - 1)) = j
    Next j
    
    'print dictionary for ref to immediate window
    'Call print_dict(sch_dic)  (only used for troubleshooting)
    
    'get the index numbers of both dictionary's if the keys exist
    If nps_dic.Exists(nps) Then
        i = nps_dic(nps) - 1 'subtract one because disctionary index is 1 to ... wherase array is base 0
    Else
        err_code = -2
        GoTo err:
    End If
    
    'get the index numbers of both dictionary's if the keys exist
    If sch_dic.Exists(sch) Then
        j = sch_dic(sch) - 1    'subtract one because disctionary index is 1 to ... wherase array is base 0
    Else
        'now check to see if a specified thickness was passed where we will check to see if the number is between 0 and 4 and then assume
        'that this is a custom thickness passed
        If IsNumeric(sch) Then
            If CDbl(sch) > 0 And CDbl(sch) < 4 Then
                thk_override = CDbl(sch)
            End If
        Else
            If returnType <> "od" Then
                err_code = -1
                GoTo err:
            End If
        End If
    End If
    
    'if both keys are retieved now get the pipe data requested
    Select Case returnType
        Case "od"
            pipedata2 = pip_ary(i, 1)
        Case "thk"
            If thk_override > 0 Then
                pipedata2 = thk_override
            Else
                pipedata2 = pip_ary(i, j)
            End If
        Case "id"
            If thk_override > 0 Then
                pipedata2 = pip_ary(i, 1) - 2 * thk_override
            Else
                pipedata2 = pip_ary(i, 1) - 2 * pip_ary(i, j)
            End If
        Case Else
            err_code = -3
            GoTo err:
    End Select

    'cleanup
    Erase pip_ary
    Erase sch_ary
    
    'if errors are not encountered then exit the function bypass error handler below
    Exit Function

err:
    pipedata2 = err_code
    'see error codes at top
End Function


Public Function buttweld_wt(nomSize As Double, sched As String, optional_thk As String)

    Dim root_gap As Double
    Dim root_face As Double
    Dim thk As Double
    Dim half_bevel_angle As Double
    Dim leg2a As Double
    Dim leg2b As Double
    Dim PI As Double
    
    Dim area1 As Double
    Dim area2 As Double
    Dim area3 As Double
    
    Dim pipe_rad As Double
    Dim rad1 As Double
    Dim rad2 As Double
    Dim rad3 As Double
    
    Dim vol1 As Double
    Dim vol2 As Double
    Dim vol3 As Double
    
    Dim sagitta_len As Double   'this is the math term for distance from chord line to outer radius,  height of weld
    Dim chord_len As Double
    Dim chord_radius As Double
    Dim chord_ang_radians As Double
    Dim dens_steel As Double    'density of steel in lbs per cubic inch
    
    PI = 3.14159265358979
    
    root_face = 0.0625   'reference weldbend catalog page 107
    
    If nomSize <= 6 Then
        root_gap = 1 / 8   'Root Gap Tolerance is an non-essential variable and shall; comply with the WPS. Normal Root Gap used is 2.4mm +/- 0.8mm and it depends on factors like Thickness, size of Filler wire used for root pass, welder's skill and no restriction as per ASME B 31.3.
    Else
        root_gap = 5 / 32
    End If
    
    half_bevel_angle = 75 / 2
    pipe_rad = pipeData(nomSize, "", "od") / 2
    thk = pipeData(nomSize, sched, "thk")

    'compute areas and volumes
    '1. Area of rectangle from root gap thru the thickness of the pipe
    area1 = root_gap * thk
    rad1 = pipe_rad - thk / 2
    vol1 = 2 * PI * rad1 * area1
    '2. Area of both triangles from root face od of pipe at the bevel angle
    
    leg2a = thk - root_face
    leg2b = leg2a * Tan((half_bevel_angle) * PI / 180)
    area2 = leg2a * leg2b
    rad2 = pipe_rad - leg2a / 3
    vol2 = 2 * PI * rad2 * area2
    
    '3. Compute the area of the outside weld cap by assuming it is equal to the area of a circular chord segment
    If nomSize <= 8 Then
        sagitta_len = 0.0625
    Else
        sagitta_len = 0.125
    End If
    
    'calculate the estimated length of weld cap
    chord_len = 2 * leg2b + root_gap + 2 * sagitta_len
    chord_radius = chord_len ^ 2 / (8 * sagitta_len) + sagitta_len / 2
    chord_ang_radians = 2 * Math.Arcsin(chord_len / (2 * chord_radius))
    area3 = 0.5 * chord_radius ^ 2 * (chord_ang_radians - Sin(chord_ang_radians))
    rad3 = pipe_rad + sagitta_len / 2    'this is the mid point which is slightly conservative
    vol3 = 2 * PI * rad3 * area3
    
    'calculate the weight of the weld area
    dens_steel = 0.2836
    buttweld_wt = (vol1 + vol2 + vol3) * dens_steel ' density of steel
End Function


Public Function getSize1(strg)
    Dim inchLoc1 As Integer
    inchLoc1 = InStr(1, strg, """", vbTextCompare)
    getSize1 = convFtInToDecIn(Left(strg, inchLoc1))
End Function

Public Function getSize2(strg)
    Dim inchLoc1, inchLoc2, locX, LenLoc As Integer
    Dim tmpSize2 As String
    
    inchLoc1 = InStr(1, strg, """", vbTextCompare)
    inchLoc2 = InStr(inchLoc1 + 1, strg, """", vbTextCompare)
    
    'Make Sure that Size 2 is not actually a length
        LenLoc = InStr(inchLoc1 + 1, LCase(strg), """ long", vbTextCompare)
    
        If LenLoc = 0 Then
            LenLoc = InStr(inchLoc1 + 1, LCase(strg), """ lg", vbTextCompare)
        End If
    
    If inchLoc2 = LenLoc Then
        inchLoc2 = 0
    End If
    
    
    If inchLoc2 = 0 Then
        getSize2 = ""
    Else
        tmpSize2 = Mid(strg, inchLoc1, inchLoc2 - inchLoc1 + 1)
        locX = InStr(1, LCase(tmpSize2), "x", vbTextCompare)
        tmpSize2 = Right(tmpSize2, Len(tmpSize2) - locX)
        getSize2 = convFtInToDecIn(tmpSize2)
    End If
End Function


Public Function get_sch_1(ByVal strg As String) As String
    Dim locX As Integer
    locX = InStr(1, strg, "x", vbBinaryCompare)
    If locX > 0 Then
        get_sch_1 = Left(strg, locX - 2)
    Else
        get_sch_1 = strg
    End If
End Function


Public Function get_sch_2(ByVal strg As String) As String
    Dim locX As Integer
    locX = InStr(1, strg, "x", vbBinaryCompare)
    
    If locX > 0 Then
        get_sch_2 = Right(strg, Len(strg) - locX - 1)
    Else
        get_sch_2 = ""
    End If
End Function


Public Function stud_nut_wt(ByVal Dia As Double, ByVal length As Double) As Double
    
    Dim nutDict As Object
    Set nutDict = CreateObject("Scripting.Dictionary")
    
    ' Data
    Dim DiaAry As Variant
    Dim WtAry As Variant
    
    DiaAry = Array(0.5, 0.625, 0.75, 0.875, 1#, 1.125, 1.25, 1.375, 1.5, 1.625, 1.75, 1.875, 2, 2.25, 2.5, 2.75, 3, 3.25, 3.5, 3.75)
    WtAry = Array(0.07, 0.12, 0.2, 0.3, 0.43, 0.59, 0.79, 1.02, 1.31, 1.62, 2.04, 2.41, 2.99, 4.19, 5.64, 7.38, 9.5, 11.94, 15.26, 18.12)
    
    ' Populate Dictionary
    For i = LBound(DiaAry) To UBound(DiaAry)
        nutDict(DiaAry(i)) = WtAry(i)
    Next i
    
    Dim PI As Double
    PI = 3.14159
    Dim dens_stl As Double
    dens_stl = 0.2836
    
    stud_nut_wt = dens_stl * PI / 4 * Dia ^ 2 * length + 2 * nutDict(Dia)
End Function