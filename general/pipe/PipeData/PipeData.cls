' filename:         PipeData.cls
'
' purpose:          class to return pipe properties
'
' usage:            varies
'
' dependencies:     none
'
' By:               T.Sciple, 09/16/2024

Option Explicit

Private pip_ary() As Double
Private sch_ary() As String
Private sch_dic As Object
Private nps_dic As Object
Const PI As Double = 3.14159265358979


' Initialize the class
Private Sub Class_Initialize()
    Call get_pipe_ary(pip_ary, sch_ary)
    Set sch_dic = CreateObject("Scripting.Dictionary")
    Set nps_dic = CreateObject("Scripting.Dictionary")
    
    Dim i As Integer
    For i = 1 To UBound(pip_ary, 1) + 1
        nps_dic(pip_ary(i - 1, 0)) = i
    Next i
    
    Dim j As Integer
    For j = 1 To UBound(sch_ary) + 1
        sch_dic(sch_ary(j - 1)) = j
    Next j
End Sub


' Methods to retrieve and / or calculate OD, ID, and Thickness, metal area, empty weight, full weight, moment of inertia, metal weight, paint area
Public Function od(size As Double) As Double
    od = GetPipeData(size, "", "od")
End Function


Public Function id(size As Double, schedule As String) As Double
    id = GetPipeData(size, schedule, "id")
End Function


Public Function thk(size As Double, schedule As String) As Double
    thk = GetPipeData(size, schedule, "thk")
End Function


Public Function area_metal_sqin(size As Double, schedule As String) As Double
    area_metal_sqin = PI / 4 * (GetPipeData(size, "", "od") ^ 2 - GetPipeData(size, schedule, "id") ^ 2)
End Function


Public Function wt_empty_lbs_per_ft(size As Double, schedule As String) As Double
    Const plus_wt_tolerance As Double = 1.1
    wt_empty_lbs_per_ft = plus_wt_tolerance * PI / 4 * (GetPipeData(size, "", "od") ^ 2 - GetPipeData(size, schedule, "id") ^ 2) * 12 * 0.2836
End Function


Public Function wt_full_h20_lbs_per_ft(size As Double, schedule As String) As Double
    Dim empty_wt As Double
    Dim water_wt As Double
    Dim id As Double
    Const plus_wt_tolerance As Double = 1.1
    
    id = GetPipeData(size, schedule, "id")
    empty_wt = plus_wt_tolerance * PI / 4 * (GetPipeData(size, "", "od") ^ 2 - id ^ 2) * 12 * 0.2836
    water_wt = PI / 4 * (id / 12) ^ 2 * 1 * 62.4
    wt_full_h20_lbs_per_ft = empty_wt + water_wt
End Function


Public Function moment_of_inertia_in4(size As Double, schedule As String) As Double
    Dim ro As Double ' outside radious
    Dim ri As Double ' inside radious

    ro = GetPipeData(size, "", "od") / 2
    ri = GetPipeData(size, schedule, "id") / 2
    moment_of_inertia_in4 = PI / 4 * (ro ^ 4 - ri ^ 4)
End Function


Public Function section_modulus_in3(size As Double, schedule As String) As Double
    Dim ro As Double ' outside radious
    Dim ri As Double ' inside radious

    ro = GetPipeData(size, "", "od") / 2
    ri = GetPipeData(size, schedule, "id") / 2
    
    section_modulus_in3 = PI / (4 * ro) * (ro ^ 4 - ri ^ 4)
End Function


Public Function paint_area_sf_per_ft(size As Double) As Double
    Dim actual_area As Double
    
    actual_area = PI * GetPipeData(size, "", "od") / 12
    If actual_area < 1 Then
        paint_area_sf_per_ft = 1
    Else
        paint_area_sf_per_ft = PI * GetPipeData(size, "", "od") / 12
    End If
End Function


' Private function to retrieve pipe data
Private Function GetPipeData(size As Double, schedule As String, returnType As String) As Double
    Dim i As Integer
    Dim j As Integer
    Dim thk_override As Double
    Dim err_code As Integer

    returnType = LCase(returnType)
    thk_override = 0
    schedule = LCase(schedule)

    If nps_dic.Exists(size) Then
        i = nps_dic(size) - 1
    Else
        GetPipeData = -2
        Exit Function
    End If

    If sch_dic.Exists(schedule) Then
        j = sch_dic(schedule) - 1
    Else
        If IsNumeric(schedule) Then
            If CDbl(schedule) > 0 And CDbl(schedule) < 4 Then
                thk_override = CDbl(schedule)
            End If
        Else
            If returnType <> "od" Then
                GetPipeData = -1
                Exit Function
            End If
        End If
    End If

    Select Case returnType
        Case "od"
            GetPipeData = pip_ary(i, 1)
        Case "thk"
            If thk_override > 0 Then
                GetPipeData = thk_override
            Else
                GetPipeData = pip_ary(i, j)
            End If
        Case "id"
            If thk_override > 0 Then
                GetPipeData = pip_ary(i, 1) - 2 * thk_override
            Else
                GetPipeData = pip_ary(i, 1) - 2 * pip_ary(i, j)
            End If
        Case Else
            GetPipeData = -3
    End Select
End Function


' Sub to read the data table for pipe sizes and schedules
Private Sub get_pipe_ary(ByRef pip_ary() As Double, ByRef sch_ary() As String)
    Dim i As Integer
    Dim j As Integer
    
    sch_ary = Split("nps,od,5,10,20,30,40,std,60,80,xs,100,120,140,160,xxs", ",")

    Dim str() As String
    ReDim str(0 To 35)
    str(0) = "0.125 0.405 .035  .049  0     0     .068  .068  0     .095  .095  0     0     0     0     0     "
    str(1) = "0.25  0.54  .049  .065  0     0     .088  .088  0     .119  .119  0     0     0     0     0     "
    str(2) = "0.375 0.675 .049  .065  0     0     .091  .091  0     .126  .126  0     0     0     0     0     "
    str(3) = "0.5   0.84  .065  .083  0     0     .109  .109  0     .147  .147  0     0     0     .187  .294  "
    str(4) = "0.75  1.05  .065  .083  0     0     .113  .113  0     .154  .154  0     0     0     .218  .308  "
    str(5) = "1     1.315 .065  .109  0     0     .133  .133  0     .179  .179  0     0     0     .250  .358  "
    str(6) = "1.25  1.66  .065  .109  0     0     .140  .140  0     .191  .191  0     0     0     .250  .382  "
    str(7) = "1.5   1.9   .065  .109  0     0     .145  .145  0     .200  .200  0     0     0     .281  .400  "
    str(8) = "2     2.375 .065  .109  0     0     .154  .154  0     .218  .218  0     0     0     .343  .436  "
    str(9) = "2.5   2.875 .083  .120  0     0     .203  .203  0     .276  .276  0     0     0     .375  .552  "
    str(10) = "3     3.5   .083  .120  0     0     .216  .216  0     .300  .300  0     0     0     .437  .600  "
    str(11) = "3.5   4     .083  .120  0     0     .226  .226  0     .318  .318  0     0     0     0     .636  "
    str(12) = "4     4.5   .083  .120  0     0     .237  .237  .281  .337  .337  0     .437  0     .531  .674  "
    str(13) = "4.5   5     0     0     0     0     0     .247  0     0     .355  0     0     0     0     .710  "
    str(14) = "5     5.563 .109  .134  0     0     .258  .258  0     .375  .375  0     .500  0     .625  .750  "
    str(15) = "6     6.625 .109  .134  0     0     .280  .280  0     .432  .432  0     .562  0     .718  .864  "
    str(16) = "7     7.625 0     0     0     0     0     .301  0     0     .500  0     0     0     0     .875  "
    str(17) = "8     8.625 .109  .148  .250  .277  .322  .322  .406  .500  .500  .593  .718  .812  .906  .875  "
    str(18) = "9     9.625 0     0     0     0     0     .342  0     0     .500  0     0     0     0     0     "
    str(19) = "10    10.75 .134  .165  .250  .307  .365  .365  .500  .593  .500  .718  .843  1.000 1.125 0     "
    str(20) = "11    11.75 0     0     0     0     0     .375  0     0     .500  0     0     0     0     0     "
    str(21) = "12    12.75 .165  .180  .250  .330  .406  .375  .562  .688  .500  .843  1.000 1.125 1.312 0     "
    str(22) = "14    14    0     .250  .312  .375  .437  .375  .593  .750  .500  .937  1.093 1.250 1.406 0     "
    str(23) = "16    16    0     .250  .312  .375  .500  .375  .656  .843  .500  1.031 1.218 1.437 1.593 0     "
    str(24) = "18    18    0     .250  .312  .437  .562  .375  .750  .937  .500  1.156 1.375 1.562 1.781 0     "
    str(25) = "20    20    0     .250  .375  .500  .593  .375  .812  1.031 .500  1.280 1.500 1.750 1.968 0     "
    str(26) = "22    22    0     .250  .375  .500  0     .375  .875  1.125 .500  1.375 1.625 1.875 2.125 0     "
    str(27) = "24    24    0     .250  .375  .562  .688  .375  .968  1.218 .500  1.531 1.812 2.062 2.343 0     "
    str(28) = "26    26    0     .312  .500  0     0     .375  0     0     .500  .750  1.000 1.250 1.500 2.000 "
    str(29) = "28    28    0     .312  .500  .625  0     .375  0     0     .500  .750  1.000 1.250 1.500 2.000 "
    str(30) = "30    30    0     .312  .500  .625  0     .375  0     0     .500  .750  1.000 1.250 1.500 2.000 "
    str(31) = "32    32    0     .312  .500  .625  .688  .375  0     0     .500  .750  1.000 1.250 1.500 2.000 "
    str(32) = "34    34    0     .312  .500  .625  .688  .375  0     0     .500  .750  1.000 1.250 1.500 2.000 "
    str(33) = "36    36    0     .312  .500  .625  .750  .375  0     0     .500  .875  1.000 1.250 1.500 2.000 "
    str(34) = "42    42    0     0     0     0     0     .375  0     0     .500  .750  1.000 1.250 1.500 2.000 "
    str(35) = "48    48    0     0     0     0     0     .375  0     0     .500  .750  1.000 1.250 1.500 2.000 "

    ReDim pip_ary(0 To 35, 0 To 15)
    For i = 0 To 35
        For j = 0 To 15
            pip_ary(i, j) = Val(Mid(str(i), j * 6 + 1, 6))
        Next j
    Next i
End Sub
