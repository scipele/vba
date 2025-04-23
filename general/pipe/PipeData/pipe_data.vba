Option Explicit
' filename:         PipeData.cls
'
' purpose:          class to return pipe properties
'
' usage:            varies
'
' dependencies:     none
'
' By:               T.Sciple, 09/16/2024

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