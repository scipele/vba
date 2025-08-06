Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | PaintSfFactors.vba                                          |
'| EntryPoint   | Public Function                                             |
'| Purpose      | Compute Square Footage of Pipe/Fittings for Paint Estimating|
'| Inputs       | nom_dia, component_types (see below)                        |
'| Outputs      | paint sf factor                                             |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 8/5/2025                                          |

' Defined constants at module level so they are available for subsequent calls from dbase or spreadsheet
Public Const PI_D As Double = 3.14159265358979
' End of Globals


Public Function GetPaintSfFactor(ByVal nps As Double, _
                                 ByVal component_desc As String, _
                                 ByVal l_rtg As String) _
                                 As Variant
                                 
    'set the nominal pipe size the same as 3 inch for anything smaller than 3 inch
    If nps < 3 Then nps = 3
    
    'Valid Codes for component_desc parameters are as as follows.
        'cap
        'elb_90
        'flg  -  note that the rating is appended in a switch statement for flanges
        'nipple
        'olet
        'pip
        'reducer
        'stub_end
        'tee
        'vlv_flg
        'vlv_wld

    Dim pipe_sf As Double
    
    On Error Resume Next
    
    Dim nps_ary As Variant
    Dim nps_indx As Integer
    nps_ary = Array(3, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36, 38, 40, 42, 44, 46, 48, 50, 52, 54, 56, 58, 60, 62, 66, 72, 74, 78, 80, 84, 90, 96, 102, 108, 114, 120, 126, 132, 138, 144)
    
    nps_indx = GetNpsIndexNumber(nps, nps_ary)
    
    Dim data As Variant
    
    ' Combineded the rating for flanges
    If component_desc = "flg" Then
        component_desc = Switch(l_rtg = "150", "flg_300", l_rtg = "300", "flg_300", l_rtg = "600", "flg_400_900", l_rtg = "900", "flg_400_900", l_rtg = "1500", "flg_1500_2500", l_rtg = "2500", "flg_1500_2500", True, "flg_300")
    End If

    Select Case component_desc
        'First cases are checked where the sf can be calculated
        Case "pip", "nipple"
            pipe_sf = CalcPipeSf(nps)
        Case "reducer", "cap", "shoe"
            pipe_sf = 2 * CalcPipeSf(nps)
        'Next are the cases that have to be looked up from tabular values
        Case "flg_300"
            data = Array(2, 2.4, 3.6, 4.6, 5.8, 6.8, 7.4, 8.4, 9.6, 10.6, 11.8, 12.9, 13.6, 14.3, 15.8, 17.3, 19, 21.1, 21.6, 22, 22.3, 24.1, 26.8, 28.6, 30.4, 32.4, 34.9, 39.1, 41.8, 44, 46.4, 48.6, 51.9, 54.9, 59.8, 63.6, 68.4, 78.1, 87.9, 98.8, 109.8, 122, 134.1, 147.6, 160.8, 175.6, 190)
        Case "flg_400_900"
            data = Array(2, 2.4, 3.6, 4.8, 6.2, 7.7, 8.5, 10, 12.6, 14.8, 18.7, 22.5, 23.3, 24.1, 24.8, 28.1, 31.7, 34.5, 34.7, 35.2, 37.5, 41.8, 46.22, 48.9, 53.8, 56.9, 60.7, 66.2, 69.5, 76.6, 0, 0, 0, 0, 0, 0, 0, 0)
        Case "flg_1500_2500"
            data = Array(2.4, 3.3, 6.1, 8.2, 12.9, 16.5, 20.2, 24, 29.5, 33.3, 40.1, 46.9, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
        Case "elb_90"
            data = Array(1.5, 1.8, 2.7, 3.6, 5.6, 8, 10.1, 13.2, 17.5, 20.6, 24.9, 30.6, 34.7, 40.3, 46.3, 52.6, 59.4, 66.7, 74.2, 82.2, 90.7, 99.5, 108.8, 118.4, 128.5, 139, 149.9, 161.2, 172.9, 185.1, 197.6, 223.9, 266.5, 281.5, 312.7, 329, 362.7, 416.4, 473.7, 534.8, 599.6, 668, 740.2, 716.1, 895.7, 978.9, 1065.9)
        Case "tee"
            data = Array(3, 3.6, 5.4, 6.9, 8.7, 10.2, 11.1, 12.6, 15.9, 19.6, 23.2, 26.7, 33.2, 38.2, 43.2, 54.6, 60.1, 62.4, 71.8, 78.8, 82.5, 92.8, 99, 106.8, 114.6, 122.4, 130.8, 139.7, 148.2, 157.1, 166.4, 185.7, 216.8, 227.6, 250.1, 261.8, 285.9, 324, 407.2, 407.2, 452.4, 499.9, 549.8, 602, 656.6, 713.5, 772.8)
        Case "stub_end"
            data = Array(1, 1, 1.3, 1.7, 2.6, 3.1, 4, 4.5, 5.1, 5.7, 6.2, 6.8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
        Case "olet"
            data = Array(0.5, 0.8, 1.3, 1.8, 2.3, 3, 3.2, 3.7, 4.2, 4.8, 5.1, 5.7, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
        Case "vlv_wld"
            data = Array(4, 4.8, 7.2, 9.2, 11.6, 13.6, 14.8, 16.8, 19.2, 21.2, 23.2, 25.2, 27.2, 29.2, 31.6, 33.6, 35.6, 38, 39.6, 42, 44, 46.4, 48, 50.4, 52.4, 54.4, 56.8, 58.8, 60.8, 63.2, 64.8, 69.2, 75.6, 77.6, 81.6, 83.6, 88, 94.4, 100.8, 106.8, 113.2, 119.6, 125.6, 132, 138.4, 144.4, 150.8)
        Case "vlv_flg"
            data = Array(6, 7.2, 10.8, 13.8, 17.4, 20.4, 22.2, 25.2, 28.8, 31.8, 34.8, 37.8, 40.8, 43.8, 47.4, 50.4, 53.4, 57, 59.4, 63, 66, 69.6, 72, 75.6, 78.6, 81.6, 85.2, 88.2, 91.2, 94.8, 97.2, 103.8, 113.4, 116.4, 122.4, 125.4, 132, 141.6, 151.2, 160.2, 169.8, 179.4, 188.4, 198, 207.6, 216.6, 226.2)
        Case Else
            pipe_sf = -1    'used if a valid case is not found
    End Select
    
    'Retrieve the Paint Sf Value depending on the index value, unless already calculated
    If pipe_sf = 0 Then
        pipe_sf = IIf(pipe_sf = -1, 1, data(nps_indx))
    End If

    ' Example: Access and use the arrays
    Debug.Print "nps: " & nps, "pipe_sf: " & pipe_sf
    
    GetPaintSfFactor = IIf((pipe_sf - 0) < 0.0001, "", Int(-100 * pipe_sf) / -100)

End Function
           
           
Private Function CalcPipeSf(ByVal nps As Double) _
                            As Double
                            
    Dim od As Double
    od = GetPipeOD(nps)
                            
    CalcPipeSf = IIf(nps < 4, 1, PI_D * od / 12)

End Function


Private Function GetNpsIndexNumber(ByVal nps As Double, _
                                   ByRef nps_ary As Variant) _
                                   As Integer
                                    
    Dim elem As Variant
    Dim i As Integer
    i = 0
    
    For Each elem In nps_ary
        If elem = nps Then
            GetNpsIndexNumber = i
            Exit Function
        End If
        i = i + 1
    Next elem

End Function
           
           
Private Function Get_Polynomial_Equation_Result(ByVal x As Double, _
                                                ByVal a As Double, _
                                                ByVal b As Double, _
                                                ByVal c As Double) _
                                                As Double
                                                
    Get_Polynomial_Equation_Result = a * x ^ 2 + b * x + c
End Function


Public Function GetPipeOD(nps As Double) As Double
    ' Static collection to store NPS-to-OD mappings
    Static PipeData As Collection
    Dim NPSArray As Variant
    Dim ODArray As Variant
    Dim i As Integer
    
    ' Initialize collection if not already done
    If PipeData Is Nothing Then
        Set PipeData = New Collection
        NPSArray = Array(0.125, 0.25, 0.375, 0.5, 0.75, 1, 1.25, 1.5, 2, 2.5, 3, 3.5, 4, 4.5, 5, 6, 7, 8, 9, 10, 11, 12)
        ODArray = Array(0.405, 0.54, 0.675, 0.84, 1.05, 1.315, 1.66, 1.9, 2.375, 2.875, 3.5, 4, 4.5, 5, 5.563, 6.625, 7.625, 8.625, 9.625, 10.75, 11.75, 12.75)
        
        ' Populate collection with NPS as key (string for precise matching) and OD as value
        For i = LBound(NPSArray) To UBound(NPSArray)
            PipeData.Add ODArray(i), CStr(NPSArray(i))
        Next i
    End If
    
    ' Check if NPS is greater than 12 then the OD is the same as the nps
    If nps > 12 Then
        GetPipeOD = nps
        Exit Function
    End If
    
    ' Retrieve OD from collection
    On Error Resume Next
    GetPipeOD = PipeData(CStr(nps))
    If Err.Number <> 0 Then
        GetPipeOD = 0 ' Return 0 if NPS not found
    End If
    On Error GoTo 0
End Function
