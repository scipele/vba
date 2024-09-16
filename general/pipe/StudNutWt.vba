Option Explicit
' filename:         StudNutWt.vba
'
' purpose:          return the weight of studs and nuts for a given stud
'                   size and stud length
'
' usage:            wt = stud_nut_wt(1.25,6.25)
'
' dependencies:     none
'
' By:               T.Sciple, 09/16/2024

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