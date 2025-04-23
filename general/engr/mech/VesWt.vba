Option Explicit

Public Function VesWt(idFt, ttL, dP, dt, matlcode, ca, jointEff, headQty, headType, nozFact, skirtL, inclSaddleYN)
    'head type "ELLIP, HEMI"
    'idFt - Inside Diameter in Feet
    'ttL - Tangent to Tangent Length in Ft
    'dP - Design Pressure in psig
    'dT - Design Temperature in deg F
    'ca - Corrosion Allowance in Inches
    'nozFact - Fraction for % to Add for Nozzles Misc as a percentage of Shell Weight
    'skirtL - Skirt Length in Feet
    'jointEff - Joint Efficiency
    'headQty - Quantity of Heads
    'matlcode - For Variou Material Types
    'inclSaddleYN - Flag for whether to include the estimated weight of Saddle Supports
    
    Dim tempRngStr As String
    tempRngStr = "100,150,200,250,300,400,500,600,650,700,750,800,850,900,950,1000,1050,1100,1150,1200"
    Dim StressStr As String
    
    Select Case matlcode
        '1-SA51670, 2-SA106B, 3-SA240316High, 4-SA240316Low, 5-SA240316L, 6-SA240304LHigh, 7-SA240304LLow, 8-SA240304Low,
        '9-SA240304High, 10-22Cr5Ni3MoN, 11-SA105, 12-SA-387-22 (2-1/4CR)
        Case "1": StressStr = "20,20,20,20,20,20,20,19.4,18.8,18.1,14.8,12,9.3,6.7,4,2.5,,,,"
        Case "2": StressStr = "17.1,17.1,17.1,17.1,17.1,17.1,17.1,17.1,17.1,15.6,13,10.8,8.7,5.9,4,2.5,,,,"
        Case "3": StressStr = "20,20,20,20,20,19.3,18,17,16.6,16.3,16.1,15.9,15.7,15.6,0,0,,,,"
        Case "4": StressStr = "20,18.65,17.3,16.45,15.6,14.3,13.3,12.6,12.3,12.1,11.9,11.8,11.6,11.5,0,0,,,,"
        Case "5": StressStr = "16.7,16.7,16.7,16.7,16.7,15.7,14.8,14,13.7,13.5,13.2,12.9,12.7,0,0,0,,,,"
        Case "6": StressStr = "16.7,16.7,16.7,16.7,16.7,15.8,14.7,14,13.7,13.5,13.3,13,12.8,11.9,0,0,,,,"
        Case "7": StressStr = "16.7,15.5,14.3,13.55,12.8,11.7,10.9,10.4,10.2,10,9.8,9.7,9.5,9.3,0,0,,,,"
        Case "8": StressStr = "20,18.35,16.7,15.85,15,13.8,12.9,12.3,12,11.7,11.5,11.2,11,10.8,0,0,,,,"
        Case "9": StressStr = "20,20,20,19.45,18.9,18.3,17.5,16.6,16.2,15.8,15.5,15.2,14.9,14.6,0,0,,,,"
        Case "10": StressStr = "25.7,25.7,25.7,25.25,24.8,23.9,23.3,23.1,0,0,0,0,0,0,0,0,,,,"
        Case "11": StressStr = "20,20,20,20,20,20,19.6,18.4,17.8,17.2,14.8,12,9.3,6.7,4,2.5,,,,"
        Case "12": StressStr = "17.1,17.1,17.1,16.85,16.6,16.6,16.6,16.6,16.6,16.6,16.6,16.6,16.6,13.6,10.8,8,5.7,3.8,2.4,1.4"
    End Select
    
    'Calculate Allowable Stress given Design Temperature and material type
    Dim tempRngAry As Variant
    tempRngAry = Split(tempRngStr, ",")
    Dim StressAry As Variant
    StressAry = Split(StressStr, ",")
    Dim i As Long
    i = 0
    Do
        i = i + 1
    Loop Until dt < Val(tempRngAry(i))
    
    Dim lowerTemp As Double
    Dim upperTemp As Double
    Dim lowerStress As Double
    Dim upperStress As Double
    Dim allowStressDT As Double
    lowerTemp = Val(tempRngAry(i - 1))
    upperTemp = Val(tempRngAry(i))
    lowerStress = Val(StressAry(i - 1))
    upperStress = Val(StressAry(i))
    allowStressDT = lowerStress - ((lowerTemp - dt) / (lowerTemp - upperTemp) * (lowerStress - upperStress))
    
    'convert from ksi to psi
    allowStressDT = allowStressDT * 1000
    
    Dim idIn As Double
    idIn = idFt * 12
    'Calculate Shell Thickness
    
    Dim thk As Double
    thk = dP * (idIn / 2 + ca) / (allowStressDT * jointEff - 0.6 * dP)
    Dim thkCorrReqd As Double
    thkCorrReqd = thk + ca
    Dim fvThk As Double
    fvThk = 0.027 * (idIn) ^ 0.56 + ca
    
    Dim thkActRounded As Double
    thkActRounded = WorksheetFunction.RoundUp(WorksheetFunction.Max(thkCorrReqd, fvThk) * 8, 0) / 8
    
    'Calculate Head Thickness per UG-32(d)
    
    Dim headThkReqd As Double
    Select Case UCase(headType)
        Case "ELLIP"
            'headThk = PD / (2SE - .2P)
            headThkReqd = dP * idIn / (2 * allowStressDT * jointEff - 0.2 * dP)
        Case "HEMI"
            'T = PL / (2SE -.2P)
            Dim l As Double
            l = idIn / 2 ' Inside Spherical Radius
            headThkReqd = dP * l / (2 * allowStressDT * jointEff - 0.2 * dP)
    End Select
        
    Dim headRhkCorrReqd As Double
    headRhkCorrReqd = headThkReqd + ca
    
    Dim headThkActRounded As Double
    headThkActRounded = WorksheetFunction.RoundUp(headRhkCorrReqd * 8, 0) / 8
    
    Dim shellWt As Double
    shellWt = 3.14159 * (idIn + thkActRounded) * ttL * 12 * thkActRounded * 0.2836 * (1 + nozFact)
    
    Dim headWt As Double
    headWt = headQty * 3.14159 * (idIn * 1.3) ^ 2 / 4 * headThkActRounded * 0.2836
    
    Dim skirtThkAssumed As Double
    skirtThkAssumed = 0.5
    Dim skirtWtEstim As Double
    skirtWtEstim = 3.14159 * (idIn + skirtThkAssumed) * skirtThkAssumed * skirtL * 12 * 0.2836 * 1.15
    
    Dim saddleWtEstim As Double
    If UCase(inclSaddleYN) = "Y" Then
        saddleWtEstim = shellWt / ttL * idFt * 0.3
    End If
    
    'Multiply by 1.06 to account for mill tolerance
    VesWt = (shellWt + headWt + skirtWtEstim + saddleWtEstim) * 1.06
End Function
