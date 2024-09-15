Public Function VolOctgCY(FlatDimFt, depthFt)
    triangleHt = FlatDimFt / 2
    triangleBase = Tan(Application.WorksheetFunction.Radians(22.5)) * triangleHt
    triangleArea = 0.5 * triangleBase * triangleHt
    OctArea = 8 * 2 * triangleArea
    VolOctgCY = OctArea * depthFt / 27
End Function


Public Function VolCylCY(DiaFt, depthFt)
    CircleArea = Application.WorksheetFunction.PI() / 4 * DiaFt ^ 2
    VolCylCY = CircleArea * depthFt / 27
End Function


Public Function VolRectCY(lFt, wFt, depthFt)
    VolRectCY = lFt * wFt * depthFt / 27
End Function


Public Function surfAreaEllipHeadSF(DiaFt)
    rad = DiaFt / 2
    Depth = rad / 2
    surfAreaEllipHeadSF = 2 * Application.WorksheetFunction.PI() * ((((rad * rad) ^ 1.6 + (rad * Depth) ^ 1.6 + (rad * Depth) ^ 1.6)) / 3) ^ (1 / 1.6)
End Function


Public Function estimPipeMH(LaborType, size, SchRtg As String)
    'Labor Type = FieldBW, Field BWHvy
    SchRtg = Trim(SchRtg)
    Select Case UCase(LaborType)
    
    Case "ERECTSR"
        'schedules 10-60
        If SchRtg = "10" Or SchRtg = "20" Or SchRtg = "30" Or SchRtg = "40" Or SchRtg = "60" Or SchRtg = "STD" Then SchRtg = "10-60"
        If SchRtg = "XS" And Val(size) >= 10 Then SchRtg = "10-60"
        'schedules 80-100
        If SchRtg = "80" Or SchRtg = "100" Then SchRtg = "80-100"
        If SchRtg = "XS" And Val(size) < 10 Then SchRtg = "80-100"
        'schedules 120-160
        If SchRtg = "120" Or SchRtg = "140" Or SchRtg = "160" Or SchRtg = "XXS" Then SchRtg = "120-160"
            
        Select Case SchRtg
            Case "10-60": mhStr = "0.16,0.16,0.16,0.17,0.17,0.18,0.19,0.2,0.21,0.23,0.24,0.25,0.26,0.28,0.34,0.43,0.52,0.64,0.75,0.88,1.03,1.15"
            Case "80-100": mhStr = "0.17,0.17,0.18,0.19,0.2,0.21,0.22,0.24,0.26,0.28,0.3,0.31,0.34,0.38,0.48,0.6,0.73,0.87,1.02,1.17,1.32,1.49"
            Case "120-160": mhStr = "0.18,0.19,0.2,0.21,0.23,0.24,0.27,0.29,0.32,0.35,0.38,0.39,0.43,0.5,0.65,0.82,1,1.19,1.39,1.6,1.81,2.04"
            Case Else:
                estimPipeMH = "Invalid Schedule"
                Exit Function
        End Select
    
        npsStr = "0.25,0.375,0.5,0.75,1,1.25,1.5,2,2.5,3,3.5,4,5,6,8,10,12,14,16,18,20,24"
        npsAry = Split(npsStr, ",")
        mhAry = Split(mhStr, ",")
        i = 0
        Do Until size = Val(npsAry(i))
            i = i + 1
        Loop
        estimPipeMH = mhAry(i)
    
    Case "FIELDBW"
        Select Case UCase(SchRtg)
            Case "STD": mhStr = "0.7,0.8,0.8,1,1.2,1.3,1.4,1.5,1.7,2,2.6,3.1,3.6,4.3,5,5.9,6.3,6.9"
            Case "XS": mhStr = "0.8,0.8,0.9,1,1.3,1.4,1.6,1.8,2.1,2.5,3.3,4,4.7,5.7,6.6,7.7,8.4,10.1"
            Case "20": mhStr = "na,na,na,na,na,na,na,na,na,na,2.6,3.1,3.6,4.3,5,5.9,6.3,6.9"
            Case "30": mhStr = "na,na,na,na,na,na,na,na,na,na,2.6,3.1,3.6,4.3,5,6.8,8.4,na"
            Case "40": mhStr = "0.7,0.8,0.8,1,1.2,1.3,1.4,1.5,1.7,2,2.6,3.1,4.1,5,6.6,8.6,9.4,13.3"
            Case "60": mhStr = "na,na,na,na,na,na,na,na,na,na,3,4,5.2,6.8,8.4,11.2,13.8,20.1"
            Case "80": mhStr = "0.8,0.8,0.9,1,1.3,1.4,1.6,1.8,2.1,2.5,3.3,5.1,6.6,9.6,12.4,16.4,19.5,25.2"
            Case "100": mhStr = "na,na,na,na,na,na,na,na,na,na,4.6,6.8,9.9,13.2,19.5,21.8,26,35.8"
            Case "120": mhStr = "na,na,na,na,na,na,na,2.8,2.9,3.8,6,9.4,12.2,16.2,20.7,25.6,31.9,43.5"
            Case "140": mhStr = "na,na,na,na,na,na,na,na,na,na,7.5,11.4,15.3,19.2,25,29.9,37,49.3"
            Case "160": mhStr = "1,1.1,1.3,1.6,1.8,2.1,na,3,3.8,4.9,8.6,13.1,17.9,22.7,27.7,33.7,40.8,59.3"
            Case Else:
                estimPipeMH = "Invalid Schedule"
                Exit Function
        End Select
    
        npsStr = "1,1.25,1.5,2,2.5,3,3.5,4,5,6,8,10,12,14,16,18,20,24"
        npsAry = Split(npsStr, ",")
        mhAry = Split(mhStr, ",")
        i = 0
        Do Until size = Val(npsAry(i))
            i = i + 1
        Loop
        estimPipeMH = mhAry(i)
    
    
    Case "FIELDBWHVY"
        Select Case UCase(SchRtg)
            
            Case "0.75": mhStr = "2.7,3.3,,,,,,,,,,,"
            Case "1": mhStr = "3.7,4.1,4.7,6.4,8.7,,,,,,,,"
            Case "1.25": mhStr = ",5.7,6.7,8.5,10.1,13.5,,,,,,,"
            Case "1.5": mhStr = ",6.8,8,10.4,13.1,16.2,19.6,23.5,,,,,"
            Case "1.75": mhStr = ",,10,13.3,16.5,20.1,23.2,26.6,29.9,,,,"
            Case "2": mhStr = ",,12.4,15.6,19.2,23.2,27.4,31.2,35.6,39.8,46.4,,"
            Case "2.25": mhStr = ",,,18.2,22.7,27.3,32.6,36.5,41.5,46.4,54.8,,"
            Case "2.5": mhStr = ",,,,27.4,32.1,37.5,43.1,49.7,54.8,66.2,72.3,78.7"
            Case "2.75": mhStr = ",,,,,36.7,42.8,48.9,56.3,62.9,75.4,82.4,89.4"
            Case "3": mhStr = ",,,,,42.1,49.1,55.5,64.7,72.9,84.5,92.7,99"
            Case "3.25": mhStr = ",,,,,,55.3,62.9,72.9,82.8,96.9,105.5,114.3"
            Case "3.5": mhStr = ",,,,,,63.1,71.3,82.8,95,109.3,119.3,129.2"
            Case "3.75": mhStr = ",,,,,,,81.2,94.4,108.4,124.1,135.6,147.4"
            Case "4": mhStr = ",,,,,,,91.1,107.6,124.1,140.8,154.1,167.4"
            Case "4.25": mhStr = ",,,,,,,,,,159.7,174.8,189.7"
            Case "4.5": mhStr = ",,,,,,,,,,173.1,193.4,209.3"
            Case "4.75": mhStr = ",,,,,,,,,,189.7,204.8,223.5"
            Case "5": mhStr = ",,,,,,,,,,203.3,219.8,240.6"
            Case "5.25": mhStr = ",,,,,,,,,,216.9,240.6,262.9"
            Case "5.5": mhStr = ",,,,,,,,,,225.9,252.6,275.5"
            Case "5.75": mhStr = ",,,,,,,,,,251.5,276.8,298.2"
            Case "6": mhStr = ",,,,,,,,,,268.1,295.1,319"
            Case Else:
                estimPipeMH = "Invalid Schedule"
                Exit Function
        End Select
            
        npsStr = "3,4,5,6,8,10,12,14,16,18,20,22,24"
        npsAry = Split(npsStr, ",")
        mhAry = Split(mhStr, ",")
        i = 0
        Do Until size = Val(npsAry(i))
            i = i + 1
        Loop
        estimPipeMH = mhAry(i)
    Case "FIELDBWLGOD"
        Select Case UCase(SchRtg)
            Case "0.375": mhStr = "8.4,10,12.5,15.5,19.4,23,27,31.6,36.9,42.8,48.3,54.5,61.4,69"
            Case "0.5": mhStr = "11.4,13.1,15.2,17.9,21.5,24.7,28.9,34.2,40.4,46.6,53.1,59.9,67.6,76.2"
            Case "0.75": mhStr = "15.1,16.4,18.9,21.5,24.4,27.8,32,36.8,42.5,49.9,58.3,68.1,79.5,92.9"
            Case "1": mhStr = "20.2,22.2,24.1,26.7,29.5,33.2,37.1,41.5,46.6,57,67.9,79.1,92.2,107.4"
            Case "1.25": mhStr = "26.7,29.3,31.7,34.9,39.3,45.2,52,59.7,68.8,74.9,82.7,90.9,99.7,109.5"
            Case "1.5": mhStr = "34.5,37.3,39.8,43,46.1,52,58.8,66.3,75,83.2,91.5,99.9,109.3,119.4"
            Case "1.75": mhStr = "43.4,46.4,49.6,52.7,56.3,62.3,68.6,75.4,82.8,90.3,98.3,106.8,116,126"
            Case "2": mhStr = "52.5,55.7,58.9,62.1,65.4,71.7,78.1,85.2,92.9,101.2,109.6,118.1,127.3,137.2"
            Case "2.75": mhStr = "61.7,66,68.3,71.4,76.9,83,,,,,,,,"
            Case "2.5": mhStr = "85,91,99.5,104,110,117.5,,,,,,,,"
            Case "2.75": mhStr = "96.3,104.4,112.9,118.9,126.5,134.9,,,,,,,,"
            Case "3": mhStr = "110,117.5,126.5,132.9,142.2,150.7,,,,,,,,"
            Case "3.25": mhStr = "123.5,138.2,144.2,153,161.4,171.7,,,,,,,,"
            Case "3.5": mhStr = "138.5,150.3,161.1,170.2,180.4,192.2,,,,,,,,"
            Case "3.75": mhStr = "159.7,174.8,185.3,196.9,209.3,222.5,,,,,,,,"
            Case "4": mhStr = "180.4,195.2,209.3,222.9,237.9,250,,,,,,,,"
            Case "4.25": mhStr = "203.3,219.8,234.9,250,268.1,282.5,,,,,,,,"
            Case "4.5": mhStr = "225.9,244.9,258.4,277.1,298.2,313.2,,,,,,,,"
            Case "4.75": mhStr = "243.9,261.4,281.5,298.2,319.3,337.4,,,,,,,,"
            Case "5": mhStr = "261.7,280.1,301.1,319.3,340.9,360.8,,,,,,,,"
            Case "5.25": mhStr = "276.8,299.6,322.3,343.4,366.9,387,,,,,,,,"
            Case "5.5": mhStr = "298.2,322,343.4,365.9,391.5,412.6,,,,,,,,"
            Case "5.75": mhStr = "323.7,345.5,371.1,394.6,421.6,445.7,,,,,,,,"
            Case "6": mhStr = "345.7,367.5,400,424.7,451.7,478.3,,,,,,,,"
    
            Case Else:
                estimPipeMH = "Invalid Schedule"
                Exit Function
        End Select
        
        npsStr = "26,28,30,32,34,36,38,40,42,44,46,48,54,60"
        npsAry = Split(npsStr, ",")
        mhAry = Split(mhStr, ",")
        i = 0
        Do Until size = Val(npsAry(i))
            i = i + 1
        Loop
        estimPipeMH = mhAry(i)
    
    End Select
End Function


Function CreateUniqueIDIncrement(curDesc, prevDesc, nextDesc, prevIncrement)
    'Note that using this stupit method the data must be sorted
    'Note that you can use the Text join to concatenate the uniqueID "True" which ignores when there is no value
    'using TEXTJOIN("-",TRUE,G12,B12)
    
    'Define Cases
    If curDesc = nextDesc And curDesc <> prevDesc Then caseDesc = "NewSeries"
    If curDesc = nextDesc And curDesc = prevDesc Then caseDesc = "ContdSeries"
    If curDesc <> nextDesc And curDesc = prevDesc Then caseDesc = "LastSeries"
    
    Select Case caseDesc
        Case "NewSeries"
            CreateUniqueIDIncrement = 1
        Case "ContdSeries"
            CreateUniqueIDIncrement = prevIncrement + 1
        Case "LastSeries"
            CreateUniqueIDIncrement = prevIncrement + 1
        Case Else
            CreateUniqueIDIncrement = ""
    End Select
End Function