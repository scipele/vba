'   filename:  workday_custom_function.vba
'
'   Purpose:  Computing the number of workdays between given dates, and inputing the work schedule
'
'   Format of WkDayPerWk Parmeter
'       Param       MTWTFSS         - indicate work 1 or nonwork 0
'       "4"         1111000         - Indicates working 4 Days/Wk     M,T,W,Th
'       "5"         1111100         - Indicates working 5 Days/Wk     M,T,W,Th,F
'       "6"         1111100         - Indicates working 6 Days/Wk     M,T,W,Th,F
'       "7"         1111100         - Indicates working 7 Days/Wk     M,T,W,Th,F,S
'       "13-1"      11111111111110  - Indicates working 13 on, 1 off
'
'   Dependencies:  None
'
'   by T. Sciple, scipele@yahoo.com,  8/7/2024

Function work_days(dateA As Date, _
                  dateB As Date, _
                  WkDayPerWk As String)
    
    'make sure date B is greater than date a
    If dateA > dateB Then
        WorkDays = "Error"
        Exit Function
    End If
    
    'determine which day of the week for each date input
    Dim dayOfWk1 As String
    Dim dayOfWk2 As String
    dayOfWk1 = Weekday(dateA, vbMonday) - 1     'less one to align with array starting at 0
    dayOfWk2 = Weekday(dateB, vbMonday) - 1     'less one to align with array starting at 0
    
    'calculated the delta number of weeks and clendar days between the dates
    Dim weeks As Long
    Dim calDays As Long
    weeks = DateDiff("W", dateA, dateB, vbMonday)
    calDays = DateDiff("d", dateA, DateAdd("d", 1, dateB))
    
    'set the week array depending on the number of days worked per week
    Dim wkAry As Variant
    ReDim wkAry(7)
    Select Case WkDayPerWk
        Case "4"
            wkAry = Array(1, 1, 1, 1, 0, 0, 0)
        Case "5"
            wkAry = Array(1, 1, 1, 1, 1, 0, 0)
        Case "6"
            wkAry = Array(1, 1, 1, 1, 1, 1, 0)
        Case "7"
            wkAry = Array(1, 1, 1, 1, 1, 1, 1)
        Case "13-1"
            ReDim wkAry(14)
            wkAry = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0)
    End Select
    
    'loop thru the weeks and create a running sum of the number of workdays in each week
    Dim i As Long
    For i = LBound(wkAry, 1) To UBound(wkAry, 1)
        Sum = Sum + wkAry(i)
    Next i

    'compute the avg work days per week
    Dim avgDayWk As Double
    avgDayWk = 7 * Sum / (UBound(wkAry, 1) - LBound(wkAry, 1) + 1)
    
    'count forward to day 7
    If weeks Mod 2 = 0 Then
        oddEven = "Even"
    Else
        oddEven = "Odd"
    End If
    
    'hande the case with 13-1 work schedule
    Dim adjust As Integer
    adjust = 0
    If (WkDayPerWk = "13-1") And oddEven = "Odd" Then
        adjust = 7
    End If
    
    'count work days
    Dim cnt As Integer
    If (dayOfWk2 - dayOfWk1) >= 0 Then
        For i = dayOfWk1 + adjust To dayOfWk2 + adjust
            If wkAry(i) = 1 Then cnt = cnt + 1
        Next i
    Else
        For i = dayOfWk1 + adjust To 6 + adjust
            If wkAry(i) = 1 Then cnt = cnt + 1
        Next i
        For i = dayOfWk2 To 0 Step -1
            If wkAry(i) = 1 Then cnt = cnt + 1
        Next i
    End If
    
    'note that the int function rounds up for 13-1 schedule if its an odd number of weeks
    work_days = Int(-1 * weeks * avgDayWk) / -1 + cnt
    
End Function
