Function WorkDaysEndDate(dateA, WkDays, WkDayPerWk)
    Dim tmpDate As Date
    Dim avgWkDays As Integer
    Dim delta, calcDays As Double

    If WkDayPerWk = "13-1" Then
        avgWkDays = 6.5
    Else
        avgWkDays = WkDayPerWk
    End If
    
    'find approximate end date
    tmpDate = DateAdd("d", ((WkDays - 1) / WkDayPerWk * 7) - 1, dateA)
    
    'check if end date is correct and if not adjust until it matches
    Do
        calcDays = WorkDays(dateA, tmpDate, WkDayPerWk)
        delta = WkDays - calcDays
        If delta = 0 Then
            Exit Do
        Else
            tmpDate = DateAdd("d", delta, tmpDate)
        End If
    Loop Until DateDiff("d", dateA, tmpDate) = WkDays

    WorkDaysEndDate = tmpDate
End Function