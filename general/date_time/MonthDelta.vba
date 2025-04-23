Public Function MonthDelta(dateA, dateB)
    Dim moFracA, moFracB As Double
    Dim Yr, curYr, lastYr, moCnt As Long
    Dim curYrMo, lastYrMo As Double
    Dim curMo, lastMo  As Integer
    Dim Choose As Integer
    Dim negFlag As Integer
    Dim tmpDate As Date
    
    negFlag = 1
    'swap dateA and dateB if A > B
    If DateDiff("d", dateA, dateB) < 0 Then
        negFlag = -1
        tmpDate = dateA
        dateA = dateB
        dateB = tmpDate
    End If
    
    'count months between two input dates
    'set first month to start counting from
    curMo = Month(dateA) + 1 'start counting with the month following the start date
    curYr = Year(dateA)
    curYrMo = curYr + curMo / 12
    'set last month to stop counting at
    lastMo = Month(dateB)
    lastYr = Year(dateB)
    lastYrMo = lastYr + lastMo / 12
    
    If curYrMo > lastYrMo Then Choose = 1
    If curYrMo = lastYrMo Then Choose = 2
    
    Select Case Choose
        Case 1
            moCnt = 0
            moFracA = (Day(dateB) - Day(dateA) + 1) / noDaysPerMo(dateA)
            moFracB = 0
        Case 2
            moFracA = (noDaysPerMo(dateA) - Day(dateA) + 1) / noDaysPerMo(dateA)
            moFracB = Day(dateB) / noDaysPerMo(dateB)
            moCnt = 0
        Case Else
            moFracA = (noDaysPerMo(dateA) - Day(dateA) + 1) / noDaysPerMo(dateA)
            moFracB = Day(dateB) / noDaysPerMo(dateB)
            Do
                moCnt = moCnt + 1
                If curMo = 12 Then
                    curYr = curYr + 1
                    curMo = 1
                Else
                    curMo = curMo + 1
                End If
                curYrMo = curYr + curMo / 12
                
            Loop Until curYrMo = lastYrMo
    End Select
    
    MonthDelta = Round(negFlag * (moFracA + moCnt + moFracB), 4)
End Function
