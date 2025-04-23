Function CheckInList(lookupText, lookupRng As Range) As Boolean
    If Application.WorksheetFunction.CountIf(lookupRng, lookupText) = 0 Then
        CheckInList = False
    Else
        CheckInList = True
    End If
End Function