Public Function noDaysPerMo(inpDate)
    noDaysPerMo = Day(DateSerial(Year(inpDate), Month(inpDate) + 1, 1) - 1)
End Function