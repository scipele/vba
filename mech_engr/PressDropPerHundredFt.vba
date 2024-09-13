Public Function PressDropPerHundredFt(nomDia, sch, OptionalThk, pph, viscCP, densPPCF)
    Dim id As Double
    id = pipeID(nomDia, sch, OptionalThk)
    Dim nRe As Double
    nRe = 6.31 * pph / (viscCP * id)
    Dim roughFactor As Double
    roughFactor = 1.5 * 10 ^ (-4)
    Dim diaRatio As Double
    diaRatio = roughFactor / (id / 12)
    Dim ffM As Double
    If nRe < 2000 Then
        ffM = Application.WorksheetFunction.Max(64 / nRe, 0.04)
    Else
        ffM = (-2 * Log10(diaRatio / 3.7 - 5.02 / nRe * Log10(diaRatio / 3.7 - 5.02 / nRe * Log10(diaRatio / 3.7 + 13 / nRe)))) ^ (-2)
    End If
    PressDropPerHundredFt = 0.000336 * ffM * pph ^ 2 / id ^ 5 / densPPCF
End Function