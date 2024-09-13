Function LogN(ByVal inp As Double) _
              As Double
    
    Dim e As Double
    e = 2.718281828459
    LogN = Log(inp) / Log(e)
End Function


Public Function Log10(x) As Double
    Log10 = Log(x) / Log(10#)
End Function


Function Arcsin(x As Double) As Double
    Arcsin = Atn(x / Sqr(-x * x + 1))
End Function
