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


Function Inpolate(ByVal x1 As Double, _
                  ByVal x2 As Double, _
                  ByVal x3 As Double, _
                  ByVal y1 As Double, _
                  ByVal y3 As Double) _
                  As Double

    'solves for y2**, interpolation in the form
    '
    '   +----+-- x1      y1   --+------+
    '   |    |                  |      |
    '   |    +-- x2      y2** --+      |
    '   |                              |
    '   +------- x3      y3 -----------+
    '
    Inpolate = y1 - ((x1 - x2) / (x1 - x3) * (y1 - y3))
End Function