Option Explicit

Private Sub PassVarToFnc()
    Dim FruitType As String
    Dim color As String
    
    FruitType = "Banana"
    color = FruitColor(FruitType)
    MsgBox color
End Sub

Private Function FruitColor(FruitType As String) As String
    Select Case FruitType
        Case "Banana"
            FruitColor = "Yellow"
        Case "Kiwi"
            FruitColor = "Green"
    End Select
End Function