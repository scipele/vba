Option Explicit

Function hash12(s As String) As String
    Dim l As Integer
    Dim l3 As Integer
    Dim s1 As String
    Dim s2 As String
    Dim s3 As String

    l = Len(s)
    l3 = Int(l / 3)
    s1 = Mid(s, 1, l3)
    s2 = Mid(s, l3 + 1, l3)
    s3 = Mid(s, 2 * l3 + 1)
    hash12 = hash4(s1) + hash4(s2) + hash4(s3)
End Function

Function hash4(txt)
    Dim x As Long
    Dim mask As Long
    Dim i As Integer
    Dim j As Integer
    Dim nC As Integer
    Dim crc As Integer
    Dim c As String
    
    crc = &HFFFF
    For nC = 1 To Len(txt)
        j = Asc(Mid(txt, nC))
        crc = crc Xor j
        For j = 1 To 8
            mask = 0
            If crc / 2 <> Int(crc / 2) Then mask = &HA001
            crc = Int(crc / 2) And &H7FFF: crc = crc Xor mask
        Next j
    Next nC
    
    c = Hex$(crc)

    While Len(c) < 4
      c = "0" & c
    Wend
    
    hash4 = c
End Function