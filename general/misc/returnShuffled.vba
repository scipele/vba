Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | return_shuffled_fisher_yates.vba                            |
'| EntryPoint   | main                                                        |
'| Purpose      | generate a random order of items                            |
'| Inputs       | max_indx                                                    |
'| Outputs      | shuffled array                                              |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 2/17/2025                                         |


Public Sub main()

    Dim max_indx As Long
    max_indx = ThisWorkbook.Sheets("shuffle").Range("max_indx_rng").Value

    Dim ary As Variant
    ary = getShuffledAry(max_indx)
    
    ' Print shuffled array
    Dim i As Long
    For i = 0 To max_indx
        ThisWorkbook.Sheets("shuffle").Range("a" & i + 4).Value = i
        ThisWorkbook.Sheets("shuffle").Range("B" & i + 4).Value = ary(i)
    Next i

End Sub


Public Function getShuffledAry(ByVal max_indx As Long) As Variant

    ' Generate a sequential array from 0 to max_indx
    Dim tmp_ary As Variant
    ReDim tmp_ary(0 To max_indx)
    Dim i As Long

    For i = 0 To max_indx
        tmp_ary(i) = i
    Next i

    ' Fisher-Yates Shuffle Algorithm
    Randomize
    Dim j As Long
    Dim temp As Long
    
    ' Loop backward from Last element to second item
    For i = max_indx To 1 Step -1    ' dont need to run for last element
        j = Int(i * Rnd)   ' Generate random index from 1 to i
        
        ' Swap the random elements
        temp = tmp_ary(i)
        tmp_ary(i) = tmp_ary(j)
        tmp_ary(j) = temp
    Next i

    getShuffledAry = tmp_ary
End Function
