Private Function ary_2d_to_1d(ByRef ary As Variant)
    Dim i As Long
    Dim tmp_ary As Variant
    ReDim tmp_ary(LBound(ary) To UBound(ary))
    
    For i = LBound(ary) To UBound(ary)
        tmp_ary(i) = ary(i, 1)
    Next i
    'reset the ary to the contents of the temporary ary
    ary_2d_to_1d = tmp_ary
    
    'cleanup
    Erase tmp_ary
End Function
