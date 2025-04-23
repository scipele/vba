Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | modGetScaledDataPeriodAmt.vba                               |
'| EntryPoint   | Function ComputeScaledDataPeriodAmt                         |
'| Purpose      | Compute Incremental Scaled Data Item as shown below         |
'| Inputs       | origAmts, scaledPeriodCount, scaledPeriod                   |
'| Outputs      | Period Amt                                                  |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 1/2/2025                                          |
'
' Example usage in spreadsheet in Table2 calc_distrib body cells
'
'   =ComputeScaledDataPeriodAmt(Table1[dataRng], _
'                               scaledPeriodCount, _
'                               [@[scaled_period]])
'
' Table1:                         Table2:
'+-------------+---------+       +---------------+--------------+
'| orig_period | dataRng |       | scaled_period | calc_distrib |
'+-------------+---------+       +---------------+--------------+
'|      1      |   7.5   |       |       1       |      6       |
'|      2      |  17.5   |       |       2       |     12       |
'|      3      |   20    |       |       3       |     15       |
'|      4      |   15    |       |       4       |     15       |
'|      5      |   45    |       |       5       |     12       |
'|      6      |  32.5   |       |       6       |     36       |
'|      7      |  24.25  |       |       7       |    28.5      |
'|      8      |  20.25  |       |       8       |    22.7      |
'|             |         |       |       9       |    18.6      |
'|             |         |       |      10       |    16.2      |
'+-------------+---------+       +---------------+--------------+

Enum InterpolationType
    itAllWithin = 0
    itBtmInterpolation = 1
    itTopInterpolation = 2
    itMidInterpolation = 3
End Enum


Public Function GetScaledDataPeriodAmt(ByRef origAmts As Range, _
                                       ByVal scaledPeriodCount As Long, _
                                       ByVal scaledPeriod As Long) As Double
    
    ' Load the range into an array
    Dim origDataAry As Variant
    origDataAry = origAmts.Value

    ' Find the overlapped region using a calculation rather than iterating
    Dim s0 As Double, s1 As Double, orig_period_count As Long
    s0 = (scaledPeriod - 1) / scaledPeriodCount
    s1 = (scaledPeriod) / scaledPeriodCount
    orig_period_count = UBound(origDataAry) - LBound(origDataAry) + 1
    
    ' Save lower and upper bound for looping below
    Dim lower_bnd_orig As Long, upper_bnd_orig As Long
    lower_bnd_orig = Int(s0 * orig_period_count) + 1
    upper_bnd_orig = Int(s1 * orig_period_count) + 1
    If upper_bnd_orig > orig_period_count Then
        upper_bnd_orig = orig_period_count
    End If
    
    ' Variables used in the loop below
    Dim indx As Long
    Dim x0 As Double 'proportionate fraction at start(x0) period
    Dim x1 As Double 'proportionate fraction at end(x1) period
    Dim frac As Double
    Dim sum_amt As Double
    sum_amt = 0
    
    ' now loop from the lower to upper bound computed above
    For indx = lower_bnd_orig To upper_bnd_orig
        x0 = (indx - 1) / orig_period_count
        x1 = (indx) / orig_period_count
        frac = GetInterpolationFrac(x0, x1, s0, s1)
        sum_amt = sum_amt + frac * origDataAry(indx, 1)
    Next indx
    'cleanup
    Erase origDataAry

    GetScaledDataPeriodAmt = sum_amt
    
End Function


Private Function GetInterpolationFrac(x0 As Double, _
                                      x1 As Double, _
                                      s0 As Double, _
                                      s1 As Double) As Double
    Dim interp_type As Integer
    interp_type = Switch( _
        s0 <= x0 And s1 >= x1, itAllWithin, _
        s0 >= x0 And s1 >= x1, itBtmInterpolation, _
        s0 <= x0 And s1 > x0, itTopInterpolation, _
        s0 > x0 And s1 < x1, itMidInterpolation, _
        True, -1)

    ' Use Select Case with formulas for different interpolation types
    Dim result As Double
    Select Case interp_type
        Case itAllWithin
            result = 1
        Case itBtmInterpolation
            result = (x1 - s0) / (x1 - x0)
        Case itTopInterpolation
            result = (s1 - x0) / (x1 - x0)
        Case itMidInterpolation
            result = (s1 + s0) / (x1 - x0)
        Case Else
            result = 0
    End Select
    GetInterpolationFrac = result
End Function