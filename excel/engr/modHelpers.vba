'==============================================================================
' Filename:    modHelpers.vba
' Purpose:     Utility functions: table lookups, conversions, validation,
'              multi-tank cascade deposit/withdraw
' Dependencies: modTypes
' By:          T. Sciple, 03/17/2026
'==============================================================================

'##############################################################################
'  MODULE: modHelpers  (Standard Module)
'  Utility functions: table lookups, conversions, validation
'##############################################################################
Option Explicit


Public Function GetConfigValue(ByVal paramName As String) As Variant
'------------------------------------------------------------------------------
' Reads a value from tblRunConfig by ParamName.
' Returns Empty if not found.
'------------------------------------------------------------------------------
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets("Config").ListObjects("tblRunConfig")

    Dim r As ListRow
    For Each r In tbl.ListRows
        If CStr(r.Range(1, 1).Value) = paramName Then
            GetConfigValue = r.Range(1, 2).Value
            Exit Function
        End If
    Next r

    GetConfigValue = Empty
End Function


Public Function ModeNameToEnum(ByVal modeName As String) As transportMode
    Select Case LCase$(Trim$(modeName))
        Case "rail":    ModeNameToEnum = tmRail
        Case "truck":   ModeNameToEnum = tmTruck
        Case "barge":   ModeNameToEnum = tmBarge
        Case Else:      ModeNameToEnum = tmRail  ' default
    End Select
End Function


Public Function IsWeekend(ByVal dt As Date) As Boolean
    Dim wd As Long
    wd = Weekday(dt, vbMonday)  ' Mon=1 .. Sun=7
    IsWeekend = (wd >= 6)       ' Sat=6, Sun=7
End Function


Public Function FindRawTankIndex(ByRef state As SimState, _
                                  ByVal tankName As String) As Long
'------------------------------------------------------------------------------
' Returns 0-based index of raw tank matching tankName, or -1 if not found.
'------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To state.num_raw_tanks - 1
        If state.raw_tanks(i).tank_name = tankName Then
            FindRawTankIndex = i
            Exit Function
        End If
    Next i
    FindRawTankIndex = -1
End Function


Public Function FindBlendTankIndex(ByRef state As SimState, _
                                    ByVal tankName As String) As Long
    Dim i As Long
    For i = 0 To state.num_blend_tanks - 1
        If state.blend_tanks(i).tank_name = tankName Then
            FindBlendTankIndex = i
            Exit Function
        End If
    Next i
    FindBlendTankIndex = -1
End Function


Public Function FindProductTankByProduct(ByRef state As SimState, _
                                          ByVal productName As String) As Long
    Dim i As Long
    For i = 0 To state.num_product_tanks - 1
        If state.product_tanks(i).product_name = productName Then
            FindProductTankByProduct = i
            Exit Function
        End If
    Next i
    FindProductTankByProduct = -1
End Function


Public Function FindRawTankByMaterial(ByRef state As SimState, _
                                       ByVal materialName As String) As Long
    Dim i As Long
    For i = 0 To state.num_raw_tanks - 1
        If state.raw_tanks(i).material_name = materialName Then
            FindRawTankByMaterial = i
            Exit Function
        End If
    Next i
    FindRawTankByMaterial = -1
End Function


Public Function FindAllRawTanksByMaterial(ByRef state As SimState, _
                                           ByVal materialName As String, _
                                           ByRef indices() As Long) As Long
'------------------------------------------------------------------------------
' Returns count of raw tanks matching materialName. Fills indices() array.
'------------------------------------------------------------------------------
    Dim count As Long
    count = 0
    Dim i As Long
    For i = 0 To state.num_raw_tanks - 1
        If state.raw_tanks(i).material_name = materialName Then
            count = count + 1
            ReDim Preserve indices(0 To count - 1)
            indices(count - 1) = i
        End If
    Next i
    FindAllRawTanksByMaterial = count
End Function


Public Function FindAllProductTanksByProduct(ByRef state As SimState, _
                                              ByVal productName As String, _
                                              ByRef indices() As Long) As Long
'------------------------------------------------------------------------------
' Returns count of product tanks matching productName. Fills indices() array.
'------------------------------------------------------------------------------
    Dim count As Long
    count = 0
    Dim i As Long
    For i = 0 To state.num_product_tanks - 1
        If state.product_tanks(i).product_name = productName Then
            count = count + 1
            ReDim Preserve indices(0 To count - 1)
            indices(count - 1) = i
        End If
    Next i
    FindAllProductTanksByProduct = count
End Function


Public Function TotalRawInventoryByMaterial(ByRef state As SimState, _
                                             ByVal materialName As String) As Double
'------------------------------------------------------------------------------
' Sums inventory across ALL raw tanks holding the given material.
'------------------------------------------------------------------------------
    Dim total As Double
    total = 0
    Dim i As Long
    For i = 0 To state.num_raw_tanks - 1
        If state.raw_tanks(i).material_name = materialName Then
            total = total + state.raw_tanks(i).inventory_bbl
        End If
    Next i
    TotalRawInventoryByMaterial = total
End Function


Public Function TotalProductInventoryByProduct(ByRef state As SimState, _
                                                ByVal productName As String) As Double
'------------------------------------------------------------------------------
' Sums inventory across ALL product tanks holding the given product.
'------------------------------------------------------------------------------
    Dim total As Double
    total = 0
    Dim i As Long
    For i = 0 To state.num_product_tanks - 1
        If state.product_tanks(i).product_name = productName Then
            total = total + state.product_tanks(i).inventory_bbl
        End If
    Next i
    TotalProductInventoryByProduct = total
End Function


Public Sub DepositToRawTanks(ByRef state As SimState, _
                              ByVal materialName As String, _
                              ByRef volume_bbl As Double, _
                              ByRef flags As String)
'------------------------------------------------------------------------------
' Distributes volume_bbl across all raw tanks matching materialName.
' Fills each tank in order until full, then moves to next.
' Reduces volume_bbl by the amount actually deposited.
'------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To state.num_raw_tanks - 1
        If state.raw_tanks(i).material_name = materialName Then
            If volume_bbl <= 0 Then Exit Sub
            Dim space As Double
            space = state.raw_tanks(i).capacity_bbl - state.raw_tanks(i).inventory_bbl
            If space <= 0 Then GoTo NextRawDeposit
            Dim deposit As Double
            If volume_bbl <= space Then
                deposit = volume_bbl
            Else
                deposit = space
            End If
            state.raw_tanks(i).inventory_bbl = state.raw_tanks(i).inventory_bbl + deposit
            volume_bbl = volume_bbl - deposit
        End If
NextRawDeposit:
    Next i
    If volume_bbl > 0 Then
        flags = flags & "OVERFLOW_ALL_RAW:" & materialName & "; "
    End If
End Sub


Public Sub WithdrawFromRawTanks(ByRef state As SimState, _
                                 ByVal materialName As String, _
                                 ByRef volume_bbl As Double, _
                                 ByRef flags As String)
'------------------------------------------------------------------------------
' Withdraws volume_bbl across all raw tanks matching materialName.
' Drains each tank in order, then moves to next.
' Reduces volume_bbl by amount actually withdrawn.
'------------------------------------------------------------------------------
    Dim actual_drawn As Double
    actual_drawn = 0
    Dim i As Long
    For i = 0 To state.num_raw_tanks - 1
        If state.raw_tanks(i).material_name = materialName Then
            If volume_bbl <= 0 Then Exit Sub
            Dim avail As Double
            avail = state.raw_tanks(i).inventory_bbl
            If avail <= 0 Then GoTo NextRawWithdraw
            Dim draw As Double
            If volume_bbl <= avail Then
                draw = volume_bbl
            Else
                draw = avail
            End If
            state.raw_tanks(i).inventory_bbl = state.raw_tanks(i).inventory_bbl - draw
            volume_bbl = volume_bbl - draw
            actual_drawn = actual_drawn + draw
            If state.raw_tanks(i).inventory_bbl < state.raw_tanks(i).min_inv_bbl Then
                flags = flags & "LOW_RAW:" & state.raw_tanks(i).tank_name & "; "
            End If
        End If
NextRawWithdraw:
    Next i
End Sub


Public Sub DepositToProductTanks(ByRef state As SimState, _
                                  ByVal productName As String, _
                                  ByRef volume_bbl As Double, _
                                  ByRef flags As String)
'------------------------------------------------------------------------------
' Distributes volume_bbl across all product tanks matching productName.
'------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To state.num_product_tanks - 1
        If state.product_tanks(i).product_name = productName Then
            If volume_bbl <= 0 Then Exit Sub
            Dim space As Double
            space = state.product_tanks(i).capacity_bbl - state.product_tanks(i).inventory_bbl
            If space <= 0 Then GoTo NextProdDeposit
            Dim deposit As Double
            If volume_bbl <= space Then
                deposit = volume_bbl
            Else
                deposit = space
            End If
            state.product_tanks(i).inventory_bbl = state.product_tanks(i).inventory_bbl + deposit
            volume_bbl = volume_bbl - deposit
        End If
NextProdDeposit:
    Next i
    If volume_bbl > 0 Then
        flags = flags & "OVERFLOW_ALL_PROD:" & productName & "; "
    End If
End Sub


Public Sub WithdrawFromProductTanks(ByRef state As SimState, _
                                     ByVal productName As String, _
                                     ByRef volume_bbl As Double, _
                                     ByRef flags As String)
'------------------------------------------------------------------------------
' Withdraws volume_bbl across all product tanks matching productName.
'------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To state.num_product_tanks - 1
        If state.product_tanks(i).product_name = productName Then
            If volume_bbl <= 0 Then Exit Sub
            Dim avail As Double
            avail = state.product_tanks(i).inventory_bbl
            If avail <= 0 Then GoTo NextProdWithdraw
            Dim draw As Double
            If volume_bbl <= avail Then
                draw = volume_bbl
            Else
                draw = avail
            End If
            state.product_tanks(i).inventory_bbl = state.product_tanks(i).inventory_bbl - draw
            volume_bbl = volume_bbl - draw
            If state.product_tanks(i).inventory_bbl < state.product_tanks(i).min_inv_bbl Then
                flags = flags & "BELOW_MIN:" & state.product_tanks(i).tank_name & "; "
            End If
        End If
NextProdWithdraw:
    Next i
End Sub


Public Function ClampValue(ByVal val As Double, _
                            ByVal min_val As Double, _
                            ByVal max_val As Double) As Double
    If val < min_val Then
        ClampValue = min_val
    ElseIf val > max_val Then
        ClampValue = max_val
    Else
        ClampValue = val
    End If
End Function
