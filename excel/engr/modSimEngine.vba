'==============================================================================
' Filename:    modSimEngine.vba
' Purpose:     Core simulation: loads data, runs time-step loop, populates
'              snapshots
' Dependencies: modTypes, modHelpers
' By:          T. Sciple, 03/17/2026
'==============================================================================

'##############################################################################
'  MODULE: modSimEngine  (Standard Module)
'  Core simulation: loads data, runs time-step loop, populates snapshots.
'##############################################################################
Option Explicit


Public Sub LoadSimData(ByRef state As SimState)
'------------------------------------------------------------------------------
' Reads all input tables into the SimState UDT arrays.
'------------------------------------------------------------------------------

    ' ── Run configuration ──
    state.run_duration_days = CLng(GetConfigValue("RunDuration_Days"))
    state.time_step_hrs = CDbl(GetConfigValue("TimeStep_Hours"))
    state.unload_on_weekends = CBool(GetConfigValue("UnloadOnWeekends"))
    state.load_on_weekends = CBool(GetConfigValue("LoadOnWeekends"))
    state.start_date = CDate(GetConfigValue("StartDate"))
    state.total_steps = CLng(state.run_duration_days * 24 / state.time_step_hrs)

    ' ── Raw tanks ──
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets("RawMaterials").ListObjects("tblRawTanks")
    state.num_raw_tanks = tbl.ListRows.Count
    If state.num_raw_tanks > 0 Then
        ReDim state.raw_tanks(0 To state.num_raw_tanks - 1)
        Dim i As Long
        For i = 1 To tbl.ListRows.Count
            With state.raw_tanks(i - 1)
                .tank_name = CStr(tbl.ListRows(i).Range(1, 1).Value)
                .material_name = CStr(tbl.ListRows(i).Range(1, 2).Value)
                .capacity_bbl = CDbl(tbl.ListRows(i).Range(1, 3).Value)
                .inventory_bbl = CDbl(tbl.ListRows(i).Range(1, 4).Value)
                .min_inv_bbl = CDbl(tbl.ListRows(i).Range(1, 5).Value)
            End With
        Next i
    End If

    ' ── Blend tanks ──
    Set tbl = ThisWorkbook.Worksheets("Blending").ListObjects("tblBlendTanks")
    state.num_blend_tanks = tbl.ListRows.Count
    If state.num_blend_tanks > 0 Then
        ReDim state.blend_tanks(0 To state.num_blend_tanks - 1)
        For i = 1 To tbl.ListRows.Count
            With state.blend_tanks(i - 1)
                .tank_name = CStr(tbl.ListRows(i).Range(1, 1).Value)
                .capacity_bbl = CDbl(tbl.ListRows(i).Range(1, 2).Value)
                .inventory_bbl = CDbl(tbl.ListRows(i).Range(1, 3).Value)
            End With
        Next i
    End If

    ' ── Blend recipes ──
    Set tbl = ThisWorkbook.Worksheets("Blending").ListObjects("tblBlendRecipes")
    state.num_recipes = tbl.ListRows.Count
    If state.num_recipes > 0 Then
        ReDim state.blend_recipes(0 To state.num_recipes - 1)
        For i = 1 To tbl.ListRows.Count
            With state.blend_recipes(i - 1)
                .blend_tank_name = CStr(tbl.ListRows(i).Range(1, 1).Value)
                .material_name = CStr(tbl.ListRows(i).Range(1, 2).Value)
                .fraction = CDbl(tbl.ListRows(i).Range(1, 3).Value)
            End With
        Next i
    End If

    ' ── Processing units ──
    Set tbl = ThisWorkbook.Worksheets("Processing").ListObjects("tblUnits")
    state.num_units = tbl.ListRows.Count
    If state.num_units > 0 Then
        ReDim state.units(0 To state.num_units - 1)
        For i = 1 To tbl.ListRows.Count
            With state.units(i - 1)
                .unit_name = CStr(tbl.ListRows(i).Range(1, 1).Value)
                .capacity_bbl_day = CDbl(tbl.ListRows(i).Range(1, 2).Value)
                .feed_source = CStr(tbl.ListRows(i).Range(1, 3).Value)
                .product_name = CStr(tbl.ListRows(i).Range(1, 4).Value)
            End With
        Next i
    End If

    ' ── Product tanks ──
    Set tbl = ThisWorkbook.Worksheets("Products").ListObjects("tblProductTanks")
    state.num_product_tanks = tbl.ListRows.Count
    If state.num_product_tanks > 0 Then
        ReDim state.product_tanks(0 To state.num_product_tanks - 1)
        For i = 1 To tbl.ListRows.Count
            With state.product_tanks(i - 1)
                .tank_name = CStr(tbl.ListRows(i).Range(1, 1).Value)
                .product_name = CStr(tbl.ListRows(i).Range(1, 2).Value)
                .capacity_bbl = CDbl(tbl.ListRows(i).Range(1, 3).Value)
                .inventory_bbl = CDbl(tbl.ListRows(i).Range(1, 4).Value)
                .min_inv_bbl = CDbl(tbl.ListRows(i).Range(1, 5).Value)
            End With
        Next i
    End If

    ' ── Unload spots ──
    Set tbl = ThisWorkbook.Worksheets("RawMaterials").ListObjects("tblUnloadSpots")
    state.num_unload_spots = tbl.ListRows.Count
    If state.num_unload_spots > 0 Then
        ReDim state.unload_spots(0 To state.num_unload_spots - 1)
        For i = 1 To tbl.ListRows.Count
            With state.unload_spots(i - 1)
                .mode_name = CStr(tbl.ListRows(i).Range(1, 1).Value)
                .mode_type = ModeNameToEnum(.mode_name)
                .num_spots = CLng(tbl.ListRows(i).Range(1, 2).Value)
                .avg_unload_hrs = CDbl(tbl.ListRows(i).Range(1, 3).Value)
                .bbl_per_load = CDbl(tbl.ListRows(i).Range(1, 4).Value)
            End With
        Next i
    End If

    ' ── Load spots ──
    Set tbl = ThisWorkbook.Worksheets("Products").ListObjects("tblLoadSpots")
    state.num_load_spots = tbl.ListRows.Count
    If state.num_load_spots > 0 Then
        ReDim state.load_spots(0 To state.num_load_spots - 1)
        For i = 1 To tbl.ListRows.Count
            With state.load_spots(i - 1)
                .mode_name = CStr(tbl.ListRows(i).Range(1, 1).Value)
                .mode_type = ModeNameToEnum(.mode_name)
                .num_spots = CLng(tbl.ListRows(i).Range(1, 2).Value)
                .avg_load_hrs = CDbl(tbl.ListRows(i).Range(1, 3).Value)
                .bbl_per_load = CDbl(tbl.ListRows(i).Range(1, 4).Value)
            End With
        Next i
    End If

    ' ── Arrival schedule ──
    Set tbl = ThisWorkbook.Worksheets("RawMaterials").ListObjects("tblUnloadSchedule")
    state.num_arrivals = tbl.ListRows.Count
    If state.num_arrivals > 0 Then
        ReDim state.arrivals(0 To state.num_arrivals - 1)
        For i = 1 To tbl.ListRows.Count
            With state.arrivals(i - 1)
                .arrival_day = CLng(tbl.ListRows(i).Range(1, 1).Value)
                .mode_name = CStr(tbl.ListRows(i).Range(1, 2).Value)
                .quantity_bbl = CDbl(tbl.ListRows(i).Range(1, 3).Value)
                .material_name = CStr(tbl.ListRows(i).Range(1, 4).Value)
            End With
        Next i
    End If

    ' ── Shipment schedule ──
    Set tbl = ThisWorkbook.Worksheets("Products").ListObjects("tblLoadSchedule")
    state.num_shipments = tbl.ListRows.Count
    If state.num_shipments > 0 Then
        ReDim state.shipments(0 To state.num_shipments - 1)
        For i = 1 To tbl.ListRows.Count
            With state.shipments(i - 1)
                .ship_day = CLng(tbl.ListRows(i).Range(1, 1).Value)
                .product_name = CStr(tbl.ListRows(i).Range(1, 2).Value)
                .quantity_bbl = CDbl(tbl.ListRows(i).Range(1, 3).Value)
                .mode_name = CStr(tbl.ListRows(i).Range(1, 4).Value)
            End With
        Next i
    End If

    ' ── Initialize results array ──
    ReDim state.snapshots(0 To state.total_steps - 1)

    ' ── Initialize active/pending queues as empty ──
    state.num_active_unloads = 0
    state.num_active_loads = 0
    state.num_pending_unloads = 0
    state.num_pending_loads = 0

End Sub


Public Sub RunSimLoop(ByRef state As SimState)
'------------------------------------------------------------------------------
' Main time-step loop.
'------------------------------------------------------------------------------
    Dim step_idx As Long
    Dim current_dt As Date
    Dim current_day As Long
    Dim hrs_per_step As Double
    hrs_per_step = state.time_step_hrs

    For step_idx = 0 To state.total_steps - 1
        current_dt = state.start_date + (step_idx * hrs_per_step) / 24#
        current_day = Int(step_idx * hrs_per_step / 24#) + 1 ' 1-based day

        Dim flags As String
        flags = ""

        Dim is_wknd As Boolean
        is_wknd = IsWeekend(current_dt)

        ' ── Step 1: Check for new arrivals this day ──
        ' (only trigger on the first step of each day to avoid duplicates)
        If IsFirstStepOfDay(step_idx, hrs_per_step) Then
            Call EnqueueArrivals(state, current_day)
        End If

        ' ── Step 2: Process unloading (respect weekend flag) ──
        If (Not is_wknd) Or state.unload_on_weekends Then
            Call ProcessUnloading(state, hrs_per_step, flags)
        End If

        ' ── Step 3: Blending — pull from raw tanks, push to blend tanks ──
        Call ProcessBlending(state, hrs_per_step, flags)

        ' ── Step 4: Unit processing — pull feed, push product ──
        Call ProcessUnits(state, hrs_per_step, flags)

        ' ── Step 5: Check for shipments this day ──
        If IsFirstStepOfDay(step_idx, hrs_per_step) Then
            Call EnqueueShipments(state, current_day)
        End If

        ' ── Step 6: Process loading (respect weekend flag) ──
        If (Not is_wknd) Or state.load_on_weekends Then
            Call ProcessLoading(state, hrs_per_step, flags)
        End If

        ' ── Step 7: Record snapshot ──
        Call RecordSnapshot(state, step_idx, current_dt, flags)

    Next step_idx

End Sub


Private Function IsFirstStepOfDay(ByVal step_idx As Long, _
                                    ByVal hrs_per_step As Double) As Boolean
'------------------------------------------------------------------------------
' Returns True if this step is the first step of its simulation day.
'------------------------------------------------------------------------------
    Dim hour_of_day As Double
    hour_of_day = (step_idx * hrs_per_step) - Int(step_idx * hrs_per_step / 24#) * 24#
    IsFirstStepOfDay = (hour_of_day < hrs_per_step)
End Function


Private Sub EnqueueArrivals(ByRef state As SimState, _
                             ByVal current_day As Long)
'------------------------------------------------------------------------------
' Finds arrivals scheduled for current_day, adds BBLs to pending unload queue.
'------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To state.num_arrivals - 1
        If state.arrivals(i).arrival_day = current_day Then
            ' Add to pending queue
            state.num_pending_unloads = state.num_pending_unloads + 1
            ReDim Preserve state.pending_unload_bbl(0 To state.num_pending_unloads - 1)
            ReDim Preserve state.pending_unload_mat(0 To state.num_pending_unloads - 1)
            state.pending_unload_bbl(state.num_pending_unloads - 1) = state.arrivals(i).quantity_bbl
            state.pending_unload_mat(state.num_pending_unloads - 1) = state.arrivals(i).material_name
        End If
    Next i
End Sub


Private Sub ProcessUnloading(ByRef state As SimState, _
                               ByVal hrs_per_step As Double, _
                               ByRef flags As String)
'------------------------------------------------------------------------------
' Moves material from pending queue into active unloads (limited by spots),
' then advances active unloads by one time step, depositing into raw tanks.
'------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long

    ' Calculate total available spots across all modes
    Dim total_spots As Long
    For i = 0 To state.num_unload_spots - 1
        total_spots = total_spots + state.unload_spots(i).num_spots
    Next i

    ' Start new unloads if spots are available
    Do While state.num_pending_unloads > 0 And state.num_active_unloads < total_spots
        ' Pop first pending item
        Dim mat_name As String
        Dim bbl_qty As Double
        mat_name = state.pending_unload_mat(0)
        bbl_qty = state.pending_unload_bbl(0)

        ' Find unload config for this material to get unload time
        Dim unload_hrs As Double
        unload_hrs = 4  ' default
        For j = 0 To state.num_unload_spots - 1
            ' Use first available spot config (simplified: all modes share spot pool)
            unload_hrs = state.unload_spots(j).avg_unload_hrs
            Exit For
        Next j

        ' Add to active
        state.num_active_unloads = state.num_active_unloads + 1
        ReDim Preserve state.active_unloads(0 To state.num_active_unloads - 1)
        With state.active_unloads(state.num_active_unloads - 1)
            .material_name = mat_name
            .bbl_remaining = bbl_qty
            .hours_remaining = unload_hrs
            .spot_index = state.num_active_unloads - 1
        End With

        ' Remove first pending item (shift array left)
        Call RemoveFirstPendingUnload(state)
    Loop

    ' Advance active unloads
    Dim items_to_remove() As Long
    Dim remove_count As Long
    remove_count = 0

    For i = 0 To state.num_active_unloads - 1
        state.active_unloads(i).hours_remaining = _
            state.active_unloads(i).hours_remaining - hrs_per_step

        If state.active_unloads(i).hours_remaining <= 0 Then
            ' Unload complete — deposit across all raw tanks for this material
            Dim deposit_vol As Double
            deposit_vol = state.active_unloads(i).bbl_remaining
            Dim any_tank As Long
            any_tank = FindRawTankByMaterial(state, state.active_unloads(i).material_name)
            If any_tank >= 0 Then
                Call DepositToRawTanks(state, state.active_unloads(i).material_name, _
                                       deposit_vol, flags)
            Else
                flags = flags & "NO_TANK_FOR:" & state.active_unloads(i).material_name & "; "
            End If

            ' Mark for removal
            remove_count = remove_count + 1
            ReDim Preserve items_to_remove(0 To remove_count - 1)
            items_to_remove(remove_count - 1) = i
        End If
    Next i

    ' Remove completed unloads (iterate backwards)
    If remove_count > 0 Then
        For i = remove_count - 1 To 0 Step -1
            Call RemoveActiveUnload(state, items_to_remove(i))
        Next i
    End If
End Sub


Private Sub RemoveFirstPendingUnload(ByRef state As SimState)
'------------------------------------------------------------------------------
' Removes the first element from pending unload arrays (shift left).
'------------------------------------------------------------------------------
    If state.num_pending_unloads <= 1 Then
        state.num_pending_unloads = 0
        Exit Sub
    End If

    Dim i As Long
    For i = 0 To state.num_pending_unloads - 2
        state.pending_unload_bbl(i) = state.pending_unload_bbl(i + 1)
        state.pending_unload_mat(i) = state.pending_unload_mat(i + 1)
    Next i
    state.num_pending_unloads = state.num_pending_unloads - 1
    ReDim Preserve state.pending_unload_bbl(0 To state.num_pending_unloads - 1)
    ReDim Preserve state.pending_unload_mat(0 To state.num_pending_unloads - 1)
End Sub


Private Sub RemoveActiveUnload(ByRef state As SimState, _
                                 ByVal idx As Long)
'------------------------------------------------------------------------------
' Removes an active unload at given index (shift remaining left).
'------------------------------------------------------------------------------
    If state.num_active_unloads <= 1 Then
        state.num_active_unloads = 0
        Exit Sub
    End If

    Dim i As Long
    For i = idx To state.num_active_unloads - 2
        state.active_unloads(i) = state.active_unloads(i + 1)
    Next i
    state.num_active_unloads = state.num_active_unloads - 1
    ReDim Preserve state.active_unloads(0 To state.num_active_unloads - 1)
End Sub


Private Sub ProcessBlending(ByRef state As SimState, _
                              ByVal hrs_per_step As Double, _
                              ByRef flags As String)
'------------------------------------------------------------------------------
' For each blend tank, pull raw materials according to recipe fractions.
' Blending rate is governed by the downstream unit that feeds from the blend
' tank. If no unit feeds from it, skip.
'------------------------------------------------------------------------------
    If state.num_blend_tanks = 0 Then Exit Sub

    Dim b As Long
    For b = 0 To state.num_blend_tanks - 1
        Dim bt_name As String
        bt_name = state.blend_tanks(b).tank_name

        ' Find the unit that feeds from this blend tank to get pull rate
        Dim pull_rate_bbl_hr As Double
        pull_rate_bbl_hr = 0
        Dim u As Long
        For u = 0 To state.num_units - 1
            If state.units(u).feed_source = bt_name Then
                pull_rate_bbl_hr = state.units(u).capacity_bbl_day / 24#
                Exit For
            End If
        Next u

        If pull_rate_bbl_hr = 0 Then GoTo NextBlendTank

        ' Calculate how much to blend this step
        Dim blend_volume As Double
        Dim space_available As Double
        space_available = state.blend_tanks(b).capacity_bbl - state.blend_tanks(b).inventory_bbl
        blend_volume = pull_rate_bbl_hr * hrs_per_step
        If blend_volume > space_available Then
            blend_volume = space_available
        End If

        If blend_volume <= 0 Then GoTo NextBlendTank

        ' Pull from each raw tank per recipe fraction
        Dim r As Long
        Dim can_blend As Boolean
        can_blend = True

        ' First check total availability across all tanks per material
        For r = 0 To state.num_recipes - 1
            If state.blend_recipes(r).blend_tank_name = bt_name Then
                Dim needed As Double
                needed = blend_volume * state.blend_recipes(r).fraction
                Dim total_avail As Double
                total_avail = TotalRawInventoryByMaterial(state, _
                               state.blend_recipes(r).material_name)
                If total_avail < needed Then
                    can_blend = False
                    flags = flags & "LOW_RAW_TOTAL:" & _
                            state.blend_recipes(r).material_name & "; "
                End If
            End If
        Next r

        ' Execute blend if all materials available (withdraw across multiple tanks)
        If can_blend Then
            For r = 0 To state.num_recipes - 1
                If state.blend_recipes(r).blend_tank_name = bt_name Then
                    Dim draw_vol As Double
                    draw_vol = blend_volume * state.blend_recipes(r).fraction
                    Call WithdrawFromRawTanks(state, _
                            state.blend_recipes(r).material_name, draw_vol, flags)
                End If
            Next r
            state.blend_tanks(b).inventory_bbl = _
                state.blend_tanks(b).inventory_bbl + blend_volume
        End If

NextBlendTank:
    Next b
End Sub


Private Sub ProcessUnits(ByRef state As SimState, _
                           ByVal hrs_per_step As Double, _
                           ByRef flags As String)
'------------------------------------------------------------------------------
' For each processing unit, pull from feed source and push to product tank.
'------------------------------------------------------------------------------
    If state.num_units = 0 Then Exit Sub

    Dim u As Long
    For u = 0 To state.num_units - 1
        Dim rate_bbl_hr As Double
        rate_bbl_hr = state.units(u).capacity_bbl_day / 24#

        Dim process_vol As Double
        process_vol = rate_bbl_hr * hrs_per_step

        ' ── Pull from feed source (raw tank or blend tank) ──
        Dim feed_name As String
        feed_name = state.units(u).feed_source

        ' Try raw tank first (by tank name — could be one of many)
        Dim feed_idx As Long
        feed_idx = FindRawTankIndex(state, feed_name)
        Dim is_raw_feed As Boolean

        If feed_idx >= 0 Then
            ' FeedSource references a specific raw tank name — pull from it
            is_raw_feed = True
            If state.raw_tanks(feed_idx).inventory_bbl < process_vol Then
                process_vol = state.raw_tanks(feed_idx).inventory_bbl
                If process_vol < rate_bbl_hr * hrs_per_step * 0.5 Then
                    flags = flags & "LOW_FEED:" & feed_name & "; "
                End If
            End If
            state.raw_tanks(feed_idx).inventory_bbl = _
                state.raw_tanks(feed_idx).inventory_bbl - process_vol
        Else
            ' Try blend tank
            is_raw_feed = False
            feed_idx = FindBlendTankIndex(state, feed_name)
            If feed_idx >= 0 Then
                If state.blend_tanks(feed_idx).inventory_bbl < process_vol Then
                    process_vol = state.blend_tanks(feed_idx).inventory_bbl
                    If process_vol < rate_bbl_hr * hrs_per_step * 0.5 Then
                        flags = flags & "LOW_FEED:" & feed_name & "; "
                    End If
                End If
                state.blend_tanks(feed_idx).inventory_bbl = _
                    state.blend_tanks(feed_idx).inventory_bbl - process_vol
            Else
                ' Maybe FeedSource is a material name — pull from all raw tanks of that material
                Dim mat_total As Double
                mat_total = TotalRawInventoryByMaterial(state, feed_name)
                If mat_total > 0 Then
                    If mat_total < process_vol Then
                        process_vol = mat_total
                        If process_vol < rate_bbl_hr * hrs_per_step * 0.5 Then
                            flags = flags & "LOW_FEED:" & feed_name & "; "
                        End If
                    End If
                    Dim draw_remaining As Double
                    draw_remaining = process_vol
                    Call WithdrawFromRawTanks(state, feed_name, draw_remaining, flags)
                Else
                    flags = flags & "NO_FEED_SOURCE:" & feed_name & "; "
                    process_vol = 0
                End If
            End If
        End If

        ' ── Push to product tank(s) ── cascade across all tanks for this product
        If process_vol > 0 Then
            Dim deposit_remaining As Double
            deposit_remaining = process_vol
            Call DepositToProductTanks(state, state.units(u).product_name, _
                                       deposit_remaining, flags)
            If deposit_remaining > 0 Then
                ' All product tanks are full
                flags = flags & "NO_PROD_TANK:" & state.units(u).product_name & "; "
            End If
        End If
    Next u
End Sub


Private Sub EnqueueShipments(ByRef state As SimState, _
                               ByVal current_day As Long)
'------------------------------------------------------------------------------
' Finds shipments scheduled for current_day, adds to pending load queue.
'------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To state.num_shipments - 1
        If state.shipments(i).ship_day = current_day Then
            state.num_pending_loads = state.num_pending_loads + 1
            ReDim Preserve state.pending_load_bbl(0 To state.num_pending_loads - 1)
            ReDim Preserve state.pending_load_prod(0 To state.num_pending_loads - 1)
            state.pending_load_bbl(state.num_pending_loads - 1) = state.shipments(i).quantity_bbl
            state.pending_load_prod(state.num_pending_loads - 1) = state.shipments(i).product_name
        End If
    Next i
End Sub


Private Sub ProcessLoading(ByRef state As SimState, _
                             ByVal hrs_per_step As Double, _
                             ByRef flags As String)
'------------------------------------------------------------------------------
' Mirrors ProcessUnloading but for outbound product loading.
'------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long

    Dim total_spots As Long
    For i = 0 To state.num_load_spots - 1
        total_spots = total_spots + state.load_spots(i).num_spots
    Next i

    ' Start new loads
    Do While state.num_pending_loads > 0 And state.num_active_loads < total_spots
        Dim prod_name As String
        Dim bbl_qty As Double
        prod_name = state.pending_load_prod(0)
        bbl_qty = state.pending_load_bbl(0)

        Dim load_hrs As Double
        load_hrs = 4  ' default
        For j = 0 To state.num_load_spots - 1
            load_hrs = state.load_spots(j).avg_load_hrs
            Exit For
        Next j

        ' Check total product inventory across all tanks for this product
        Dim total_prod_inv As Double
        total_prod_inv = TotalProductInventoryByProduct(state, prod_name)
        If total_prod_inv > 0 Then
            If total_prod_inv < bbl_qty Then
                flags = flags & "INSUF_PROD:" & prod_name & "; "
                bbl_qty = total_prod_inv
            End If
            ' Withdraw across all product tanks for this product
            Dim load_draw As Double
            load_draw = bbl_qty
            Call WithdrawFromProductTanks(state, prod_name, load_draw, flags)
        Else
            flags = flags & "INSUF_PROD:" & prod_name & "; "
            bbl_qty = 0
        End If

        state.num_active_loads = state.num_active_loads + 1
        ReDim Preserve state.active_loads(0 To state.num_active_loads - 1)
        With state.active_loads(state.num_active_loads - 1)
            .product_name = prod_name
            .bbl_remaining = bbl_qty
            .hours_remaining = load_hrs
            .spot_index = state.num_active_loads - 1
        End With

        Call RemoveFirstPendingLoad(state)
    Loop

    ' Advance active loads
    Dim items_to_remove() As Long
    Dim remove_count As Long
    remove_count = 0

    For i = 0 To state.num_active_loads - 1
        state.active_loads(i).hours_remaining = _
            state.active_loads(i).hours_remaining - hrs_per_step

        If state.active_loads(i).hours_remaining <= 0 Then
            ' Loading complete — product already deducted
            remove_count = remove_count + 1
            ReDim Preserve items_to_remove(0 To remove_count - 1)
            items_to_remove(remove_count - 1) = i
        End If
    Next i

    If remove_count > 0 Then
        For i = remove_count - 1 To 0 Step -1
            Call RemoveActiveLoad(state, items_to_remove(i))
        Next i
    End If
End Sub


Private Sub RemoveFirstPendingLoad(ByRef state As SimState)
    If state.num_pending_loads <= 1 Then
        state.num_pending_loads = 0
        Exit Sub
    End If

    Dim i As Long
    For i = 0 To state.num_pending_loads - 2
        state.pending_load_bbl(i) = state.pending_load_bbl(i + 1)
        state.pending_load_prod(i) = state.pending_load_prod(i + 1)
    Next i
    state.num_pending_loads = state.num_pending_loads - 1
    ReDim Preserve state.pending_load_bbl(0 To state.num_pending_loads - 1)
    ReDim Preserve state.pending_load_prod(0 To state.num_pending_loads - 1)
End Sub


Private Sub RemoveActiveLoad(ByRef state As SimState, _
                               ByVal idx As Long)
    If state.num_active_loads <= 1 Then
        state.num_active_loads = 0
        Exit Sub
    End If

    Dim i As Long
    For i = idx To state.num_active_loads - 2
        state.active_loads(i) = state.active_loads(i + 1)
    Next i
    state.num_active_loads = state.num_active_loads - 1
    ReDim Preserve state.active_loads(0 To state.num_active_loads - 1)
End Sub


Private Sub RecordSnapshot(ByRef state As SimState, _
                             ByVal step_idx As Long, _
                             ByVal current_dt As Date, _
                             ByVal flags As String)
'------------------------------------------------------------------------------
' Captures current inventories and flags into the snapshots array.
'------------------------------------------------------------------------------
    With state.snapshots(step_idx)
        .sim_step = step_idx + 1
        .date_time = current_dt

        ' Raw tank inventories
        If state.num_raw_tanks > 0 Then
            ReDim .raw_inventories(0 To state.num_raw_tanks - 1)
            Dim i As Long
            For i = 0 To state.num_raw_tanks - 1
                .raw_inventories(i) = state.raw_tanks(i).inventory_bbl
            Next i
        End If

        ' Blend tank inventories
        If state.num_blend_tanks > 0 Then
            ReDim .blend_inventories(0 To state.num_blend_tanks - 1)
            For i = 0 To state.num_blend_tanks - 1
                .blend_inventories(i) = state.blend_tanks(i).inventory_bbl
            Next i
        End If

        ' Product tank inventories
        If state.num_product_tanks > 0 Then
            ReDim .product_inventories(0 To state.num_product_tanks - 1)
            For i = 0 To state.num_product_tanks - 1
                .product_inventories(i) = state.product_tanks(i).inventory_bbl
            Next i
        End If

        ' Unit throughputs (actual capacity for this step)
        If state.num_units > 0 Then
            ReDim .unit_throughputs(0 To state.num_units - 1)
            For i = 0 To state.num_units - 1
                .unit_throughputs(i) = state.units(i).capacity_bbl_day / 24# * _
                                        state.time_step_hrs
            Next i
        End If

        .unloading_active = (state.num_active_unloads > 0)
        .loading_active = (state.num_active_loads > 0)
        .flag_text = flags
    End With
End Sub
