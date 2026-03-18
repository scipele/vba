'==============================================================================
' Filename:    frmDashboard.vba
' Purpose:     Graphical UserForm dashboard for stepping through simulation
'              snapshots with dynamic tank fill indicators
' Dependencies: modTypes, modHelpers, modSimEngine
' By:          T. Sciple, 03/17/2026
'==============================================================================

'##############################################################################
'  USERFORM: frmDashboard
'  Place the code below in a UserForm named "frmDashboard".
'  UserForm properties:
'    Caption = "Plant Logistics Dashboard"
'    Width   = 680
'    Height  = 520
'
'  Controls to add via the VBA editor toolbox:
'    - ScrollBar named "scrStep" (horizontal, Min=1, Max=720, Top=10)
'    - Label named "lblStepInfo" (shows current step/date)
'    - CommandButton named "btnPrev" (Caption="<<")
'    - CommandButton named "btnNext" (Caption=">>")
'    - Image control named "imgCanvas" (or use the UserForm backgound itself)
'
'  The drawing is done directly on the UserForm using GDI-like shape objects
'  that VBA provides via .Controls.Add or manual Shape drawing.  For
'  simplicity we use Labels with colored backgrounds as "tank fill" indicators.
'##############################################################################
Option Explicit

' ── Module-level state ──
Private m_state As SimState
Private m_current_step As Long
Private m_loaded As Boolean

' ── Layout constants ──
Private Const TANK_WIDTH As Single = 60
Private Const TANK_HEIGHT As Single = 80
Private Const TANK_GAP As Single = 20
Private Const LABEL_HEIGHT As Single = 14
Private Const TOP_MARGIN As Single = 60
Private Const LEFT_START As Single = 20
Private Const SECTION_GAP As Single = 40
Private Const CAR_SIZE As Single = 28


Private Sub UserForm_Initialize()
'------------------------------------------------------------------------------
' Load simulation data and first snapshot when form opens.
'------------------------------------------------------------------------------
    Call LoadSimData(m_state)
    Call RunSimLoop(m_state)
    m_current_step = 1
    m_loaded = True

    ' Configure scrollbar
    With Me.scrStep
        .Min = 1
        .Max = m_state.total_steps
        .Value = 1
        .SmallChange = 1
        .LargeChange = CLng(m_state.total_steps / 20)
    End With

    Call DrawDashboard
End Sub


Private Sub scrStep_Change()
    If Not m_loaded Then Exit Sub
    m_current_step = Me.scrStep.Value
    Call DrawDashboard
End Sub


Private Sub btnPrev_Click()
    If m_current_step > 1 Then
        m_current_step = m_current_step - 1
        Me.scrStep.Value = m_current_step
    End If
End Sub


Private Sub btnNext_Click()
    If m_current_step < m_state.total_steps Then
        m_current_step = m_current_step + 1
        Me.scrStep.Value = m_current_step
    End If
End Sub


Private Sub DrawDashboard()
'------------------------------------------------------------------------------
' Clears dynamic controls and redraws all tanks/railcars for current step.
'------------------------------------------------------------------------------
    If Not m_loaded Then Exit Sub

    ' Remove previously drawn dynamic controls
    Call ClearDynamicControls

    Dim snap As StepSnapshot
    snap = m_state.snapshots(m_current_step - 1)

    ' Update step info label
    Me.lblStepInfo.Caption = "Step " & snap.sim_step & " / " & m_state.total_steps & _
                              "    " & Format$(snap.date_time, "yyyy-mm-dd hh:mm") & _
                              IIf(snap.unloading_active, "  [UNLOADING]", "") & _
                              IIf(snap.loading_active, "  [LOADING]", "")

    Dim x_pos As Single
    Dim y_pos As Single
    x_pos = LEFT_START
    y_pos = TOP_MARGIN

    ' ── Section: Inbound Railcars ──
    Call DrawSectionLabel(x_pos, y_pos - 18, "INBOUND")
    If snap.unloading_active Then
        Call DrawRailcar(x_pos, y_pos, "R/C", vbYellow)
        Call DrawRailcar(x_pos + CAR_SIZE + 6, y_pos, "R/C", vbYellow)
    End If
    Call DrawArrow(x_pos + CAR_SIZE * 2 + 14, y_pos + CAR_SIZE / 2 - 4, 30)

    x_pos = x_pos + CAR_SIZE * 2 + SECTION_GAP + 30

    ' ── Section: Raw Material Tanks ──
    Call DrawSectionLabel(x_pos, y_pos - 18, "RAW TANKS")
    Dim i As Long
    For i = 0 To m_state.num_raw_tanks - 1
        Dim raw_pct As Double
        If m_state.raw_tanks(i).capacity_bbl > 0 Then
            raw_pct = snap.raw_inventories(i) / m_state.raw_tanks(i).capacity_bbl
        End If
        Call DrawTank(x_pos + i * (TANK_WIDTH + TANK_GAP), y_pos, _
                      m_state.raw_tanks(i).tank_name, raw_pct, _
                      CLng(snap.raw_inventories(i)))
    Next i

    ' ── Section: Blend Tanks (below raw tanks if applicable) ──
    y_pos = y_pos + TANK_HEIGHT + LABEL_HEIGHT + SECTION_GAP
    If m_state.num_blend_tanks > 0 Then
        x_pos = LEFT_START + CAR_SIZE * 2 + SECTION_GAP + 30
        Call DrawSectionLabel(x_pos, y_pos - 18, "BLEND TANKS")
        For i = 0 To m_state.num_blend_tanks - 1
            Dim bl_pct As Double
            If m_state.blend_tanks(i).capacity_bbl > 0 Then
                bl_pct = snap.blend_inventories(i) / m_state.blend_tanks(i).capacity_bbl
            End If
            Call DrawTank(x_pos + i * (TANK_WIDTH + TANK_GAP), y_pos, _
                          m_state.blend_tanks(i).tank_name, bl_pct, _
                          CLng(snap.blend_inventories(i)))
        Next i
        y_pos = y_pos + TANK_HEIGHT + LABEL_HEIGHT + SECTION_GAP
    End If

    ' ── Section: Processing Units ──
    x_pos = LEFT_START + CAR_SIZE * 2 + SECTION_GAP + 30
    Call DrawSectionLabel(x_pos, y_pos - 18, "PROCESSING UNITS")
    For i = 0 To m_state.num_units - 1
        Call DrawUnitBox(x_pos + i * (TANK_WIDTH + TANK_GAP), y_pos, _
                         m_state.units(i).unit_name, _
                         m_state.units(i).capacity_bbl_day)
    Next i
    y_pos = y_pos + 40 + SECTION_GAP

    ' ── Section: Product Tanks ──
    x_pos = LEFT_START + CAR_SIZE * 2 + SECTION_GAP + 30
    Call DrawSectionLabel(x_pos, y_pos - 18, "PRODUCT TANKS")
    For i = 0 To m_state.num_product_tanks - 1
        Dim pr_pct As Double
        If m_state.product_tanks(i).capacity_bbl > 0 Then
            pr_pct = snap.product_inventories(i) / m_state.product_tanks(i).capacity_bbl
        End If
        Call DrawTank(x_pos + i * (TANK_WIDTH + TANK_GAP), y_pos, _
                      m_state.product_tanks(i).tank_name, pr_pct, _
                      CLng(snap.product_inventories(i)))
    Next i

    ' ── Section: Outbound Railcars ──
    Dim out_x As Single
    out_x = x_pos + (m_state.num_product_tanks) * (TANK_WIDTH + TANK_GAP) + 20
    Call DrawArrow(out_x - 30, y_pos + TANK_HEIGHT / 2 - 4, 30)
    Call DrawSectionLabel(out_x, y_pos - 18, "OUTBOUND")
    If snap.loading_active Then
        Call DrawRailcar(out_x, y_pos + 10, "R/C", vbGreen)
    End If
End Sub


' ── Drawing helper: Tank (rounded rect with fill level) ──
Private Sub DrawTank(ByVal x As Single, ByVal y As Single, _
                      ByVal tankName As String, _
                      ByVal fillPct As Double, _
                      ByVal bblValue As Long)
    Dim tag_prefix As String
    tag_prefix = "dyn_"

    ' Tank outline (label acting as border)
    Dim lbl_border As MSForms.Label
    Set lbl_border = Me.Controls.Add("Forms.Label.1", tag_prefix & "tb_" & tankName)
    With lbl_border
        .Left = x: .Top = y
        .Width = TANK_WIDTH: .Height = TANK_HEIGHT
        .BackColor = &HC0C0C0  ' light gray
        .BorderStyle = fmBorderStyleSingle
        .Caption = ""
        .Tag = "dynamic"
    End With

    ' Fill level (colored label inside)
    Dim fill_height As Single
    fill_height = CSng(TANK_HEIGHT * ClampValue(fillPct, 0, 1))

    Dim lbl_fill As MSForms.Label
    Set lbl_fill = Me.Controls.Add("Forms.Label.1", tag_prefix & "tf_" & tankName)
    With lbl_fill
        .Left = x + 1
        .Top = y + TANK_HEIGHT - fill_height
        .Width = TANK_WIDTH - 2
        .Height = fill_height
        .Caption = ""
        .Tag = "dynamic"

        ' Color: green < 80%, yellow 80-95%, red > 95%
        If fillPct > 0.95 Then
            .BackColor = &H4040FF      ' red
        ElseIf fillPct > 0.8 Then
            .BackColor = &H80FFFF      ' yellow
        ElseIf fillPct < 0.15 Then
            .BackColor = &H4040FF      ' red (too low)
        Else
            .BackColor = &H80FF80      ' green
        End If
    End With

    ' Percentage text overlay
    Dim lbl_pct As MSForms.Label
    Set lbl_pct = Me.Controls.Add("Forms.Label.1", tag_prefix & "tp_" & tankName)
    With lbl_pct
        .Left = x: .Top = y + TANK_HEIGHT / 2 - 7
        .Width = TANK_WIDTH: .Height = LABEL_HEIGHT
        .Caption = Format$(fillPct * 100, "0") & "%"
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleTransparent
        .Font.Bold = True
        .Tag = "dynamic"
    End With

    ' Tank name below
    Dim lbl_name As MSForms.Label
    Set lbl_name = Me.Controls.Add("Forms.Label.1", tag_prefix & "tn_" & tankName)
    With lbl_name
        .Left = x: .Top = y + TANK_HEIGHT + 2
        .Width = TANK_WIDTH: .Height = LABEL_HEIGHT
        .Caption = tankName & vbCrLf & Format$(bblValue, "#,##0")
        .TextAlign = fmTextAlignCenter
        .Font.Size = 7
        .BackStyle = fmBackStyleTransparent
        .Tag = "dynamic"
    End With
End Sub


' ── Drawing helper: Railcar rectangle ──
Private Sub DrawRailcar(ByVal x As Single, ByVal y As Single, _
                          ByVal caption As String, _
                          ByVal clr As Long)
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", "dyn_car_" & x & "_" & y)
    With lbl
        .Left = x: .Top = y
        .Width = CAR_SIZE: .Height = CAR_SIZE
        .BackColor = clr
        .BorderStyle = fmBorderStyleSingle
        .Caption = caption
        .TextAlign = fmTextAlignCenter
        .Font.Size = 7
        .Font.Bold = True
        .Tag = "dynamic"
    End With
End Sub


' ── Drawing helper: Arrow (simple ">>>") ──
Private Sub DrawArrow(ByVal x As Single, ByVal y As Single, _
                       ByVal arrowWidth As Single)
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", "dyn_arr_" & x & "_" & y)
    With lbl
        .Left = x: .Top = y
        .Width = arrowWidth: .Height = 14
        .Caption = Chr$(9654) & Chr$(9654) & Chr$(9654)   ' ▶▶▶
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleTransparent
        .Font.Size = 8
        .Tag = "dynamic"
    End With
End Sub


' ── Drawing helper: Processing unit box ──
Private Sub DrawUnitBox(ByVal x As Single, ByVal y As Single, _
                          ByVal unitName As String, _
                          ByVal capacityBBLDay As Double)
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", "dyn_unit_" & unitName)
    With lbl
        .Left = x: .Top = y
        .Width = TANK_WIDTH: .Height = 32
        .BackColor = &HFFFFC0   ' light cyan
        .BorderStyle = fmBorderStyleSingle
        .Caption = unitName & vbCrLf & Format$(capacityBBLDay, "#,##0") & " B/D"
        .TextAlign = fmTextAlignCenter
        .Font.Size = 7
        .Tag = "dynamic"
    End With
End Sub


' ── Drawing helper: Section label ──
Private Sub DrawSectionLabel(ByVal x As Single, ByVal y As Single, _
                               ByVal caption As String)
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", "dyn_sec_" & caption)
    With lbl
        .Left = x: .Top = y
        .Width = 120: .Height = 14
        .Caption = caption
        .Font.Bold = True
        .Font.Size = 8
        .BackStyle = fmBackStyleTransparent
        .Tag = "dynamic"
    End With
End Sub


' ── Clear all dynamically created controls ──
Private Sub ClearDynamicControls()
    Dim ctrl As MSForms.Control
    Dim ctrls_to_remove() As String
    Dim count As Long
    count = 0

    For Each ctrl In Me.Controls
        If ctrl.Tag = "dynamic" Then
            count = count + 1
            ReDim Preserve ctrls_to_remove(1 To count)
            ctrls_to_remove(count) = ctrl.Name
        End If
    Next ctrl

    Dim i As Long
    For i = 1 To count
        Me.Controls.Remove ctrls_to_remove(i)
    Next i
End Sub
