'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | Form_ff_qtys_code.vba                                       |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 11/17/2025                                        |

Option Explicit
Option Compare Database

    
Private Sub Form_BeforeInsert(Cancel As Integer)
    If IsNull(Parent!est_id) Then
        Parent!iso = "Need Iso Number "
        Parent.Refresh
        MsgBox ("Created a parent record, so now please re-enter the size/schedule")
    End If
End Sub


'****************************************************************************
'********************* Before Update Event Code *****************************
'*************************** Quantities *************************************
'****************************************************************************
Private Sub spool_qty_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Or IsNull(Me.sch_id) Then
        MsgBox ("missing size and/or schedule")
        Me.spool_qty.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub str_run_qty_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Or IsNull(Me.sch_id) Then
        MsgBox ("missing size and/or schedule")
        Me.str_run_qty.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub fld_butt_wld_qty_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Or IsNull(Me.sch_id) Then
        MsgBox ("missing size and/or schedule")
        Me.fld_butt_wld_qty.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub fld_sw_qty_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.fld_sw_qty.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub bu_qty_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.bu_qty.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub vlv_handling_qty_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.vlv_handling_qty.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub make_on_qty_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.make_on_qty.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub mo_bckwld_qty_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.mo_bckwld_qty.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub cut_bev_qty_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.cut_bev_qty.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


'****************************************************************************
'********************* Before Update Event Code *****************************
'***************************** Manhours *************************************
'****************************************************************************
Private Sub spool_mhs_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.spool_mhs.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub str_run_mhs_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.str_run_mhs.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub butt_wld_mhs_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.butt_wld_mhs.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub sw_mhs_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.sw_mhs.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub bu_mhs_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.bu_mhs.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub vlv_hnd_mhs_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.vlv_handling_qty.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub make_on_mhs_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.make_on_mhs.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub mo_bckwld_mhs_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.mo_bckwld_mhs.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub cut_bev_mhs_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.cut_bev_mhs.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


'****************************************************************************
'***************** After Update Event Code - Quantities**********************
'****************************************************************************
Private Sub size_id_AfterUpdate()
    'Turn off screen painting - eliminates 99% of flickering
    Application.Echo False
    
    On Error GoTo CleanExit
    
    'Capture the primary key as early as possible
    Dim currentPK As Variant
    currentPK = Me.qty_id
    
    'Capture size_id before anything changes it
    Dim sizeIDValue As Variant
    sizeIDValue = Me.size_id
    
    '=== Update all the dependent manhours fields ===
    If Not IsNull(Me.spool_qty) Then Call spool_qty_AfterUpdate
    If Not IsNull(Me.str_run_qty) Then Call str_run_qty_AfterUpdate
    If Not IsNull(Me.butt_wld_qty) Then Call fld_butt_wld_qty_AfterUpdate
    If Not IsNull(Me.sw_qty) Then Call fld_sw_qty_AfterUpdate
    If Not IsNull(Me.bu_qty) Then Call bu_qty_AfterUpdate
    If Not IsNull(Me.vlv_handling_qty) Then Call vlv_handling_qty_AfterUpdate
    If Not IsNull(Me.make_on_qty) Then Call make_on_qty_AfterUpdate
    If Not IsNull(Me.mo_bckwld_qty) Then Call mo_bckwld_qty_AfterUpdate
    If Not IsNull(Me.cut_bev_qty) Then Call cut_bev_qty_AfterUpdate
    
    'Recalculate totals
    mGetMhs.TotalMhForIso
    
    'Only ONE requery at the very end, if needed
    If Me.Parent.ParentHasKey Then
        Me.Requery
        Call ReturnToRecordByPK(currentPK)
    End If
    
    Me.sch_id.SetFocus

CleanExit:
    Application.Echo True   'Always turn it back on
    'If you re-queried, repaint once
    If Me.Parent.ParentHasKey Then Me.Repaint
End Sub


Private Sub sch_id_AfterUpdate()
    'Capture the primary key 'PK' as early as possible
    Dim currentPK As Variant
    currentPK = Me.qty_id
    
    'check to see if any manhours need to be updated following a potential schedule change
    If Not IsNull(Me.spool_qty) Then Call spool_qty_AfterUpdate
    If Not IsNull(Me.str_run_qty) Then Call str_run_qty_AfterUpdate
    If Not IsNull(Me.butt_wld_qty) Then Call fld_butt_wld_qty_AfterUpdate
    If Not IsNull(Me.cut_bev_qty) Then Call cut_bev_qty_AfterUpdate
    If Me.Parent.ParentHasKey Then
        Me.Requery
    Else
        Call Me.Parent.EnsureParentRecord
    End If
    mGetMhs.TotalMhForIso
    Call ReturnToRecordByPK(currentPK)
    Me.spool_qty.SetFocus
End Sub


Private Sub spool_qty_AfterUpdate()
    'Capture the primary key 'PK' as early as possible
    Dim currentPK As Variant
    currentPK = Me.qty_id
    
    If IsNull(Me.spool_qty) Then
        Me.spool_mhs.Value = Null
    Else
        Me.spool_mhs.Value = mGetMhs.GetSpoolMhs(Me.size_id.Value, Me.sch_id)
    End If
    
    'safe requery
    If Me.Parent.ParentHasKey Then
        Me.Requery
    Else
        Call Me.Parent.EnsureParentRecord
    End If
    
    mGetMhs.TotalMhForIso
    Call ReturnToRecordByPK(currentPK)
    Me.str_run_qty.SetFocus
End Sub


Private Sub str_run_qty_AfterUpdate()
    'Capture the primary key 'PK' as early as possible
    Dim currentPK As Variant
    currentPK = Me.qty_id
    
    ' Update MHS value
    If IsNull(Me.str_run_qty) Then
        Me.str_run_mhs = Null
    Else
        Me.str_run_mhs = mGetMhs.GetStrMhs(Me.size_id.Value, Me.sch_id)
    End If
    
    'safe requery
    If Me.Parent.ParentHasKey Then
        Me.Requery
    Else
        Call Me.Parent.EnsureParentRecord
    End If
    
    mGetMhs.TotalMhForIso
    Call ReturnToRecordByPK(currentPK)
    Me.fld_butt_wld_qty.SetFocus
End Sub


Private Sub fld_butt_wld_qty_AfterUpdate()
    'Capture the primary key 'PK' as early as possible
    Dim currentPK As Variant
    currentPK = Me.qty_id
    
    ' Update MHS value
    If IsNull(Me.butt_wld_qty) Then
        Me.butt_wld_mhs.Value = Null
    Else
        Me.butt_wld_mhs.Value = mGetMhs.GetBwMhs(Me.size_id.Value, Me.sch_id)
    End If
    
    'safe requery
    If Me.Parent.ParentHasKey Then
        Me.Requery
    Else
        Call Me.Parent.EnsureParentRecord
    End If
    
    mGetMhs.TotalMhForIso
    Call ReturnToRecordByPK(currentPK)
    Me.fld_sw_qty.SetFocus
End Sub


Private Sub fld_sw_qty_AfterUpdate()
    'Capture the primary key 'PK' as early as possible
    Dim currentPK As Variant
    currentPK = Me.qty_id
    
    ' Update MHS value
    If IsNull(Me.sw_qty) Then
        Me.sw_mhs.Value = Null
    Else
        Me.sw_mhs.Value = mGetMhs.GetSwMhs(Me.size_id.Value)
    End If
    
    'safe requery
    If Me.Parent.ParentHasKey Then
        Me.Requery
    Else
        Call Me.Parent.EnsureParentRecord
    End If
    
    mGetMhs.TotalMhForIso
    Call ReturnToRecordByPK(currentPK)
    Me.bu_qty.SetFocus
End Sub


Private Sub bu_qty_AfterUpdate()
    'Capture the primary key 'PK' as early as possible
    Dim currentPK As Variant
    currentPK = Me.qty_id
    
    ' Update MHS value
    If IsNull(Me.bu_qty) Then
        Me.bu_mhs.Value = Null
    Else
        Me.bu_mhs.Value = mGetMhs.GetBuMhs(Me.size_id.Value)
    End If
    
    'safe requery
    If Me.Parent.ParentHasKey Then
        Me.Requery
    Else
        Call Me.Parent.EnsureParentRecord
    End If
    
    mGetMhs.TotalMhForIso
    Call ReturnToRecordByPK(currentPK)
    Me.vlv_handling_qty.SetFocus
End Sub


Private Sub vlv_handling_qty_AfterUpdate()
    'Capture the primary key 'PK' as early as possible
    Dim currentPK As Variant
    currentPK = Me.qty_id
    
    ' Update MHS value
    If IsNull(Me.vlv_handling_qty) Then
        Me.vlv_hnd_mhs.Value = Null
    Else
        Me.vlv_hnd_mhs.Value = mGetMhs.GetVhMhs(Me.size_id.Value)
    End If
    
    'safe requery
    If Me.Parent.ParentHasKey Then
        Me.Requery
    Else
        Call Me.Parent.EnsureParentRecord
    End If
    
    mGetMhs.TotalMhForIso
    Call ReturnToRecordByPK(currentPK)
    Me.make_on_qty.SetFocus
End Sub


Private Sub make_on_qty_AfterUpdate()
    'Capture the primary key 'PK' as early as possible
    Dim currentPK As Variant
    currentPK = Me.qty_id
    
    ' Update MHS value
    If IsNull(Me.make_on_qty) Then
        Me.make_on_mhs.Value = Null
    Else
        Me.make_on_mhs.Value = mGetMhs.GetMoMhs(Me.size_id.Value)
    End If
    
    'safe requery
    If Me.Parent.ParentHasKey Then
        Me.Requery
    Else
        Call Me.Parent.EnsureParentRecord
    End If
    
    mGetMhs.TotalMhForIso
    Call ReturnToRecordByPK(currentPK)
    Me.mo_bckwld_qty.SetFocus
End Sub


Private Sub mo_bckwld_qty_AfterUpdate()
    'Capture the primary key 'PK' as early as possible
    Dim currentPK As Variant
    currentPK = Me.qty_id
    ' Update MHS value
    If IsNull(Me.mo_bckwld_qty) Then
        Me.mo_bckwld_mhs.Value = Null
    Else
        Me.mo_bckwld_mhs.Value = mGetMhs.GetMbMhs(Me.size_id.Value)
    End If
    
    'safe requery
    If Me.Parent.ParentHasKey Then
        Me.Requery
    Else
        Call Me.Parent.EnsureParentRecord
    End If
    
    mGetMhs.TotalMhForIso
    Call ReturnToRecordByPK(currentPK)
    Me.cut_bev_qty.SetFocus
End Sub


Private Sub cut_bev_qty_AfterUpdate()
    'Capture the primary key 'PK' as early as possible
    Dim currentPK As Variant
    currentPK = Me.qty_id
    
    ' Update MHS value
    If IsNull(Me.cut_bev_qty) Then
        Me.cut_bev_mhs.Value = Null
    Else
        Me.cut_bev_mhs.Value = mGetMhs.GetCbMhs(Me.size_id.Value, Me.sch_id)
    End If
    
    'safe requery
    If Me.Parent.ParentHasKey Then
        Me.Requery
    Else
        Call Me.Parent.EnsureParentRecord
    End If
    
    mGetMhs.TotalMhForIso
    Call ReturnToRecordByPK(currentPK)
    Me.cut_bev_qty.SetFocus
End Sub


'****************************************************************************
'***************************** Misc Functions********************************
'****************************************************************************
Private Function IsNotInLibrary(ByVal val As Double) As Boolean
    If val = 0 Then
        MsgBox ("unit man hour not found in table 'tx_mhs'")
       IsNotInLibrary = True
    Else
        IsNotInLibrary = False
    End If
End Function


Private Sub ReturnToRecordByPK(Optional pkValue As Variant)
    If IsNull(pkValue) Then Exit Sub
    
    Dim rs As DAO.Recordset
    Set rs = Me.RecordsetClone
    
    rs.FindFirst "[qty_id] = " & pkValue
    If Not rs.NoMatch Then
        Me.Bookmark = rs.Bookmark
    End If
    
    Set rs = Nothing
End Sub