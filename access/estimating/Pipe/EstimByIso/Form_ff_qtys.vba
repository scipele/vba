'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | Form_ff_qtys_code.vba                                       |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 10/2/2025                                         |

Option Explicit
Option Compare Database


Private Sub spool_qty_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.spool_qty.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub spool_qty_AfterUpdate()
    If IsNull(Me.spool_qty) Then
        Me.spool_mhs.Value = Null
    Else
        Me.spool_mhs.Value = mGetMhs.GetSpoolMhs(Me.size_id.Value)
    End If
    Me.Requery
    mGetMhs.TotalMhForIso
End Sub


Private Sub str_run_qty_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.str_run_qty.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub str_run_qty_AfterUpdate()
    If IsNull(Me.str_run_qty) Then
        Me.str_run_mhs.Value = Null
    Else
        Me.str_run_mhs.Value = mGetMhs.GetStrMhs(Me.size_id.Value)
    End If
    Me.Requery
    mGetMhs.TotalMhForIso
End Sub


Private Sub fld_butt_wld_qty_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.fld_butt_wld_qty.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub fld_butt_wld_qty_AfterUpdate()
    If IsNull(Me.butt_wld_qty) Then
        Me.butt_wld_mhs.Value = Null
    Else
        Me.butt_wld_mhs.Value = mGetMhs.GetBwMhs(Me.size_id.Value)
    End If
    Me.Requery
    mGetMhs.TotalMhForIso
End Sub


Private Sub fld_sw_qty_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.fld_sw_qty.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub fld_sw_qty_AfterUpdate()
    If IsNull(Me.sw_qty) Then
        Me.sw_mhs.Value = Null
    Else
        Me.sw_mhs.Value = mGetMhs.GetSwMhs(Me.size_id.Value)
    End If
    Me.Requery
    mGetMhs.TotalMhForIso
End Sub


Private Sub bu_qty_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.bu_qty.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub bu_qty_AfterUpdate()
    If IsNull(Me.bu_qty) Then
        Me.bu_mhs.Value = Null
    Else
        Me.bu_mhs.Value = mGetMhs.GetBuMhs(Me.size_id.Value)
    End If
    Me.Requery
    mGetMhs.TotalMhForIso
End Sub


Private Sub sw_mhs_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.sw_mhs.Undo
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


Private Sub vlv_handling_qty_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.vlv_handling_qty.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub vlv_handling_qty_AfterUpdate()
    If IsNull(Me.vlv_handling_qty) Then
        Me.vlv_hnd_mhs.Value = Null
    Else
        Me.vlv_hnd_mhs.Value = mGetMhs.GetVhMhs(Me.size_id.Value)
    End If
    Me.Requery
    mGetMhs.TotalMhForIso
End Sub


Private Sub make_on_qty_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.make_on_qty.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub make_on_qty_AfterUpdate()
    If IsNull(Me.make_on_qty) Then
        Me.make_on_mhs.Value = Null
    Else
        Me.make_on_mhs.Value = mGetMhs.GetMoMhs(Me.size_id.Value)
    End If
    Me.Requery
    mGetMhs.TotalMhForIso
End Sub


Private Sub mo_bckwld_qty_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.mo_bckwld_qty.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub mo_bckwld_qty_AfterUpdate()
    If IsNull(Me.mo_bckwld_qty) Then
        Me.mo_bckwld_mhs.Value = Null
    Else
        Me.mo_bckwld_mhs.Value = mGetMhs.GetMbMhs(Me.size_id.Value)
    End If
    Me.Requery
    mGetMhs.TotalMhForIso
End Sub


Private Sub cut_bev_qty_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.cut_bev_qty.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub


Private Sub cut_bev_qty_AfterUpdate()
    If IsNull(Me.cut_bev_qty) Then
        Me.cut_bev_mhs.Value = Null
    Else
        Me.cut_bev_mhs.Value = mGetMhs.GetCbMhs(Me.size_id.Value)
    End If
    Me.Requery
    mGetMhs.TotalMhForIso
End Sub


Private Sub str_run_mhs_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.str_run_mhs.Undo
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


Private Sub butt_wld_mhs_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.butt_wld_mhs.Undo
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


Private Sub spool_mhs_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.size_id) Then
        MsgBox ("missing size")
        Me.spool_mhs.Undo
        Cancel = True ' Cancel the update to prevent saving the record
    End If
End Sub



Private Function IsNotInLibrary(ByVal val As Double) As Boolean
    If val = 0 Then
        MsgBox ("unit man hour not found in table 'tx_mhs'")
       IsNotInLibrary = True
    Else
        IsNotInLibrary = False
    End If
End Function