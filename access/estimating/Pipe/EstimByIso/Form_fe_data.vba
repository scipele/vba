'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | Form_fe_data.vba                                            |
'| EntryPoint   | command button / after update events                        |
'| Purpose      | various                                                     |
'| Inputs       | user inputs                                                 |
'| Outputs      | various                                                     |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple / BJR, 11/17/2025                                  |

Option Explicit


Private Sub cmdMainMenu_Click()
    On Error Resume Next
    DoCmd.Close acForm, "fe_data"
    On Error GoTo 0
    DoCmd.OpenForm "fa_main"
End Sub


Private Sub cmdViewProjTotals_Click()
    DoCmd.OpenForm "fg_totals"
    DoCmd.SelectObject acForm, "fg_totals", True
End Sub


Private Sub cmdClose_Click()
    On Error Resume Next
    DoCmd.Close acForm, "fe_data"
End Sub


Private Sub cmdCopyPrev_Click()
    On Error GoTo ErrHandler
    
    Dim old_id As Long
    Dim new_est_id As Long
    Dim db As DAO.Database
    Dim str_sql As String
    
    ' 1. Get the ID to copy FROM
    old_id = InputBox("Enter estimate ID number to copy from:", "Copy Previous", "")
    If old_id = 0 Then Exit Sub
    
    ' 2. Make sure we're on a saved main record and capture its est_id
    If Me.Dirty Then Me.Dirty = False          ' Save any pending changes
    If Me.NewRecord Then
        MsgBox "You must save the main estimate first before copying quantities.", vbExclamation
        Exit Sub
    End If
    
    new_est_id = Me.tbx_est_id                       ' This is the correct current est_id
    If new_est_id = 0 Or IsNull(new_est_id) Then
        MsgBox "Current estimate has no ID. Save it first."
        Exit Sub
    End If
    
    ' 3. Copy the header fields from ta_data (old ? current)
    Me.cbx_matl_id = Nz(DLookup("matl_id", "ta_data", "est_id = " & old_id), Null)
    Me.cbx_flg_rtg_id = Nz(DLookup("flg_rtg_id", "ta_data", "est_id = " & old_id), Null)
    Me.cbx_sb_rtg_id = Nz(DLookup("sb_rtg_id", "ta_data", "est_id = " & old_id), Null)
    Me.tbx_rt_pct = Nz(DLookup("rt_pct", "ta_data", "est_id = " & old_id), Null)
    Me.tbx_instr_mh = Nz(DLookup("instr_mh", "ta_data", "est_id = " & old_id), Null)
    Me.tbx_sp_mh = Nz(DLookup("sp_mh", "ta_data", "est_id = " & old_id), Null)
    Me.tbx_tie_mh = Nz(DLookup("tie_mh", "ta_data", "est_id = " & old_id), Null)
    Me.tbx_supt_qty = Nz(DLookup("supt_qty", "ta_data", "est_id = " & old_id), Null)
    Me.tbx_supt_mh = Nz(DLookup("supt_mh", "ta_data", "est_id = " & old_id), Null)
    Me.tbx_grout_mh = Nz(DLookup("grout_mh", "ta_data", "est_id = " & old_id), Null)
    Me.tbx_misc_mh = Nz(DLookup("misc_mh", "ta_data", "est_id = " & old_id), Null)
    
    ' 4. Copy the quantities from tb_qtys
    If DCount("*", "tb_qtys", "est_id = " & old_id) > 0 Then
        str_sql = "INSERT INTO tb_qtys (est_id, size_id, sch_id, spool_qty, str_run_qty, butt_wld_qty, " & _
                  "sw_qty, bu_qty, vlv_hnd_qty, make_on_qty, mo_bckwld_qty, cut_bev_qty, " & _
                  "spool_mhs, str_run_mhs, butt_wld_mhs, sw_mhs, bu_mhs, vlv_hnd_mhs, " & _
                  "make_on_mhs, mo_bckwld_mhs, cut_bev_mhs) " & _
                  "SELECT " & new_est_id & " As est_id, size_id, sch_id, spool_qty, str_run_qty, butt_wld_qty, " & _
                  "sw_qty, bu_qty, vlv_hnd_qty, make_on_qty, mo_bckwld_qty, cut_bev_qty, " & _
                  "spool_mhs, str_run_mhs, butt_wld_mhs, sw_mhs, bu_mhs, vlv_hnd_mhs, " & _
                  "make_on_mhs, mo_bckwld_mhs, cut_bev_mhs " & _
                  "FROM tb_qtys WHERE est_id = " & old_id
        
        Set db = CurrentDb
        db.Execute str_sql, dbFailOnError
    End If
    
    ' 5. Requery the subform to show new rows
    Me.ff_qtys.Form.Requery        ' ? assuming subform control is named ff_qtys
    
    ' 6. Re-total MH
    mGetMhs.CalculateTotalMhForCurrentIso
    
    ' 7. OPTIONAL: Stay on the same record (usually not needed if no navigation happened)
    ' But if you ever lose position, this brings you back safely:
    Dim rs As DAO.Recordset
    Set rs = Me.RecordsetClone
    rs.FindFirst "est_id = " & new_est_id
    If Not rs.NoMatch Then Me.Bookmark = rs.Bookmark
    Set rs = Nothing
    
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub


Private Sub Form_Current()
' Check if the form's record source is ta_data and a record is loaded
    If Me.RecordSource = "ta_data" And Not Me.NewRecord Then
        ' Example: Display the value of the "ID" field from the current record
        Dim currentID As Variant
        currentID = Me.est_id ' Replace "ID" with the actual field name in ta_data
        'MsgBox "You are now on record with ID: " & currentID, vbInformation, "Record Changed"
        mGetMhs.CalculateTotalMhForCurrentIso
    End If
End Sub


Private Sub tbx_instr_mh_AfterUpdate()
    mGetMhs.CalculateTotalMhForCurrentIso
End Sub


Private Sub tbx_sp_mh_AfterUpdate()
    mGetMhs.CalculateTotalMhForCurrentIso
End Sub


Private Sub tbx_supt_mh_AfterUpdate()
    mGetMhs.CalculateTotalMhForCurrentIso
End Sub


Private Sub tbx_grout_mh_AfterUpdate()
    mGetMhs.CalculateTotalMhForCurrentIso
End Sub


Private Sub tbx_misc_mh_AfterUpdate()
    mGetMhs.CalculateTotalMhForCurrentIso
End Sub