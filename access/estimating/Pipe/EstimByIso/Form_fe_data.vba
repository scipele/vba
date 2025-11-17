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
    'Get old id from user
    Dim old_id As Long
    old_id = InputBox("enter id number to copy ")
    
    'set all textbox values from the user input id
    Me.matl_id = Nz(DLookup("matl_id", "ta_data", "est_id = " & old_id), "Not Found")
    Me.sch_id = Nz(DLookup("sch_id", "ta_data", "est_id = " & old_id), "Not Found")
    Me.flg_rtg_id = Nz(DLookup("flg_rtg_id", "ta_data", "est_id = " & old_id), "Not Found")
    Me.sb_rtg_id = Nz(DLookup("sb_rtg_id", "ta_data", "est_id = " & old_id), "Not Found")
    Me.rt_pct = Nz(DLookup("rt_pct", "ta_data", "est_id = " & old_id), "Not Found")
    Me.instr_mh = Nz(DLookup("instr_mh", "ta_data", "est_id = " & old_id), "Not Found")
    Me.sp_mh = Nz(DLookup("sp_mh", "ta_data", "est_id = " & old_id), "Not Found")
    Me.tie_mh = Nz(DLookup("tie_mh", "ta_data", "est_id = " & old_id), "Not Found")
    Me.supt_qty = Nz(DLookup("supt_qty", "ta_data", "est_id = " & old_id), "Not Found")
    Me.supt_mh = Nz(DLookup("supt_mh", "ta_data", "est_id = " & old_id), "Not Found")
    Me.grout_mh = Nz(DLookup("grout_mh", "ta_data", "est_id = " & old_id), "Not Found")

    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim str_sql As String
    Dim new_est_id As Long
    
    ' Get the current est_id from the main form
    new_est_id = Me.est_id ' Adjust control name if different
    
    ' Verify records exist for the old est_id
    If DCount("*", "tb_qtys", "est_id = " & old_id) = 0 Then
        Exit Sub
    End If
    
    ' Build the INSERT INTO query
    ' Insert fields: all except qty_id
    ' Select: new est_id literal, then the rest from old records
    str_sql = "INSERT INTO tb_qtys (est_id, size_id, spool_qty, str_run_qty, butt_wld_qty, " & _
             "sw_qty, bu_qty, vlv_hnd_qty, make_on_qty, mo_bckwld_qty, cut_bev_qty, " & _
             "spool_mhs, str_run_mhs, butt_wld_mhs, sw_mhs, bu_mhs, vlv_hnd_mhs, " & _
             "make_on_mhs, mo_bckwld_mhs, cut_bev_mhs) " & _
             "SELECT " & new_est_id & ", size_id, spool_qty, str_run_qty, butt_wld_qty, " & _
             "sw_qty, bu_qty, vlv_hnd_qty, make_on_qty, mo_bckwld_qty, cut_bev_qty, " & _
             "spool_mhs, str_run_mhs, butt_wld_mhs, sw_mhs, bu_mhs, vlv_hnd_mhs, " & _
             "make_on_mhs, mo_bckwld_mhs, cut_bev_mhs " & _
             "FROM tb_qtys WHERE est_id = " & old_id
    
    ' Execute the copy
    Set db = CurrentDb
    db.Execute str_sql, dbFailOnError
    
    ' Requery the subform to show the new records
    Me.fb_qtys.Requery
    
    'Total up mhs
    mGetMhs.TotalMhForIso
    Exit Sub
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub


Private Sub Form_Current()
' Check if the form's record source is ta_data and a record is loaded
    If Me.RecordSource = "ta_data" And Not Me.NewRecord Then
        ' Example: Display the value of the "ID" field from the current record
        Dim currentID As Variant
        currentID = Me.est_id ' Replace "ID" with the actual field name in ta_data
        'MsgBox "You are now on record with ID: " & currentID, vbInformation, "Record Changed"
        mGetMhs.TotalMhForIso
    End If
End Sub


Private Sub grout_mh_AfterUpdate()
    mGetMhs.TotalMhForIso
End Sub


Private Sub instr_mh_AfterUpdate()
    mGetMhs.TotalMhForIso
End Sub


Private Sub shop_supt_AfterUpdate()
    mGetMhs.TotalMhForIso
End Sub


Private Sub sp_mh_AfterUpdate()
    mGetMhs.TotalMhForIso
End Sub


Private Sub tie_mh_AfterUpdate()
    mGetMhs.TotalMhForIso
End Sub