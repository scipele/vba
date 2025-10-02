'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | Form_fa_main.vba                                            |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple / BJR, 10/2/2025                                   |

Option Explicit


Private Sub cmdClose_Click()
    On Error Resume Next
    DoCmd.Close acForm, "fa_main"
End Sub


Private Sub cmdDeleteAll_Click()
    mSetup.DeleteAllRecords
End Sub


Private Sub cmdGotoFadata_Click()
    On Error Resume Next
    DoCmd.Close acForm, "fa_main"
    On Error GoTo 0
    DoCmd.OpenForm "fe_data"
End Sub


Private Sub cmdGotoTotals_Click()
    On Error Resume Next
    DoCmd.Close acForm, "fa_main"
    On Error GoTo 0
    DoCmd.OpenForm "fg_totals"
End Sub


Private Sub cmdIsoList_Click()
    DoCmd.OpenForm "fd_isos", acFormDS
End Sub


Private Sub cmdMainSpecs_Click()
    DoCmd.OpenForm "fc_specs", acFormDS
End Sub


Private Sub cmdOpenQueryByIso_Click()
    DoCmd.OpenQuery "qb_by_iso", acViewNormal, acReadOnly
End Sub


Private Sub cmdSetupAreas_Click()
    DoCmd.OpenForm "fb_areas", acFormDS
End Sub