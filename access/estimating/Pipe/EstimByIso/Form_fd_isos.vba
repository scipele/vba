'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | Form_fd_isos.vba                                            |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple / BJR, 11/17/2025                                  |

Option Explicit


Private Sub cmdClose_Click()
    On Error Resume Next
    DoCmd.Close acForm, "fg_totals"
End Sub


Private Sub cmdMainMenu_Click()
    On Error Resume Next
    DoCmd.Close acForm, "fg_totals"
    On Error GoTo 0
    DoCmd.OpenForm "fa_main"
End Sub


Private Sub cmdSpecsClose_Click()
    On Error Resume Next
    DoCmd.Close acForm, "fc_specs"
End Sub


Private Sub cmdSpecsToMainMenu_Click()
    On Error Resume Next
    DoCmd.Close acForm, "fc_specs"
    On Error GoTo 0
    DoCmd.OpenForm "fa_main"
End Sub