'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | Form_fb_areas.vba                                           |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple / BJR, 10/2/2025                                   |

Option Explicit


Private Sub cmdClose_Click()
    DoCmd.Close acForm, "fg_totals"
End Sub


Private Sub cmdAreasClose_Click()
    DoCmd.Close acForm, "fb_areas"
End Sub


Private Sub cmdMainMenu_Click()
    On Error Resume Next
    DoCmd.Close acForm, "fg_totals"
    On Error GoTo 0
    DoCmd.OpenForm "fa_main"
End Sub


Private Sub cmdAreasGotoMainMenu_Click()
    DoCmd.Close acForm, "fb_areas"
    DoCmd.OpenForm "fa_main"
End Sub
