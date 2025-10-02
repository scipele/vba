'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | Form_fg_totals.vba                                          |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple / BJR, 10/1/2025                                   |

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


Private Sub Form_Load()
    Me.tbx_rt_trips.Value = mNdeCalcs.GetRtTrips
    Me.tbx_pwht_trips.Value = mNdeCalcs.GetPwhtTrips
End Sub