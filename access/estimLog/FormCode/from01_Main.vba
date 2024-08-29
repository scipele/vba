Option Compare Database
Option Explicit

Private Sub cmdEstimDataEntry_Click()
    DoCmd.Close
    DoCmd.OpenForm "frm03EstimData"
End Sub

Private Sub cmdEstimDataAll_Click()
    DoCmd.Close
    DoCmd.OpenForm "frm04EstimDataAll", acFormDS
    DoCmd.GoToRecord , , acLast
End Sub

Private Sub cmdClientCityState_Click()
    DoCmd.Close
    DoCmd.OpenForm "frm07ClientCityState"
    DoCmd.GoToRecord acDataForm, "frm07ClientCityState", acNewRec
End Sub

Private Sub cmdQry01Pending_Click()
    DoCmd.Close
    DoCmd.OpenForm "frm09Pending", acFormDS
End Sub

Private Sub cmdOpenEstimates_Click()
    DoCmd.Close
    DoCmd.OpenForm "frm08OpenEstimates", acFormDS
End Sub

Private Sub cmdVersionHistory_Click()
    DoCmd.Close
    DoCmd.OpenForm "frm10VersionHistory"
    DoCmd.GoToRecord , , acLast
End Sub