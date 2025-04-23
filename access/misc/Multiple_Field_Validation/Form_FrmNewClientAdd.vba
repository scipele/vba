Option Compare Database
Option Explicit

'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | Form_FrmNewClientAdd.vba                                    |
'| EntryPoint   | cmdEnterClientCityState, or AfterUpdate Events              |
'| Purpose      | Handle Events and call module level code                    |
'| Inputs       | Combo Boxes                                                 |
'| Outputs      | none                                                        |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 12/13/2024                                        |


Private Sub cbxCity_AfterUpdate()
    If Not IsNull(Me.Controls("cbxCity").Value) Then
        Call ModClientCityState.AddNewTableData("tlkpCity", "city", Me.Controls("cbxCity").Value)
    End If
End Sub


Private Sub cbxCompany_AfterUpdate()
    If Not IsNull(Me.Controls("cbxCompany").Value) Then
        Call ModClientCityState.AddNewTableData("tlkpCompany", "company", Me.Controls("cbxCompany").Value)
    End If
End Sub


Private Sub cmdEnterClientCityState_Click()
    Call ModClientCityState.confirmClientCityStateUnique
End Sub