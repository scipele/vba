VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm07ClientCityState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cboxCityID_NotInList(NewData As String, Response As Integer)
    Dim strMsg As String
    Dim strTitle As String
    Dim strSQL As String

    ' Prompt the user to confirm adding the new company name
    strMsg = "The city name you entered is not in the list. Please check Spelling Before Adding it.  Do you want to add it?"
    strTitle = "Add New City Name"

    If MsgBox(strMsg, vbQuestion + vbYesNo, strTitle) = vbYes Then
        ' Add the new company name to the tlkp2CompanyName table
        strSQL = "INSERT INTO tlkpCity (City) VALUES ('" & NewData & "');"
        CurrentDb.Execute strSQL

        ' Set the response to acDataErrAdded so that the combo box is updated with the new value
        Response = acDataErrAdded
    Else
        ' If the user chooses not to add the new company name, set the response to acDataErrContinue
        Response = acDataErrContinue
    End If
    
End Sub

Private Sub cboxCompany_NotInList(NewData As String, Response As Integer)
    Dim strMsg As String
    Dim strTitle As String
    Dim strSQL As String

    ' Prompt the user to confirm adding the new company name
    strMsg = "The company name you entered is not in the list. Do you want to add it?"
    strTitle = "Add New Company Name"

    If MsgBox(strMsg, vbQuestion + vbYesNo, strTitle) = vbYes Then
        ' Add the new company name to the tlkp2CompanyName table
        strSQL = "INSERT INTO tlkpCompany (Company) VALUES ('" & NewData & "');"
        CurrentDb.Execute strSQL

        ' Set the response to acDataErrAdded so that the combo box is updated with the new value
        Response = acDataErrAdded
    Else
        ' If the user chooses not to add the new company name, set the response to acDataErrContinue
        Response = acDataErrContinue
    End If

End Sub

Private Sub cmdConfirmUnique_Click()
    Call ClientCityStateUniq.confirmClientCityStateUnique
End Sub

Private Sub cmdSaveAndClose_Click()
    Me.Refresh
    Me.Requery
    DoCmd.Close acForm, "frm07ClientCityState", acSaveYes
    DoCmd.OpenForm "frm03EstimData"
    
End Sub

Private Sub Form_Close()
    DoCmd.OpenForm "frm03EstimData"
End Sub