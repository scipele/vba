VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm03EstimData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cboClientCityState_Click()
    Me.Refresh
End Sub

Private Sub ClientContact_Click()
    Me.Refresh
End Sub

Private Sub cboClientContact_Click()
    cboClientContact.Requery
End Sub

Private Sub cboClientContact_Gotfocus()
    Me.Refresh
End Sub

Private Sub cboClientCityState_Gotfocus()
    Me.Refresh
End Sub

Private Sub cmdAddClientCityState_Click()
    DoCmd.OpenForm "frm07ClientCityState"
    DoCmd.GoToRecord acDataForm, "frm07ClientCityState", acNewRec
End Sub

Private Sub cmdAddNewEstim_Click()
    On Error GoTo ErrorHandler
    Dim strYr As String
    strYr = Right(Year(Date), 2)
    Dim maxExistNo As String
    maxExistNo = DMax("[EstimNo]", "tblEstimData", "[ID] > 0")
        
    'If its a new year then start over with numbering at 0 where it will be added to by a 1 further down
    Dim maxExistNo_woYr As Long
    If (CLng(strYr) - CLng(Left(maxExistNo, 2))) > 0 Then
        maxExistNo_woYr = 0
    Else
        maxExistNo_woYr = Right(maxExistNo, 4)
    End If
    
    'Pad Estimate numbers with Zeros depending on the value
    Dim padZero As String
    If maxExistNo_woYr < 9 Then padZero = "000"
    If maxExistNo_woYr >= 9 And maxExistNo_woYr < 99 Then padZero = "00"
    If maxExistNo_woYr >= 99 And maxExistNo_woYr < 999 Then padZero = "0"
        
    Dim EstimateLog_BE As DAO.Database
    Set EstimateLog_BE = CurrentDb
    Dim rsttblEstimData As DAO.Recordset
    Set rsttblEstimData = EstimateLog_BE.OpenRecordset("tblEstimData")
     
    With rsttblEstimData
        .AddNew
        !EstimNo = strYr & "-" & padZero & maxExistNo_woYr + 1
        .Update
        .MoveFirst
        .MoveLast
        .Close
    End With
    Me.Requery
    
    DoCmd.RunCommand acCmdRecordsGoToLast
    
    Exit Sub ' Exit the subroutine if there are no errors
  
ErrorHandler:
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Err.Clear
    
End Sub

Private Sub cmdMkFolder_Click()
    Folders.CreateFolderStructure
End Sub

Private Sub cmdOpenFolder_Click()
    Folders.OpenFolderStructure
End Sub

Private Sub cmdRefreshSave_Click()
    Dim curRec As Long
    curRec = Me.CurrentRecord
    Me.Requery
    DoCmd.Save acForm, "frm03EstimData"
    DoCmd.GoToRecord acDataForm, "frm03EstimData", acGoTo, curRec
End Sub

Private Sub cmdAddClientContact_Click()
    DoCmd.OpenForm "frm07ClientCityState", , , "[tlkpClientCityState].[ID]=" & cboClientCityState
    Forms!frm07ClientCityState.SetFocus
    Forms!frm07ClientCityState!frm06Client.SetFocus
End Sub

Private Sub Command643_Click()
    RecordNavigation.NavigateToRecordInOpenForm
End Sub

Private Sub Form_Close()
    DoCmd.OpenForm "frm01Main"
End Sub

Private Sub Form_GotFocus()
    Me.Requery
    Me.Refresh
    cboClientContact.Requery
    cboClientCityState.Requery
End Sub

Private Sub Form_Load()
    DoCmd.RunCommand acCmdRecordsGoToLast
    Application.SetOption "Default Find/Replace Behavior", 1
    Me.EstimNo.Locked = True
    DoCmd.Restore
End Sub

Private Sub Form_Current()
    Me.Refresh
End Sub