VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm06Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Form_Load()
    Me.OrderBy = "LastFirst"
    ' Temporarily remove the mask
    Me.tboxPhoneNo.InputMask = ""
    Me.tboxPhoneNoMobile.InputMask = ""
End Sub
Private Sub Form_Current()
    Me.Refresh
End Sub
Private Sub Form_Close()
    DoCmd.OpenForm "frm03EstimData"
End Sub

Private Sub tboxPhoneNo_afterUpdate()

    ' Process the entered data
    Dim enteredNumber As String
    If IsNull(Me.tboxPhoneNo.text) Then
        Exit Sub
    Else
        enteredNumber = Me.tboxPhoneNo.text
    End If

    ' Remove non-numeric characters
    Me.tboxPhoneNo.Value = Misc.GetNumericText(enteredNumber)

    ' Reapply the original mask
    Me.tboxPhoneNo.InputMask = "!\(999" & Chr(34) & ") ""000\-0000"""

End Sub


Private Sub tboxPhoneNo_GotFocus()
    Me.tboxPhoneNo.InputMask = ""
End Sub

Private Sub tboxPhoneNoMobile_AfterUpdate()
    ' Temporarily remove the mask
    Me.tboxPhoneNoMobile.InputMask = ""
    
    ' Process the entered data
    Dim enteredNumber As String
    If IsNull(Me.tboxPhoneNoMobile.text) Then
        Exit Sub
    Else
        enteredNumber = Me.tboxPhoneNoMobile.text
    End If

    ' Remove non-numeric characters
    Me.tboxPhoneNoMobile.Value = Misc.GetNumericText(enteredNumber)
    
    ' Reapply the original mask
    Me.tboxPhoneNoMobile.InputMask = "!\(999" & Chr(34) & ") ""000\-0000"""
End Sub

Private Sub tboxPhoneNo_Click()
    Me.tboxPhoneNo.InputMask = ""
End Sub

Private Sub tboxPhoneNoMobile_Click()
    Me.tboxPhoneNoMobile.InputMask = ""
End Sub

Private Sub tboxPhoneNoMobile_GotFocus()
    Me.tboxPhoneNoMobile.InputMask = ""
End Sub