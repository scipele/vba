Option Compare Database

Public Sub NavigateToRecordInOpenForm()
    Dim currentRecordID As Variant
    ' Assume you've got the ID as before
    
   
    currentRecordID = Forms("frm09Pending").Recordset.Fields("ID").Value

    If Not IsNull(currentRecordID) Then
        If IsFormOpen("frm03EstimData") Then
            ' If frmDetails is open, apply the filter and requery
            With Forms("frm03EstimData")
                .Filter = "ID = " & currentRecordID
                .FilterOn = True
                .Requery
            End With
        Else
            ' If frm03EstimDataComments is not open, open it with the filter
            DoCmd.OpenForm "frm03EstimData", , , "ID = " & currentRecordID
        End If
    Else
        MsgBox "The current record's ID is not available.", vbInformation, "Error"
    End If
End Sub


Function IsFormOpen(FormName As String) As Boolean
    Dim frm As Form
    On Error Resume Next ' In case the form is not open
    Set frm = Forms(FormName)
    IsFormOpen = Err.Number = 0
    On Error GoTo 0 ' Reset error handling
End Function

