Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | frmUserDataEntry.vba                                        |
'| EntryPoint   | InitializeForm called from Main                             |
'| Purpose      | form code to get user name and which test they are taking   |
'| Inputs       | User Entry                                                  |
'| Outputs      | pass variables back to module code                          |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 12/21/2024                                        |


Public Sub InitializeForm()
    modMakeTest.isDataEntered = False
    PopulateTestList ' Call the subroutine to populate the combo box
    Me.Show
End Sub


Private Sub cmdSubmitUserData_Click()

    ' Check if Last Name, First Name, and Test are filled in
    If Me.tbxLastName = "" Or Me.tbxFirstName = "" Or Me.tbxMiddleName = "" Then
        MsgBox "Please fill in all name fields.", vbExclamation, "Missing Information"
        Exit Sub ' Exit the subroutine if fields are missing
    End If

    ' Check if a test has been selected
    If Me.cbxSelectedTest.Value = "" Then
        MsgBox "Please select a test.", vbExclamation, "Missing Information"
        Exit Sub ' Exit if no test is selected
    End If

    ' Concatenate the full name from text boxes
    Dim currentUserFullName As String
    currentUserFullName = Me.tbxLastName & "_" & Me.tbxFirstName & "_" & Me.tbxMiddleName
    
    ' Pass variables to the main module code
    Call modMakeTest.PassUserDataToMod(currentUserFullName, _
                                       Me.cbxSelectedTest.ListIndex + 1, _
                                       Me.cbxSelectedTest.Value)
    modMakeTest.isDataEntered = True
    
    'unload the form so that the program exectuion will return to main
    Unload Me

End Sub


Private Sub PopulateTestList()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim tblRow As ListRow
    Dim colIndex As Integer

    ' Reference the worksheet and table
    Set ws = ThisWorkbook.Worksheets("data_hide")
    Set lo = ws.ListObjects("Table0")
    
    ' Get the column index for the 'test_name' field
    On Error Resume Next
    colIndex = lo.ListColumns("test_name").Index
    On Error GoTo 0
    
    If colIndex = 0 Then
        MsgBox "The 'test_name' field was not found in Table0.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Clear existing items in the combo box
    Me.cbxSelectedTest.Clear

    ' Loop through each row in the table and add the 'test_name' values to the combo box
    For Each tblRow In lo.ListRows
        If tblRow.Range(1, colIndex).Value <> "" Then
            Me.cbxSelectedTest.AddItem tblRow.Range(1, colIndex).Value
        End If
    Next tblRow
End Sub