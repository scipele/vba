Option Explicit
Private m_ad As ClsApplicantData
'
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | FrmApplicantDataEntry.vba                                   |
'| EntryPoint   | InitializeApplicantDataEntryForm called from ModMain        |
'| Purpose      | form code to get user name and which test they are taking   |
'| Inputs       | test_name read from method 'm_ad.GetTestList'               |
'| Outputs      | pass variables to class module 'ClsApplicantData' Setters   |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 12/26/2024                                        |


Public Sub InitializeApplicantDataEntryForm()
    ' Initialize 'ClsApplicantData' with same pointer from 'ModMain'
    Set m_ad = modMain.ad

    ' Use the GetTestList Method to get array of the Test Names
    Dim test_names As Variant
    test_names = m_ad.TestNames
    
    ' Load the combobox with the test names
    Dim test_name As Variant
    For Each test_name In test_names
        Me.cbxSelectedTest.AddItem test_name
    Next test_name
    
    ' Display the form
    Me.Show
End Sub


Private Sub cmdSubmitUserData_Click()
    ' Use the following jmethod to set all the names to private member
    ' variables int the Class 'ClsApplicantData'
    m_ad.SetNames Me.tbxLastName, Me.tbxFirstName, Me.tbxMiddleName
    
    ' Exit the subroutine if fields are missing
    If m_ad.IsNameDataIncomplete Then Exit Sub
    
    ' Run setter function in the Class 'ClsApplicantData' to set the
    ' property of the selected combo box index
    m_ad.SelectedTestIndx = Me.cbxSelectedTest.ListIndex
   
    ' Exit if not combo box is not selected and throw an error message
    If m_ad.IsSelectedTestIndxNotSet Then Exit Sub
    
    ' Set flag to indicated whether the applicant data was submitted
    m_ad.IsSubmitted = True
    
    ' Unload the form or 'Me' object
    Unload Me
End Sub
