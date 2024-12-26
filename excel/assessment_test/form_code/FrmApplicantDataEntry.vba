Option Explicit
Private m_ad As ClsApplicantData
Public Test_Closed As Boolean
'

'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | FrmApplicantDataEntry.vba                                   |
'| EntryPoint   | InitializeForm called from ModMain                          |
'| Purpose      | form code to get user name and which test they are taking   |
'| Inputs       | User Entry                                                  |
'| Outputs      | pass variables to class module 'ClsApplicantData' Setters   |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 12/26/2024                                        |


Public Sub InitializeApplicantDataEntryForm()
    
    ' Initialize the applicant object same pointer from module main
    Set m_ad = modMain.ad

    Dim test_name As Variant
    For Each test_name In m_ad.GetTestList
        Me.cbxSelectedTest.AddItem test_name
    Next test_name
    
    Me.Show
End Sub


Private Sub cmdSubmitUserData_Click()
    ' Use the following jmethod to set all the names to private member
    ' variables int the Class 'ClsApplicantData'
    m_ad.SetNames Me.tbxLastName, Me.tbxFirstName, Me.tbxMiddleName
    
    ' Exit the subroutine if fields are missing
    If m_ad.IsNameDataIncomplete Then Exit Sub
    
    ' Run setter function in the Class 'ClsApplicantData' to set the property of
    ' the selected combo box index
    m_ad.SelectedTest = Me.cbxSelectedTest.ListIndex
   
    ' Exit if not combo box is not selected and throw an error message
    If m_ad.IsSelectedTestIndxNotSet Then Exit Sub ' Exit the subroutine test is not set
    
    m_ad.IsSubmitted = True
    
    Unload Me
End Sub
