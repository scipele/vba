'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | modMain.vba                                                 |
'| EntryPoint   | StartTest from excel sheet active-X cmd Button              |
'| Purpose      | create an assessment test by reading questions/answers and  |
'|              | recording the user answers                                  |
'| Inputs       | Excel table data in data_hide, User Input                   |
'| Outputs      | selected test answers written to hidden table               |
'| Dependencies | Indicate if any libraries are used or none                  |
'| By Name,Date | T.Sciple, 12/26/2024                                        |


Public ad As ClsApplicantData
Public td As ClsTestData

Public Sub StartTest()

    ' Note that a single Public instances is created for each of the objects
    ' here so that subsequent form module code and class modules can refer
    ' back to these instances to avoid losing the data
    Set ad = New ClsApplicantData
    
    ' Initiatiate the Form for applicant data entry and allow the applicant
    ' to submit their data
    FrmApplicantDataEntry.InitializeApplicantDataEntryForm
    
    'Check if the window was closed
    If Not ad.IsSubmitted Then
        MsgBox "Form was closed without submitting data.", _
            vbExclamation, "Operation Cancelled"
        Exit Sub
    End If
    
    ' a. Initialize the Form for Test Data which will call related class
    '    to get the test data, display the question, and potential answers.
    ' b. Navigation buttons will allow the applicant to navigate to the
    '    next/previous next questions and submit final answers when they
    '    have completed the test.
    Set td = New ClsTestData
    FrmTestManager.InitializeTestDataForm
End Sub
