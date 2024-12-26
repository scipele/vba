'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | frmTestManager.vba                                          |
'| EntryPoint   | InitializeTestDataForm (called from module code)            |
'| Purpose      | form code for displaying questions, answers, w/ option btns |
'| Inputs       | User input with combo buttons                               |
'| Outputs      | selected test answers written to 'ClsTestData' members      |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 12/26/2024                                        |


Option Explicit
Private m_ad As ClsApplicantData
Private m_td As ClsTestData
Dim currentQuestion As ClsQuestionData

Private questionIndx As Integer


Public Sub InitializeTestDataForm()
    
    ' Initialize the applicant object
    Set m_ad = modMain.ad
    Set m_td = modMain.td
    
    ' Method to gather test questions and potential answers from the spreadsheet
    Dim selected_test As Long
    
    selected_test = m_ad.SelectedTest
    
    ' Call the following method to read the test questions amd potential answers using
    ' the Class 'ClsTestData' and
    ' its refeenced Class 'ClsQuestionData'
    
    m_td.ReadTableData m_ad.SelectedTest
    questionIndx = 1
    
    Call DisplayQuestionAndPotentialAnswers
    Me.Show
    
End Sub


Private Sub DisplayQuestionAndPotentialAnswers()

    ' Display the question and restore the saved answer
    If questionIndx > 0 And questionIndx <= m_td.NumQuestions Then

        Set currentQuestion = m_td.GetQuestionData(questionIndx)
        
        'Update the label captions
        lblQuestion.Caption = currentQuestion.QuestionText
        lblAnswer1.Caption = currentQuestion.PotentialAnswer(1)
        lblAnswer2.Caption = currentQuestion.PotentialAnswer(2)

        If (currentQuestion.PotentialAnswer(3) <> "") Then
            lblAnswer3.Visible = True
            optAnswer3.Visible = True
            lblAnswer3.Caption = currentQuestion.PotentialAnswer(3)
        Else
            lblAnswer3.Visible = False
            optAnswer3.Visible = False
        End If

        If (currentQuestion.PotentialAnswer(4) <> "") Then
            lblAnswer4.Visible = True
            optAnswer4.Visible = True
            lblAnswer4.Caption = currentQuestion.PotentialAnswer(4)
        Else
            lblAnswer4.Visible = False
            optAnswer4.Visible = False
        End If

        ' Restore the previously selected answer
        Select Case currentQuestion.SelectedAnswer
            Case 1: optAnswer1.value = True
            Case 2: optAnswer2.value = True
            Case 3: optAnswer3.value = True
            Case 4: optAnswer4.value = True
            Case Else
                ' Clear any previous selection
                optAnswer1.value = False
                optAnswer2.value = False
                optAnswer3.value = False
                optAnswer4.value = False
        End Select
    End If
    UpdateNavigationButtons
End Sub


Private Sub UpdateNavigationButtons()
    cmdPrevious.Enabled = questionIndx > 1
    cmdNext.Enabled = questionIndx < m_td.NumQuestions

End Sub


Private Sub cmdNext_Click()
    If questionIndx > 0 And questionIndx <= m_td.NumQuestions Then
        If currentQuestion.SelectedAnswer = 0 Then
            MsgBox "Please select an answer before proceeding.", vbExclamation
            Exit Sub
        End If
    End If
    
    If questionIndx < m_td.NumQuestions Then
        questionIndx = questionIndx + 1
        DisplayQuestionAndPotentialAnswers
    End If
End Sub


Private Sub cmdPrevious_Click()
    If questionIndx > 1 Then
        questionIndx = questionIndx - 1
        DisplayQuestionAndPotentialAnswers
    End If
End Sub


Private Sub cmdSubmit_Click()
    ' Call the module's SubmitAnswers procedure
    Dim success As Boolean
    success = m_td.SubmitAnswers()
    
    If success Then
        MsgBox "All answers have been submitted successfully!", vbInformation, "Submission Complete"
        Unload Me
    Else
        MsgBox "Please answer all questions before submitting.", vbExclamation, "Incomplete Answers"
    End If

End Sub


Private Sub optAnswer1_Click()
    UpdateSelectedAnswer 1
End Sub


Private Sub optAnswer2_Click()
    UpdateSelectedAnswer 2
End Sub


Private Sub optAnswer3_Click()
    UpdateSelectedAnswer 3
End Sub


Private Sub optAnswer4_Click()
    UpdateSelectedAnswer 4
End Sub


Private Sub UpdateSelectedAnswer(SelectedAnswer As Integer)
    If questionIndx > 0 And questionIndx <= m_td.NumQuestions Then
        currentQuestion.SelectedAnswer = SelectedAnswer
    End If
    
    If questionIndx = m_td.NumQuestions Then
        MsgBox ("Note that this is the last question:" & vbCr & _
                "Make sure that you answered all questions and Press Submit Completed Answers")
    End If
    
End Sub