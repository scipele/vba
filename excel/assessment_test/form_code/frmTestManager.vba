'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | frmTestManager.vba                                          |
'| EntryPoint   | InitializeTestDataForm (called from module code)            |
'| Purpose      | form code for displaying questions, answers, w/ option btns |
'| Inputs       | User input with combo buttons                               |
'| Outputs      | selected test answers written to 'ClsTestData' members      |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 2/23/2025 v05 (shuffled questions)                |


Option Explicit
Private m_ad As ClsApplicantData
Private m_td As ClsTestData
Dim currentQuestion As ClsQuestionData
Private questionIndx As Integer
Dim shuffledIndxAry As Variant


Public Sub InitializeTestDataForm()
    ' Initialize the applicant object
    Set m_ad = modMain.ad
    Set m_td = modMain.td
    
    ' Method to get test questions and potential answers from the spreadsheet
    Dim selected_test As Long
    selected_test = m_ad.SelectedTestIndx
    
    ' Call the following method to read the test questions amd potential
    ' answers using the Class 'ClsTestData' and  its referenced
    ' Class 'ClsQuestionData'
    
    m_td.ReadTableData m_ad.SelectedTestIndx
    questionIndx = 1
    
    shuffledIndxAry = getShuffledAry(m_td.NumQuestions)
    
    'Sub to set the label values and visibility
    DisplayQuestionAndPotentialAnswers
    
    ' Display the form
    Me.Show
End Sub


Private Sub DisplayQuestionAndPotentialAnswers()
    ' Display the question and restore the saved answer
    If questionIndx > 0 And questionIndx <= m_td.NumQuestions Then

        ' Set the current question object based on the current index.
        ' Note that the new keyword is not used because the private member
        ' data is already set with the method 'm_td.ReadTableData'
        Dim current_shuffled_question_id As Integer
        current_shuffled_question_id = shuffledIndxAry(questionIndx)
        
        Set currentQuestion = m_td.GetQuestionData(current_shuffled_question_id)
        
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
            MsgBox "Please select an answer before proceeding.", _
            vbExclamation
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
        MsgBox "All answers have been submitted successfully!", _
            vbInformation, "Submission Complete"
        Unload Me
    Else
        MsgBox "Please answer all questions before submitting.", _
            vbExclamation, "Incomplete Answers"
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
                "Make sure that you answered all questions and " & _
                "Press Submit Completed Answers")
    End If
End Sub


Private Function getShuffledAry(ByVal max_indx As Integer) As Variant

    ' Generate a sequential array from 0 to max_indx
    Dim tmp_ary As Variant
    ReDim tmp_ary(1 To max_indx)
    Dim i As Integer

    For i = 1 To max_indx
        tmp_ary(i) = i
    Next i

    ' Fisher-Yates Shuffle Algorithm
    Randomize
    Dim j As Integer
    Dim temp As Integer
    
    ' Loop backward from Last element to second item
    For i = max_indx - 1 To 2 Step -1 ' dont need to run for last element
        j = Int(i * Rnd) + 1            ' Generate random index from 1 to i
        
        ' Swap the random elements
        temp = tmp_ary(i)
        tmp_ary(i) = tmp_ary(j)
        tmp_ary(j) = temp
    Next i

    getShuffledAry = tmp_ary
End Function