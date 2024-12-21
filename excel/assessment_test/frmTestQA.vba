Private currentQuestionIndex As Integer
Private tdFrm() As TestData
Private ud_frm As userData

'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | frmTestQA.vba                                               |
'| EntryPoint   | LoadQuestions (called from module code)                     |
'| Purpose      | form code for displaying qusetions, answers, w/ option btns |
'| Inputs       | Getters set global                          |
'| Outputs      | selected test answers written to hidden table               |
'| Dependencies | Indicate if any libraries are used or none                  |
'| By Name,Date | T.Sciple, 12/20/2024                                        |


Public Sub LoadQuestions(ByRef td() As TestData, _
                         ByRef ud As userData)
    
    'Sets the passed user defined typers to module global variables so there are accessible
    tdFrm = td
    ud_frm = ud
    currentQuestionIndex = 1
    DisplayQuestion
End Sub


Private Sub DisplayQuestion()
    ' Display the question and restore the saved answer
    If currentQuestionIndex > 0 And currentQuestionIndex <= UBound(tdFrm) Then
        With tdFrm(currentQuestionIndex)
            lblQuestion.Caption = .question_text
                lblAnswer1.Caption = .answer_1
            lblAnswer2.Caption = .answer_2
            
            If (.answer_3 <> "") Then
                lblAnswer3.Visible = True
                optAnswer3.Visible = True
                lblAnswer3.Caption = .answer_3
            Else
                lblAnswer3.Visible = False
                optAnswer3.Visible = False
            End If
            
            If (.answer_4 <> "") Then
                lblAnswer4.Visible = True
                optAnswer4.Visible = True
                lblAnswer4.Caption = .answer_4
            Else
                lblAnswer4.Visible = False
                optAnswer4.Visible = False
            End If
            
            ' Restore the previously selected answer
            Select Case .selected_answer
                Case 1: optAnswer1.Value = True
                Case 2: optAnswer2.Value = True
                Case 3: optAnswer3.Value = True
                Case 4: optAnswer4.Value = True
                Case Else
                    ' Clear any previous selection
                    optAnswer1.Value = False
                    optAnswer2.Value = False
                    optAnswer3.Value = False
                    optAnswer4.Value = False
            End Select
        End With
    End If
    UpdateNavigationButtons
End Sub


Private Sub UpdateNavigationButtons()
    cmdPrevious.Enabled = currentQuestionIndex > 1
    cmdNext.Enabled = currentQuestionIndex < UBound(tdFrm)
End Sub


Private Sub cmdNext_Click()
    If currentQuestionIndex > 0 And currentQuestionIndex <= UBound(tdFrm) Then
        If tdFrm(currentQuestionIndex).selected_answer = 0 Then
            MsgBox "Please select an answer before proceeding.", vbExclamation
            Exit Sub
        End If
    End If
    
    If currentQuestionIndex < UBound(tdFrm) Then
        currentQuestionIndex = currentQuestionIndex + 1
        DisplayQuestion
    End If
End Sub


Private Sub cmdPrevious_Click()
    If currentQuestionIndex > 1 Then
        currentQuestionIndex = currentQuestionIndex - 1
        DisplayQuestion
    End If
End Sub


Private Sub cmdSubmit_Click()
    ' Call the module's SubmitAnswers procedure
    Dim success As Boolean
    success = SubmitAnswers(tdFrm, ud_frm)
    
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


Private Sub UpdateSelectedAnswer(selectedAnswer As Integer)
    If currentQuestionIndex > 0 And currentQuestionIndex <= UBound(tdFrm) Then
        tdFrm(currentQuestionIndex).selected_answer = selectedAnswer
    End If
End Sub

