Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | ClsQuestionData.cls                                         |
'| EntryPoint   | Each index is instantiated in the 'ClsTestData'             |
'|              | .ReadTableData method for each row of table data in the     |
'|              | spreadhseet                                                 |
'| Purpose      | Store Job Applicant Data                                    |
'| Inputs       | Various Setters                                             |
'| Outputs      | Various Getters                                             |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 12/26/2024

Private m_quest_id As Integer
Private m_question_text As String
Private m_potentialAnswer(1 To 4) As String
Private m_selectedAnswer As Integer
Private m_correctAnswer As Integer
Private m_isCorrect As Integer
'

' Public Setter or 'Let' class function to set private member variables
Public Property Let QuestionId(ByVal value As Integer)
    m_quest_id = value
End Property


' Public Setter or 'Let' class function to set private member variables
Public Property Let QuestionText(ByVal value As String)
    m_question_text = value
End Property


' Setter for a specific potential answer by index
Public Property Let PotentialAnswer(ByVal index As Integer, _
                                    ByVal value As String)
                                    
    If index >= LBound(m_potentialAnswer) _
    And index <= UBound(m_potentialAnswer) Then
        m_potentialAnswer(index) = value
    Else
        Err.Raise vbObjectError + 1, "ClsQuestionData", _
        "Index out of bounds for potential answers."
    End If
End Property


' Public Setter or 'Let' class function to set private member variables
Public Property Let SelectedAnswer(ByVal value As Integer)
    m_selectedAnswer = value
End Property


' Public Setter or 'Let' class function to set private member variables
Public Property Let CorrectAnswer(ByVal value As Integer)
    m_correctAnswer = value
End Property


' Public Setter or 'Let' class function to set private member variables
Public Property Let IsCorrect(ByVal value As Integer)
    m_isCorrect = value
End Property


' Getter Function to access the private member variables
Public Property Get QuestionId() As Integer
    QuestionId = m_quest_id
End Property


' Getter Function to access the private member variables
Public Property Get QuestionText() As String
    QuestionText = m_question_text
End Property


' Getter Function to access the private member variables
Public Property Get PotentialAnswer(ByVal index As Integer) As String
    PotentialAnswer = m_potentialAnswer(index)
End Property


' Getter Function to access the private member variables
Public Property Get SelectedAnswer() As Integer
    SelectedAnswer = m_selectedAnswer
End Property


' Getter Function to access the private member variables
Public Property Get CorrectAnswer() As Integer
    CorrectAnswer = m_correctAnswer
End Property


' Getter Function to access the private member variables
Public Property Get IsCorrect() As Integer
    IsCorrect = m_isCorrect
End Property