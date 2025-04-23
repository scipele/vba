Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | ClsTestData.cls                                             |
'| EntryPoint   | Instantiated in Modmain                                     |
'| Purpose      | Store Test Data                                             |
'| Inputs       | Various Setters                                             |
'| Outputs      | Various Getters                                             |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 12/26/2024

Private m_numQuestions As Long
Private m_ad As ClsApplicantData
Private m_qd() As ClsQuestionData
'

Public Function GetQuestionData(ByVal index As Integer) As ClsQuestionData
    If index >= LBound(m_qd) And index <= UBound(m_qd) Then
        Set GetQuestionData = m_qd(index)
    Else
        Err.Raise vbObjectError + 513, "ClsTestData", "Index out of bounds"
    End If
End Function


'Getter Method to return number of questions
Public Property Get NumQuestions() As Long
    NumQuestions = m_numQuestions
End Property


Public Sub ReadTableData(ByVal SelectedTestIndx As Long)
    ' Construct the table name based on the zero based index
    Dim tblName As String
    tblName = "Table" & SelectedTestIndx + 1

    ' Check if the table exists
    On Error Resume Next
    Dim sht_name As String
    sht_name = "data_hide"
    
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets(sht_name).ListObjects(tblName)
    On Error GoTo 0

    If tbl Is Nothing Then
        MsgBox "Table " & tblName & " does not exist!", vbExclamation
        Exit Sub
    End If

    ' set the private member variable based on the number of questions
    m_numQuestions = tbl.ListRows.count

    ' Resize the array of
    ReDim m_qd(1 To tbl.ListRows.count)

    ' Loop through the rows of the table
    Dim indx As Integer
    indx = 1
    Dim tblRow As ListRow
    For Each tblRow In tbl.ListRows
    
        ' Instantiate the 'ClsQuestionData' object for each index 'indx'
        Set m_qd(indx) = New ClsQuestionData
        
        m_qd(indx).QuestionId = _
            tbl.DataBodyRange.Cells(indx, tbl.ListColumns("quest_id").index).value
        m_qd(indx).QuestionText = _
            tbl.DataBodyRange.Cells(indx, tbl.ListColumns("question_text").index).value
        m_qd(indx).PotentialAnswer(1) = _
            tbl.DataBodyRange.Cells(indx, tbl.ListColumns("answer_1").index).value
        m_qd(indx).PotentialAnswer(2) = _
            tbl.DataBodyRange.Cells(indx, tbl.ListColumns("answer_2").index).value
        m_qd(indx).PotentialAnswer(3) = _
            tbl.DataBodyRange.Cells(indx, tbl.ListColumns("answer_3").index).value
        m_qd(indx).PotentialAnswer(4) = _
            tbl.DataBodyRange.Cells(indx, tbl.ListColumns("answer_4").index).value
        m_qd(indx).CorrectAnswer = _
            tbl.DataBodyRange.Cells(indx, tbl.ListColumns("correct_answer").index).value
        indx = indx + 1
    Next tblRow
End Sub


Public Function SubmitAnswers() As Boolean
    Dim i As Integer
    Dim allAnswered As Boolean
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tblRow As ListRow
    Dim num_correct_answers As Integer
    num_correct_answers = 0

    ' Ensure all questions are answered
    allAnswered = True
    For i = 1 To m_numQuestions
        If m_qd(i).SelectedAnswer = 0 Then
            allAnswered = False
            Exit For
        End If
    Next i

    If Not allAnswered Then
        SubmitAnswers = False
        Exit Function
    End If

    ' Set private object reference to be the same as the original instance
    Set m_ad = modMain.ad

    ' Save answers to the Excel table
    Dim sheet_name As String
    sheet_name = "data_hide"
    
    Set ws = ThisWorkbook.Sheets(sheet_name)
    Set tbl = ws.ListObjects("Table" & m_ad.SelectedTestIndx + 1)
    For i = 1 To m_numQuestions
        ' Find the corresponding row in the table using the quest_id
        For Each tblRow In tbl.ListRows
            If tblRow.Range(tbl.ListColumns("quest_id").index).value = _
                m_qd(i).QuestionId Then
                
                ' Update the selected_answer column
                tblRow.Range(tbl.ListColumns("selected_answer").index).value = _
                    m_qd(i).SelectedAnswer
                
                ' Check to see if the answer is correct or not
                If m_qd(i).SelectedAnswer = m_qd(i).CorrectAnswer Then
                    m_qd(i).IsCorrect = 1
                    tblRow.Range(tbl.ListColumns("is_correct").index).value = 1
                    num_correct_answers = num_correct_answers + 1
                Else
                    m_qd(i).IsCorrect = 0
                    tblRow.Range(tbl.ListColumns("is_correct").index).value = 0
                End If
                    
                Exit For
            End If
        Next tblRow
    Next i

    ' If everything is successful, return True
    SubmitAnswers = True
    
    ' Calculate the grade
    m_ad.TestScore = num_correct_answers / m_numQuestions * 100
    MsgBox ("Your Test Score was " & m_ad.TestScore & " Percent")
    
    Call SaveTestResultsToCSV
    
End Function


Private Sub SaveTestResultsToCSV()
    Dim testFileName As String
    testFileName = ThisWorkbook.Path & _
                   "\test_results_" & _
                   m_ad.GetLastFirstMiddleName & "_" & _
                   m_ad.SelectedTestIndx & _
                   ".csv"
    
    ' Open the CSV file for appending data (or create it if it doesn't exist)
    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open testFileName For Output As fileNumber
    
    Dim i As Integer
    Dim delimeter As String
    Dim selected_answers As String
    Dim correct_answers As String
    
    
    'Indicate answers for record purposes
    
    ' Write the test taker's information as a new row in the CSV
    Print #fileNumber, "Last_First_Middle_Name: " & _
                            m_ad.GetLastFirstMiddleName & vbCr & _
                        "Test Taken: " & _
                            m_ad.GetSelectedTestName(m_ad.SelectedTestIndx + 1) & vbCr & _
                        "Test Score: " & _
                            m_ad.TestScore & vbCr & _
                        "+---------+----------+------------------------------------------------------------------------------------------------------------|" & vbCr & _
                        "| Correct | Selected | Questions                                                                                                  |" & vbCr & _
                        "| Answer  | Answer   |    Answers                                                                                                 |"


    ' dim variables use in the loops below
    Dim suffix As String
    Dim j As Integer
    Dim is_correct_x As String
    Dim is_selected_x As String
    Dim current_answer_ascii_letter As String
    Dim pad_end As String
        
    ' Print Questions to the results file
    For i = 1 To m_numQuestions
        suffix = IIf(i < 10, ".  ", ". ")
        Print #fileNumber, "+---------+----------+------------------------------------------------------------------------------------------------------------|" & vbCr & _
                            "|         |          | " & i & suffix & m_qd(i).QuestionText & Space(103 - Len(m_qd(i).QuestionText)) & "|"
        
        'Print the Answers and notations for correct and selected answers
        For j = 1 To 4
            is_selected_x = IIf(j = m_qd(i).SelectedAnswer, "X", " ")
            is_correct_x = IIf(j = m_qd(i).CorrectAnswer, "X", " ")
            current_answer_ascii_letter = Chr(96 + j)
            
            If (m_qd(i).PotentialAnswer(j) <> "") Then
                Print #fileNumber, "|   " & _
                    is_correct_x & _
                    "     |     " & _
                    is_selected_x & _
                    "    |    " & _
                    current_answer_ascii_letter & _
                    ". " & m_qd(i).PotentialAnswer(j) & _
                    Space(101 - Len(m_qd(i).PotentialAnswer(j))) & "|"
            End If
        Next j
    Next i
    Print #fileNumber, "+---------+----------+------------------------------------------------------------------------------------------------------------|"

    ' Close the file
    Close fileNumber
End Sub