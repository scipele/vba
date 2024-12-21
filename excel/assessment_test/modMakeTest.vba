Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | modMakeTest.vba                                             |
'| EntryPoint   | main                                                        |
'| Purpose      | create an assessment test by reading questions/answers and  |
'|              | recording the user                                          |
'| Inputs       | Table Data in data_hide, User Input                         |
'| Outputs      | selected test answers written to hidden table               |
'| Dependencies | Indicate if any libraries are used or none                  |
'| By Name,Date | T.Sciple, 12/21/2024                                        |

Public ud As userData
Public td() As TestData

Public isDataEntered As Boolean

Public Type userData
    last_first_middle_name As String
    selected_test_id As Integer
    test_name As String
    test_score As Integer
End Type

Public Type TestData
    quest_id As Integer
    question_text As String
    answer_1 As String
    answer_2 As String
    answer_3 As String
    answer_4 As String
    selected_answer As Integer
    correct_answer As Integer
    is_correct As Integer
End Type


Public Sub main()
    '1. Get User General Data
    Call frmUserDataEntry.InitializeForm
    
    ' Check if the form was completed
    If isDataEntered = False Then
        MsgBox "Form was closed without entering data. Program will terminate.", vbExclamation, "Operation Cancelled"
        Exit Sub
    End If
    
    '2. Read the question and answer data from the table into the user defined type above
    Call ReadTableData(td(), ud.selected_test_id)
    
    '2.  Show the user Form
    Call ShowUserQAForm(td(), ud)
    
    'cleanup
    Erase td()
End Sub


'getters for this module
Public Sub PassUserDataToMod(ByVal tmpName As String, _
                             ByVal tmpTest As Long, _
                             ByVal tmpTestName As String)
                             
    ud.last_first_middle_name = tmpName
    ud.selected_test_id = tmpTest
    ud.test_name = tmpTestName
End Sub


Private Sub ReadTableData(ByRef td() As TestData, ByVal tableIndex As Integer)
    
    ' Construct the table name based on the index
    Dim tblName As String
    tblName = "Table" & tableIndex

    ' Check if the table exists
    On Error Resume Next
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("data_hide").ListObjects(tblName)
    On Error GoTo 0

    If tbl Is Nothing Then
        MsgBox "Table " & tblName & " does not exist!", vbExclamation
        Exit Sub
    End If

    ' Resize the array to hold the table data
    ReDim td(1 To tbl.ListRows.Count)

    ' Loop through the rows of the table
    Dim i As Integer
    i = 1
    Dim tblRow As ListRow
    For Each tblRow In tbl.ListRows
        With td(i)
            .quest_id = tbl.DataBodyRange.Cells(i, tbl.ListColumns("quest_id").Index).Value
            .question_text = tbl.DataBodyRange.Cells(i, tbl.ListColumns("question_text").Index).Value
            .answer_1 = tbl.DataBodyRange.Cells(i, tbl.ListColumns("answer_1").Index).Value
            .answer_2 = tbl.DataBodyRange.Cells(i, tbl.ListColumns("answer_2").Index).Value
            .answer_3 = tbl.DataBodyRange.Cells(i, tbl.ListColumns("answer_3").Index).Value
            .answer_4 = tbl.DataBodyRange.Cells(i, tbl.ListColumns("answer_4").Index).Value
            .correct_answer = tbl.DataBodyRange.Cells(i, tbl.ListColumns("correct_answer").Index).Value
        End With
        i = i + 1
    Next tblRow

End Sub



Public Sub ShowUserQAForm(ByRef td() As TestData, _
                          ByRef ud As userData)
    Dim formInstance As frmTestQA
    Set formInstance = New frmTestQA
    
    formInstance.LoadQuestions td, ud
    formInstance.Show
End Sub


Public Function SubmitAnswers(ByRef tdFrm() As TestData, _
                              ByRef ud As userData) _
                              As Boolean
    
    Dim i As Integer
    Dim allAnswered As Boolean
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tblRow As ListRow
    Dim num_correct_answers As Integer
    num_correct_answers = 0

    ' Ensure all questions are answered
    allAnswered = True
    For i = 1 To UBound(tdFrm)
        If tdFrm(i).selected_answer = 0 Then
            allAnswered = False
            Exit For
        End If
    Next i

    If Not allAnswered Then
        SubmitAnswers = False
        Exit Function
    End If

    ' Save answers to the Excel table
    Set ws = ThisWorkbook.Sheets("data_hide") ' Replace with your actual sheet name
    Set tbl = ws.ListObjects("Table" & ud.selected_test_id)  ' Adjust table naming as needed
    For i = 1 To UBound(tdFrm)
        ' Find the corresponding row in the table using the quest_id
        For Each tblRow In tbl.ListRows
            If tblRow.Range(tbl.ListColumns("quest_id").Index).Value = tdFrm(i).quest_id Then
                ' Update the selected_answer column
                tblRow.Range(tbl.ListColumns("selected_answer").Index).Value = tdFrm(i).selected_answer
                
                ' Check to see if the answer is correct or not
                
                If tdFrm(i).selected_answer = tdFrm(i).correct_answer Then
                    tdFrm(i).is_correct = 1
                    tblRow.Range(tbl.ListColumns("is_correct").Index).Value = 1
                    num_correct_answers = num_correct_answers + 1
                Else
                    tdFrm(i).is_correct = 0
                    tblRow.Range(tbl.ListColumns("is_correct").Index).Value = 0
                End If
                    
                Exit For
            End If
        Next tblRow
    Next i

    ' If everything is successful, return True
    SubmitAnswers = True
    
    ' Calculate the grade
    ud.test_score = num_correct_answers / UBound(tdFrm) * 100
    MsgBox ("Your Test Score was " & ud.test_score & " Percent")
    
    
    Call SaveTestResultsToCSV(tdFrm(), ud)
    
    'cleanup
    Erase tdFrm
    
    
End Function


Private Sub SaveTestResultsToCSV(ByRef tdFrm() As TestData, _
                                 ByRef ud As userData)

    Dim testFileName As String
    testFileName = ThisWorkbook.Path & _
                   "\test_results_" & _
                   ud.last_first_middle_name & _
                   ".csv"
    
    ' Open the CSV file for appending data (or create it if it doesn't exist)
    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open testFileName For Output As fileNumber
    
    Dim i As Integer
    Dim delimeter As String
    Dim selected_answers As String
    Dim correct_answers As String
    
    For i = 1 To UBound(tdFrm())
        delimeter = IIf(i = UBound(tdFrm()), "", ",")
        selected_answers = selected_answers & tdFrm(i).selected_answer & delimeter
        correct_answers = correct_answers & tdFrm(i).correct_answer & delimeter
    Next i
    
    'Indicate answers for record purposes
    
    ' Write the test taker's information as a new row in the CSV
    Print #fileNumber, "Last_First_Middle_Name: " & _
                        ud.last_first_middle_name & vbCr & _
                        "Test Taken: " & _
                        ud.test_name & vbCr & _
                        "Test Score: " & _
                        ud.test_score & vbCr & _
                        "Selected_Answers_Numeric: " & _
                        selected_answers & vbCr & _
                        "Correct_Answers_Numeric : " & _
                        correct_answers & vbCr & ""
    
    ' Close the file
    Close fileNumber

End Sub