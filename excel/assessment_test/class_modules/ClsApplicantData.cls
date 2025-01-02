Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | ClsApplicantData.cls                                        |
'| EntryPoint   | Instatiated in Modmain                                      |
'| Purpose      | Store Job Applicant Data                                    |
'| Inputs       | Various Setters                                             |
'| Outputs      | Various Getters                                             |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 12/26/2024

Private m_lastName As String
Private m_firstName As String
Private m_middleName As String
Private m_selectedTestIndx As Long
Private m_testNames As Variant
Private m_testScore As Integer
Public IsSubmitted As Boolean


' Method to return LastFirstMiddle Names with underscore
Public Function GetLastFirstMiddleName() As String

    GetLastFirstMiddleName = m_lastName & "_" & _
                             m_firstName & "_" & _
                             m_middleName
End Function


' Method to Set All Name Properties
Public Sub SetNames(ByVal LastName As String, _
                    ByVal FirstName As String, _
                    ByVal MiddleName As String)
                    
    m_lastName = Trim(LastName)
    m_firstName = Trim(FirstName)
    m_middleName = Trim(MiddleName)
End Sub


' Setter 'Let' for the Test Selection
Public Property Let SelectedTestIndx(ByVal value As Long)
    m_selectedTestIndx = value
End Property


'Getter 'Get' for the Test Selection
Public Property Get SelectedTestIndx() As Long
    SelectedTestIndx = m_selectedTestIndx
End Property


' Setter 'Let'
Public Property Let TestScore(ByVal value As Long)
    m_testScore = value
End Property


' Getter 'Get'
Public Property Get TestScore() As Long
    TestScore = m_testScore
End Property


' Method to get the name of the test taken
Public Function GetSelectedTestName(ByVal indx As Long)
    GetSelectedTestName = m_testNames(indx)
End Function


' Method to see if any names are incomplete
Public Function IsNameDataIncomplete() As Boolean
    If m_firstName = "" Or m_lastName = "" Or m_middleName = "" Then
        MsgBox "Please fill in all name fields.", _
        vbExclamation, "Missing Information"
       
        IsNameDataIncomplete = True
    Else
        IsNameDataIncomplete = False
    End If
End Function


' Method to see if the 'SelectedTest' combo box was selected
Public Function IsSelectedTestIndxNotSet() As Boolean
    If (m_selectedTestIndx = -1) Then
        MsgBox "Please select the Test", _
        vbExclamation, "Missing Information"
        
        IsSelectedTestIndxNotSet = True
    Else
        IsSelectedTestIndxNotSet = False
    End If
End Function


' Public method to get the list of Test Names
Public Function TestNames() As Variant
    ' Set the list if not already done
    If IsEmpty(m_testNames) Then
        SetTestList
    End If
    
    TestNames = m_testNames
End Function


' Method to get this names of the Tests that are in the Excel Hidden
' Tab named Table0
Public Sub SetTestList()

    ' Reference the worksheet and table
    Dim sht_name As String
    sht_name = "data_hide"
    Dim tbl_name As String
    tbl_name = "Table0"
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sht_name)
    Dim lo As ListObject
    Set lo = ws.ListObjects(tbl_name)
    
    ' Get the column index for the 'test_name' field
    On Error Resume Next
    Dim colIndex As Integer
    colIndex = lo.ListColumns("test_name").index
    On Error GoTo 0
    
    If colIndex = 0 Then
        MsgBox "The 'test_name' field was not found in Table0.", _
        vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Reference the data range for the 'test_name' column
    Dim dataRange As Range
    Set dataRange = lo.ListColumns(colIndex).DataBodyRange
    
    ' If the table has rows, populate the array
    Dim resultArray() As Variant
    Dim indx As Long
    If Not dataRange Is Nothing Then
        ReDim resultArray(1 To dataRange.Rows.count)
        For indx = 1 To dataRange.Rows.count
            resultArray(indx) = dataRange.Cells(indx, 1).value
        Next indx
    Else
        ' Return an empty array if no rows are present
        resultArray = Array()
    End If

    ' Set Private Member Variable
    m_testNames = resultArray
End Sub