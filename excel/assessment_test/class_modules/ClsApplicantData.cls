'Class Module 'ClsApplicantData'

Option Explicit

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
Public Sub SetNames(ByVal LastName As String, ByVal FirstName As String, ByVal MiddleName As String)
    m_lastName = Trim(LastName)
    m_firstName = Trim(FirstName)
    m_middleName = Trim(MiddleName)
End Sub


'Setter 'Let' for the Test Selection
Public Property Let SelectedTest(ByVal value As Long)
    m_selectedTestIndx = value
End Property


'Getter 'Get' for the Test Selection
Public Property Get SelectedTest() As Long
    SelectedTest = m_selectedTestIndx
End Property


'Setter 'Let'
Public Property Let TestScore(ByVal value As Long)
    m_testScore = value
End Property


'Getter 'Get'
Public Property Get TestScore() As Long
    TestScore = m_testScore
End Property


'Method to get the name of the test taken
Public Function GetSelectedTestName(ByVal indx As Long)
    GetSelectedTestName = m_testNames(indx)
End Function


'Method to see if any names are incomplete
Public Function IsNameDataIncomplete() As Boolean
    If m_firstName = "" Or m_lastName = "" Or m_middleName = "" Then
        MsgBox "Please fill in all name fields.", vbExclamation, "Missing Information"
        IsNameDataIncomplete = True
    Else
        IsNameDataIncomplete = False
    End If
End Function


'Method to see if the 'SelectedTest' combo box was selected
Public Function IsSelectedTestIndxNotSet() As Boolean
    If (m_selectedTestIndx = -1) Then
        MsgBox "Please select the Test", vbExclamation, "Missing Information"
        IsSelectedTestIndxNotSet = True
    Else
        IsSelectedTestIndxNotSet = False
    End If
End Function


'Method to get this names of the Tests that are in the Excel Hidden Tab that is named Table0
Public Function GetTestList() As Variant
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim colIndex As Integer
    Dim dataRange As Range
    Dim resultArray() As Variant
    Dim i As Long

    ' Reference the worksheet and table
    Dim sht_name As String
    sht_name = "data_hide"
    Dim tbl_name As String
    tbl_name = "Table0"
    
    Set ws = ThisWorkbook.Worksheets(sht_name)
    Set lo = ws.ListObjects(tbl_name)
    
    ' Get the column index for the 'test_name' field
    On Error Resume Next
    colIndex = lo.ListColumns("test_name").index
    On Error GoTo 0
    
    If colIndex = 0 Then
        MsgBox "The 'test_name' field was not found in Table0.", vbExclamation, "Error"
        Exit Function
    End If
    
    ' Reference the data range for the 'test_name' column
    Set dataRange = lo.ListColumns(colIndex).DataBodyRange
    
    ' If the table has rows, populate the array
    If Not dataRange Is Nothing Then
        ReDim resultArray(1 To dataRange.Rows.count)
        For i = 1 To dataRange.Rows.count
            resultArray(i) = dataRange.Cells(i, 1).value
        Next i
    Else
        ' Return an empty array if no rows are present
        resultArray = Array()
    End If

    'Set Private Member Variable
    m_testNames = resultArray
    
    ' Return the result array
    GetTestList = resultArray
End Function