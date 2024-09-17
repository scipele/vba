Option Explicit

' filename:         PopulateHeaderFromSheetData.vba
'
' purpose:          update header information from sheet data
'
' usage:            run sub UpdateHeadersFooters()
'
' dependencies:     none
'
' By:               T.Sciple, 09/17/2024

Sub UpdateHeadersFooters()
    Dim ws As Worksheet
    Dim centerStr As String
    Dim leftStr As String
    Dim rightStr As String
    Dim i As Integer
    
    ' Get the value from cell A1 in Sheet1
    centerStr = ThisWorkbook.Sheets("title").Range("A1").Value
    
    leftStr = vbCr & vbCr   'Place two Carriage Returns at the Start of the string
    Dim left_info_start_row
    Dim left_info_last_row
    left_info_start_row = 4
    left_info_last_row = 7
    
    For i = left_info_start_row To left_info_last_row
        leftStr = leftStr & ThisWorkbook.Sheets("title").Range("a" & i).Value & ": " & ThisWorkbook.Sheets("title").Range("b" & i).Value
        If i < 7 Then
            leftStr = leftStr & vbCr
        End If
    Next i

    rightStr = vbCr    'Place one Carriage Returns at the Start of the string
    
    Dim right_info_start_row
    Dim right_info_last_row
    right_info_start_row = 3
    right_info_last_row = 7
    
    For i = right_info_start_row To right_info_last_row
        rightStr = rightStr & ThisWorkbook.Sheets("title").Range("d" & i).Value & ": " & ThisWorkbook.Sheets("title").Range("e" & i).Value
        If i < 7 Then
            rightStr = rightStr & vbCr
        End If
    Next i

    Set ws = ThisWorkbook.Sheets("test_sheet")
    
    With ws.PageSetup
        .CenterHeader = centerStr
        .LeftHeader = leftStr
        .RightHeader = rightStr
    End With
End Sub