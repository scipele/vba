Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | headerUpdate.vba                                            |
'| EntryPoint   | Sub UpdateHeadersFooters                                    |
'| Purpose      | Read data from ThisWorkbook and write thet data to the hdr  |
'| Inputs       | Worksheet Ranges Columns A/B, D/E                           |
'| Outputs      | header data is updated                                      |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 03/04/2025                                        |

Sub UpdateHeadersFooters()
    
    Call speedup_restore(False)
    
    ' Get the value from cell A1 in Sheet1
    Dim data As String
    data = "Title"
    
    Dim centerStr As String
    centerStr = ThisWorkbook.Sheets(data).Range("A1").Value
    
    ' Get the data that will be used for the left side header and concatenate the string together
    Dim leftStr As String
    leftStr = vbCr & vbCr   'Place two Carriage Returns at the Start of the string
    Dim i As Integer
    For i = 4 To 7
        leftStr = leftStr & ThisWorkbook.Sheets(data).Range("A" & i).Value & ": " & ThisWorkbook.Sheets(data).Range("B" & i).Value
        If i < 7 Then
            leftStr = leftStr & vbCr
        End If
    Next i

    ' Get the data that will be used for the right side header and concatenate the string together
    Dim rightStr As String
    rightStr = vbCr    'Place one Carriage Returns at the Start of the string
    For i = 3 To 7
        rightStr = rightStr & ThisWorkbook.Sheets(data).Range("D" & i).Value & ": " & ThisWorkbook.Sheets(data).Range("E" & i).Value
        If i < 7 Then
            rightStr = rightStr & vbCr
        End If
    Next i

    ' Set the Worksheet Object to the sheet name that you want to update
    Dim ws_name As String
    ws_name = "TmEstim"
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(ws_name)
    
    With ws.PageSetup
        .CenterHeader = centerStr
        .LeftHeader = leftStr
        .RightHeader = rightStr
    End With
    
    Set ws = Nothing
    Call speedup_restore(True)
    
End Sub


Sub speedup_restore(ByVal at_end As Boolean)
    'Use the boolean 'at_end' to restore settings if true or make them false at the start
    Application.ScreenUpdating = at_end
    Application.DisplayStatusBar = at_end
    Application.EnableEvents = at_end
    ActiveSheet.DisplayPageBreaks = at_end
    Application.Calculation = IIf(at_end, xlAutomatic, xlManual)
End Sub
