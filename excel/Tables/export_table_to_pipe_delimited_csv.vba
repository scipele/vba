Option Explicit

' Export Excel Table to Pipe-Delimited CSV
' Description: Exports all data from an Excel table to a pipe-delimited CSV file
' Usage: Run ExportTableToPipeDelimitedCSV macro

Sub ExportTableToPipeDelimitedCSV()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim filePath As String
    Dim fileNum As Integer
    Dim rowIdx As Long
    Dim colIdx As Long
    Dim cellValue As String
    Dim pipeLine As String
    Dim lastRow As Long
    Dim lastCol As Long
    
    ' Prompt user to select a table if multiple exist
    On Error Resume Next
    Set ws = ActiveSheet
    
    ' Check if worksheet has tables
    If ws.ListObjects.Count = 0 Then
        MsgBox "No tables found on the active sheet. Please select a sheet with a table.", vbExclamation
        Exit Sub
    End If
    
    ' Use first table if only one exists, otherwise let user select
    If ws.ListObjects.Count = 1 Then
        Set tbl = ws.ListObjects(1)
    Else
        Dim tableNames As String
        Dim i As Long
        tableNames = "Available tables:" & vbCrLf
        For i = 1 To ws.ListObjects.Count
            tableNames = tableNames & i & ". " & ws.ListObjects(i).Name & vbCrLf
        Next i
        MsgBox tableNames & vbCrLf & "Using the first table.", vbInformation
        Set tbl = ws.ListObjects(1)
    End If
    
    ' Get file path from user
    filePath = Application.GetSaveAsFilename( _
        fileFilter:="CSV Files (*.csv), *.csv", _
        Title:="Save Pipe-Delimited CSV As")
    
    If filePath = "False" Then
        MsgBox "Export cancelled.", vbInformation
        Exit Sub
    End If
    
    ' Get free file number
    fileNum = FreeFile
    
    ' Open file for output
    On Error GoTo ErrorHandler
    Open filePath For Output As fileNum
    
    ' Export header row
    pipeLine = ""
    For colIdx = 1 To tbl.ListColumns.Count
        cellValue = tbl.ListColumns(colIdx).Name
        cellValue = EscapeSpecialChars(cellValue)
        If colIdx = 1 Then
            pipeLine = cellValue
        Else
            pipeLine = pipeLine & "|" & cellValue
        End If
    Next colIdx
    Print #fileNum, pipeLine
    
    ' Export data rows
    lastRow = tbl.ListRows.Count
    
    For rowIdx = 1 To lastRow
        pipeLine = ""
        For colIdx = 1 To tbl.ListColumns.Count
            cellValue = CStr(tbl.DataBodyRange.Cells(rowIdx, colIdx).Value)
            cellValue = EscapeSpecialChars(cellValue)
            If colIdx = 1 Then
                pipeLine = cellValue
            Else
                pipeLine = pipeLine & "|" & cellValue
            End If
        Next colIdx
        Print #fileNum, pipeLine
    Next rowIdx
    
    Close fileNum
    
    MsgBox "Export completed successfully!" & vbCrLf & _
           "File saved to: " & filePath & vbCrLf & _
           "Rows exported: " & lastRow + 1 & " (including header)", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    On Error Resume Next
    Close fileNum
End Sub

' Helper function to escape special characters in CSV values
Private Function EscapeSpecialChars(cellValue As String) As String
    ' Replace pipe characters with alternative representation
    ' Replace line breaks with spaces
    cellValue = Replace(cellValue, "|", "~")
    cellValue = Replace(cellValue, vbCrLf, " ")
    cellValue = Replace(cellValue, vbCr, " ")
    cellValue = Replace(cellValue, vbLf, " ")
    ' Remove non-printable characters
    cellValue = RemoveNonPrintableChars(cellValue)
    EscapeSpecialChars = cellValue
End Function

' Helper function to remove non-printable characters
Private Function RemoveNonPrintableChars(inputStr As String) As String
    Dim result As String
    Dim i As Long
    Dim charCode As Long
    
    result = ""
    For i = 1 To Len(inputStr)
        charCode = Asc(Mid(inputStr, i, 1))
        ' Keep printable ASCII characters (32-126) and common extended ASCII (128-255)
        ' Allow tabs (9), spaces (32-126), and extended ASCII (128-255)
        If (charCode >= 32 And charCode <= 126) Or charCode >= 128 Or charCode = 9 Then
            result = result & Mid(inputStr, i, 1)
        End If
    Next i
    
    RemoveNonPrintableChars = result
End Function

' Alternative: Export from specific range (not table)
Sub ExportRangeToPipeDelimitedCSV()
    Dim selectedRange As Range
    Dim filePath As String
    Dim fileNum As Integer
    Dim rowIdx As Long
    Dim colIdx As Long
    Dim cellValue As String
    Dim pipeLine As String
    
    ' Prompt user to select range
    On Error Resume Next
    Set selectedRange = Application.InputBox( _
        "Select the range to export (including headers):", _
        Type:=8)
    
    If selectedRange Is Nothing Then
        MsgBox "No range selected. Export cancelled.", vbInformation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Get file path from user
    filePath = Application.GetSaveAsFilename( _
        fileFilter:="CSV Files (*.csv), *.csv", _
        Title:="Save Pipe-Delimited CSV As")
    
    If filePath = "False" Then
        MsgBox "Export cancelled.", vbInformation
        Exit Sub
    End If
    
    ' Get free file number
    fileNum = FreeFile
    
    ' Open file for output
    On Error GoTo ErrorHandler
    Open filePath For Output As fileNum
    
    ' Export selected range
    For rowIdx = 1 To selectedRange.Rows.Count
        pipeLine = ""
        For colIdx = 1 To selectedRange.Columns.Count
            cellValue = CStr(selectedRange.Cells(rowIdx, colIdx).Value)
            cellValue = EscapeSpecialChars(cellValue)
            If colIdx = 1 Then
                pipeLine = cellValue
            Else
                pipeLine = pipeLine & "|" & cellValue
            End If
        Next colIdx
        Print #fileNum, pipeLine
    Next rowIdx
    
    Close fileNum
    
    MsgBox "Export completed successfully!" & vbCrLf & _
           "File saved to: " & filePath & vbCrLf & _
           "Rows exported: " & selectedRange.Rows.Count, vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    On Error Resume Next
    Close fileNum
End Sub
