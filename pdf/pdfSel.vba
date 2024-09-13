Sub pdfVis()
    wsName = Left(ActiveWorkbook.name, InStr(ActiveWorkbook.name, ".xls") - 1)
    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=ActiveWorkbook.Path & "\" & wsName & ".pdf"
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
End Sub


Sub pdfSel()
    Dim wsA As Worksheet
    Dim wbA As Workbook
    Dim strName As String
    Dim strPath As String
    Dim strFile As String
    Dim strPathFile As String
    On Error GoTo errHandler
    
    Set wbA = ActiveWorkbook
    Set wsA = ActiveSheet
    
    'get active workbook folder, if saved
    strPath = wbA.Path
    If strPath = "" Then
      strPath = Application.DefaultFilePath
    End If
    strPath = ActiveWorkbook.Path & "\"
    
    If InStr(ActiveWorkbook.name, "xls") = 0 Then
        MsgBox "You need to save the Workbook before printing"
        Exit Sub
    Else
        wsName = Left(ActiveWorkbook.name, InStr(ActiveWorkbook.name, ".xls") - 1) & ".pdf"
    End If
    
    strPathFile = strPath & wsName
    
    'export to PDF to current path with same name
    If myFile <> "False" Then
        wsA.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=strPathFile, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
    End If
    
    'Unselect Multiple Sheets by Selecting all sheets then only the last active one
    For Each ws In ActiveWorkbook.Sheets
        ws.Select False
    Next
    ActiveSheet.Select
    ActiveWorkbook.Save
    ActiveWorkbook.Close

exitHandler:
    Exit Sub
errHandler:
    MsgBox "Could not create PDF file"
    Resume exitHandler
End Sub