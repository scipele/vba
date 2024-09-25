' SubName:  pdfSel
'
' Purpose:  Create a pdf of the selected sheets of the active workbook
'           PDF is created with exact filename as the excel file
'
' By:       JAS, 9/24/2024

Sub pdfSel()
    
    Dim wbA As Workbook
    Set wbA = ActiveWorkbook
    
    Dim wsA As Worksheet
    Set wsA = ActiveSheet
    
    'get active workbook folder, if saved
    Dim strPath As String
    strPath = wbA.Path
    If strPath = "" Then
      strPath = Application.DefaultFilePath
    End If
    strPath = ActiveWorkbook.Path & "\"
        
    'set error handler
    On Error GoTo errHandler
    
    'Check to see if the user has default "book1" open
    Dim wsName As String
    If InStr(ActiveWorkbook.name, "xls") = 0 Then
        MsgBox "You need to save the Workbook before printing"
        Exit Sub
    Else
        wsName = Left(ActiveWorkbook.name, InStr(ActiveWorkbook.name, ".") - 1) & ".pdf"
    End If
    
    Dim strPathFile As String
    strPathFile = strPath & wsName
    
    'export to PDF to current path with same name
    Dim myFile As Variant
    If myFile <> "False" Then
        wsA.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=strPathFile, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
    End If
    
    'The following removes the multiple selected tabs by selecting the first sheet at index (1)
    wbA.Sheets(1).Select
    
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
exitHandler:
    Exit Sub
errHandler:
    MsgBox "Could not create PDF file"
    Resume exitHandler
End Sub