Option Explicit

' SubName:   pdfVis
'
' Purpose:  Create a pdf of the visable sheets in the active workbook
'           PDF is created with exact filename as the excel file
'
' By:       JAS, 9/24/2024

Sub pdfVis()
    Dim wsName As String
    wsName = Left(ActiveWorkbook.name, InStr(ActiveWorkbook.name, ".") - 1)
    
    ActiveWorkbook.ExportAsFixedFormat _
                   Type:=xlTypePDF, _
                   Filename:=ActiveWorkbook.Path & "\" & wsName & ".pdf"
    
    ActiveWorkbook.Save
    ActiveWorkbook.Close
End Sub