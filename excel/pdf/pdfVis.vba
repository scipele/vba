Sub pdfVis()
    wsName = Left(ActiveWorkbook.name, InStr(ActiveWorkbook.name, ".xls") - 1)
    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=ActiveWorkbook.Path & "\" & wsName & ".pdf"
    ActiveWorkbook.Save
    ActiveWorkbook.Close
End Sub