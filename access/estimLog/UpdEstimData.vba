Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Option Compare Database


Public Sub rename_estim_sht(ByRef ed As est_data)
    'rename estimate filename
    Dim name_orig As String
    name_orig = ed.folderName & "4. Estim\R0\1. Overall\24-XXXX Estim_R0.xlsm"
    ed.estimFileName = ed.folderName & "4. Estim\R0\1. Overall\" & ed.estimateNo & " Estim_R0.xlsm"
    
    'Pause before and after renaming the file
    Sleep 1500  'Pause for 1.5 seconds
    Name name_orig As ed.estimFileName
    Sleep 1500  'Pause for 1.5 seconds
End Sub


Public Sub prep_data_then_update(ByRef ed As est_data)
    
    ' Get the selected ID from the control and then the related name using Dlookup
    ed.estimatorNameID = Forms("frm03EstimData").Controls("EstimatorMech").Value
    ed.estimatorName = DLookup("LastName", "tlkpEstimator", "ID = " & ed.estimatorNameID)
    
    If IsNull(Forms("frm03EstimData").Controls("BidDate")) Then
        ed.BidDate = ""
    Else
        ed.BidDate = Forms("frm03EstimData").Controls("BidDate")
    End If
    
    ed.jobLoc = Right(ed.curClient, Len(ed.curClient) - Len(ed.shortClient) - 2)
    
    Call upd_estim(ed)

End Sub


Private Sub upd_estim(ByRef ed As est_data)
    ' Instantiate new Excel instance
    Dim xlApp As Object
    Set xlApp = CreateObject("Excel.Application")
    
    ' Open the Excel file
    Dim xlWbk As Object
    Set xlWbk = xlApp.Workbooks.Open(FileName:=ed.estimFileName, UpdateLinks:=False, ReadOnly:=False)
    
    ' Make it visable
    xlApp.Visible = True
    
    Dim xlSht As Object
    Set xlSht = xlWbk.Sheets("Title")
   
    ' Update the specified cells with data
    xlSht.Range("B4").Value = ed.estimateNo
    xlSht.Range("B5").Value = ed.titleOrig
    xlSht.Range("B6").Value = ed.shortClient
    xlSht.Range("B7").Value = ed.jobLoc
    xlSht.Range("E4").Value = Date
    xlSht.Range("E5").Value = ed.BidDate
    xlSht.Range("E6").Value = ed.estimatorName
    
    ' Run the UpdateHeadersFooters Sub located in the workbook
    xlApp.Run "'" & xlWbk.Name & "'!UpdateHeadersFooters"
    
    xlWbk.Save
    xlWbk.Close False
    xlApp.Quit
    
    ' Release the object references
    Set xlSht = Nothing
    Set xlWbk = Nothing
    Set xlApp = Nothing
End Sub