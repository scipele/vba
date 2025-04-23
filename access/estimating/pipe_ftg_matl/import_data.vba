' filename:  import_pvf_pricing_data.vba
'
' Purpose:
'   1. imports data from excel workbooks into an MSAccess Database
'
' Dependencies: None
'
' By:  T.Sciple, 8/29/2024

Option Compare Database


Sub importXLDataMain()

    Dim tmpDataAry As Variant   'This array will be stored from each sheet from which the data is retrieved
    Dim pathAndFileName As String
    Dim fileData As Variant
    Dim searchParameters, refParameters As Variant
    Dim i As Long
    Dim keyColmNo As Integer
    Dim StartTime, SecondsElapsed As Double

    'Remember time when macro starts
      StartTime = Timer
        
    '1.0 Get file list with path information
        fileData = getTableData("tbl1_file_list")
    
    '2.0 Search Parameters
        'That represent LeftColmHdr, RightColmHdr,  BtmRowLessOne
        searchParameters = Split("Line Item|Line Total|SUBTOTAL", "|")
    
    '3.0 Get Active Sheet Data and place it into the table
        
        For i = LBound(fileData, 1) To UBound(fileData, 1)
            pathAndFileName = fileData(i, 2) & fileData(i, 1)
            
            '3.1 Get Active Sheet Data
                tmpDataAry = getSheetData(pathAndFileName, searchParameters)
            
            '3.2 Setup an Array for only the Column Headings that match the data that is being imported
                If Not IsEmpty(tmpDataAry) = True Then
                    captureHeadings = Split("Line Item|Qty.|UOM|DESCRIPTION|Code|Unit Price|Line Total", "|")
                    tmpDataAry = cleanupData(i, captureHeadings, tmpDataAry)
                Else
                    GoTo nextIteration
                End If
                
            '3.3 Remove Array Elements that do not have a value in the'6th' column over
                keyColmNo = 5
                tmpDataAry = removeEmptyData(keyColmNo, tmpDataAry)
            
            '3.3 Place the data into a table
                'Sub updateTableRecords(tblName As String, aryData As Variant)
                    Call updateTableRecords(i, "tbl2_data", tmpDataAry)
nextIteration:
        Next i
    
        'Determine how many seconds code took to run
        SecondsElapsed = Round(Timer - StartTime, 2)
        
        Response = MsgBox("Importing is Done, Files are imported in " & SecondsElapsed & "Seconds", vbOKOnly, "Import Mesage")

        'cleanup
        Erase tmpDataAry
       
End Sub


'1.0 Get file list with path information
Private Function getTableData(tblName)
    Dim numRecords, i As Long
    Dim rs As Object
    Set rs = CurrentDb.OpenRecordset(tblName)
    Dim tempAry As Variant
    numRecords = rs.RecordCount
    
    'Change Dimensions of the Array as needed
    ReDim tempAry(1 To numRecords, 1 To 2)
    
    For i = 1 To numRecords
        tempAry(i, 1) = rs!file_name
        tempAry(i, 2) = rs!file_path
        rs.MoveNext
    Next i

    getTableData = tempAry
    Erase tempAry

End Function


'3.0 Get Active Sheet Data and place it into the table
Function getSheetData(path_file As String, srchAry)

    Dim xls As Object
    Dim xlSht As Object
    Dim xlsWrkBk As Object
    Dim xlRng As Object
    Dim tmpArySht As Variant
    Dim foundCell As Object
    Dim startRow, lastRow As Long
    Dim startColm, endColm As String
            
    DoCmd.SetWarnings False
    Set xls = CreateObject("Excel.Application")
    xls.Application.Visible = False
    xls.Application.displayalerts = False
    
    On Error GoTo beforeEndFunction
    Set xlsWrkBk = GetObject(path_file)

    'Copy Active Sheet Only
    Set xlSht = xlsWrkBk.activesheet 'Sets the sheet as whatever sheet was active when it opens

    'Find the First and Last Columns to Import assuming they are within A to Z and 1000 lines
    Set xlRng = xlSht.Range("A1:Z1000")
    
    'Search for the key parameter text within the range specified above
    For i = 0 To 2
        Set foundCell = xlRng.Find(srchAry(i))
        
        ' Check if the text was found and then set key range parameters
        If Not foundCell Is Nothing Then
            Select Case i
                Case 0
                    startColm = Split(foundCell.Address, "$")(1)
                    startRow = foundCell.Row
                Case 1
                    endColm = Split(foundCell.Address, "$")(1)
                Case 2
                    lastRow = foundCell.Row - 2
            End Select
        
        Else
            Exit Function
        End If
        
    Next i
    
    'Now Set the Range to match the cell data that you want to retrive
    Set xlRng = xlSht.Range(startColm & startRow & ":" & endColm & lastRow)
    tmpArySht = xlRng.Value

    DoCmd.SetWarnings True
    
    'close excel
    xlsWrkBk.Application.Quit
    Set xls = Nothing
    
    getSheetData = tmpArySht

beforeEndFunction:
End Function


'3.2 Setup an Array for only the Column Headings that match the data that is being imported
    Function cleanupData(fileCounter, captureHeadings, tmpDataAry)
    Dim i, j As Long
    Dim tmpAryClean As Variant
    Dim matchColmNoAry, num As Variant
    
    'First run through the Columns in the tmpData Array and record the Column number where it matches the captureHeadings Array
    ReDim matchColmNoAry(LBound(captureHeadings, 1) To UBound(captureHeadings, 1))
    
    For i = LBound(captureHeadings, 1) To UBound(captureHeadings, 1)
        For j = LBound(tmpDataAry, 2) To UBound(tmpDataAry, 2)
            If captureHeadings(i) = tmpDataAry(1, j) Then
                matchColmNoAry(i) = j
            End If
        Next j
    Next i
    
    'Now create the clean array now that we have the matching column numbers in the matchColmNoAry
    
    ReDim tmpAryClean(LBound(tmpDataAry, 1) To UBound(tmpDataAry, 1), LBound(captureHeadings, 1) To UBound(captureHeadings, 1))
    
    'skip the first row with '+1' in i loop
    For i = LBound(tmpDataAry, 1) + 1 To UBound(tmpDataAry, 1)
        For j = LBound(captureHeadings, 1) To UBound(captureHeadings, 1)
            If Not IsEmpty(matchColmNoAry(j)) Then
                tmpAryClean(i, j) = tmpDataAry(i, matchColmNoAry(j))
            End If
        Next j
    Next i

    cleanupData = tmpAryClean
End Function


'3.3 removeEmptyData(keyColmNo, tmpDataAry)
Function removeEmptyData(keyColmNo, tmpDataAry)

    Dim newDataAry As Variant

    'count the nonblank lines using the 'keyColmNo' passed column number
    k = LBound(tmpDataAry, 1)
    For i = LBound(tmpDataAry, 1) To UBound(tmpDataAry, 1)
        If Not IsEmpty(tmpDataAry(i, keyColmNo)) = True Then
            k = k + 1
        End If
    Next i

    If k <= 1 Then k = 2  'reset k so that you at least have one line
    ReDim newDataAry(LBound(tmpDataAry, 1) To k - 1, LBound(tmpDataAry, 2) To UBound(tmpDataAry, 2))
    
    k = LBound(tmpDataAry, 1)
    For i = LBound(tmpDataAry, 1) To UBound(tmpDataAry, 1)
        If Not IsEmpty(tmpDataAry(i, keyColmNo)) = True Then
            For j = LBound(tmpDataAry, 2) To UBound(tmpDataAry, 2)
                    newDataAry(k, j) = tmpDataAry(i, j)
            Next j
            k = k + 1
        End If
    Next i

    removeEmptyData = newDataAry

End Function


'4.0 Sub that will update records in table
Sub updateTableRecords(fileIdLg As Long, tblName As String, aryData As Variant)
    Dim rs As DAO.Recordset
    Dim i, j As Long
    Dim fldNameAry As Variant
    Dim temp As Variant
    Dim pad As String
    
    'open table
    'Listing of all Fields to be updated
    
    'Set number of fields in the data table
    numFields = 13
    
    'resize the array according to the number of fields
    ReDim fldNameAry(1 To numFields)
    
    For i = 1 To numFields
        If i < 10 Then pad = "0" Else pad = ""
        fldNameAry(i) = "f" & pad & Trim(Str(i))
    Next i
    
    'Update all fields in the Array
    Set rs = CurrentDb.OpenRecordset(tblName)
    
    
    For i = LBound(aryData, 1) To UBound(aryData, 1)
        rs.AddNew
        rs.Fields("fileID") = fileIdLg

        For j = LBound(aryData, 2) To UBound(aryData, 2)
                rs.Fields(fldNameAry(j + 1)) = aryData(i, j)
        Next j
        rs.Update
    Next i
End Sub