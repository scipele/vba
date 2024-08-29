' filename:  GetEstimData.vba
'
' Purpose:
'   1. This Macro helps to pull data from specific Tabs of Estimate Sheet into a common
'       MSAccess Database Table
'
' Dependencies: None
'
' By:  T.Sciple, 8/29/2024

Option Compare Database


Sub importXLDataMain()

    Dim path As String
    Dim xlFileNames As Variant
    
    Dim tmpAry As Variant
    Dim i, j As Long
    Dim ColmNoFirst, ColmNoLast As Long
    Dim ColmLtrFirst, ColmLtrLast As String
    Dim rngUse As Long
    Dim shtNames As Variant
    
    '1. Function to Get Path Where Excel Files are located
    path = getPathName()
    
    '2. Function Get an Array listing of file Names
    xlFileNames = getListOfXlsFiles(path)
        
    '3. Loop Thru all File Names that are to be imported and get data
    For i = LBound(xlFileNames, 1) To UBound(xlFileNames, 1)
        pathAndFileName = path & xlFileNames(i)
                        
        '3.1 Function to Return an Array of Sheet Names
        shtNames = getSheetNames()
        
        '3.2. Function to return the range to import as a two dimensional Array
        tmpAry = getImportSheetData(i, pathAndFileName, shtNames)
        
        '3.4 Sub that will update records in table
        Call updateTableRecords("tbl1EstimData", tmpAry)
    Next i
        
    Response = MsgBox("Importing is Done, Files are imported!", vbOKOnly, "Import Mesage")
    
    'cleanup
    Erase tmpAry
End Sub


'1. Function to Get Path Where Excel Files are located
Function getPathName()
Dim xlsPath As String
    
    'get the Path Name from the Form Text Box
    xlsPath = Forms!frmImport.tbxPath
    If Right(xlsPath, 1) <> "\" Then xlsPath = xlsPath + "\"
    getPathName = xlsPath
    
End Function


'2. Function Get an Array listing of file Names
Function getListOfXlsFiles(xlsPath)

    Dim xlsFile As String
    Dim cnt As Integer
    Dim fileNameList As Variant
    Dim i As Long
    
    'First Count the number of Items to Determine the Array Size
    tmpFileName = Dir(xlsPath & "*.xlsx", vbNormal)     ' Retrieve the first entry.
    Do While tmpFileName <> ""
        cnt = cnt + 1
        tmpFileName = Dir()
    Loop
    
    ReDim fileNameList(1 To cnt)
    'Now redo and store filenames in the array
    
    'First Count the number of Items to Determine the Array Size
    fileNameList(1) = Dir(xlsPath & "*.xlsx", vbNormal)     ' Retrieve the first entry.
        
    For i = 2 To cnt
        fileNameList(i) = Dir()
    Next i
    
    getListOfXlsFiles = fileNameList
End Function


'3.1 Utilize Function to Return an Array of Sheet Names
Private Function getSheetNames()
    Dim tmpStrAry As Variant
    
    tmpStrAry = Split("Pile,Conc,PipUG,Steel,Equip,PipShp,PipFld,Insul,Trace,FirePrf,SPaint,FPaint,EI,Bldg,Demo,SpSub,Supt,Indir", ",")
    getSheetNames = tmpStrAry
End Function


'3.2. Utilize Function to return the range to import as a two dimensional Array
Function getImportSheetData(EstimID, pathAndFileName, shtNames)
    Dim xls As Object
    Dim xlSht As Object
    Dim xlsWrkBk As Object
    Dim xlRng As Object
    Dim colmSrch As String
    Dim LastRow As Variant
    Dim tmpAry As Variant
    Dim tmpArySht As Variant
    
            
    DoCmd.SetWarnings False
    Set xls = CreateObject("Excel.Application")
    xls.Application.Visible = False
    xls.Application.displayalerts = False
    
    Set xlsWrkBk = GetObject(pathAndFileName)
    
    'Loop Thru Each Sheet
    For i = LBound(shtNames, 1) To UBound(shtNames, 1)
    Set xlSht = xlsWrkBk.Worksheets(shtNames(i)) 'NestSummary
    
        'Find the First and Last Columns to Import assuming they are within A to Z
        Set xlRng = xlSht.range("A1:Z1")
        tmpArySht = xlRng.Value
            
        For j = LBound(tmpArySht, 2) To UBound(tmpArySht, 2)
            If tmpArySht(1, j) = "X" Then  'Identification of the First Column to Get
                ColmNoFirst = j
                '3.3.1 Function to return Excel Column Letter given Numeric Designation
                ColmLtrFirst = colmNoLetter(j)
            End If
            If tmpArySht(1, j) = "TOTAL" Then  'Identification of the Last Column to Get
                ColmNoLast = j
                '3.3.1 Function to return Excel Column Letter given Numeric Designation
                ColmLtrLast = colmNoLetter(j)
            End If
        Next j
            
        'Find the Last Row to Import
        'Get Look for Keyword in this column
            colmSrch = colmNoLetter(ColmNoFirst + 4)
            Set xlRng = xlSht.range(colmSrch & "1:" & colmSrch & "1000")
            tmpArySht = xlRng.Value
            
            For j = LBound(tmpArySht, 1) To UBound(tmpArySht, 1)
                If tmpArySht(j, 1) = "TOTALS" Then 'Identification of the First Column to Get
                    LastRow = j - 1
                End If
            Next j
            
        'Now Reset the Range to match the cell data that you want to retrive
        Set xlRng = xlSht.range(ColmLtrFirst & "2:" & ColmLtrLast & LastRow)
        tmpArySht = xlRng.Value
    
    
        '3.2.1  Function to Delete Data that is Zero Value or Does not Contain "X"
        If IsArray(tmpArySht) = True Then
            tmpArySht = deleteUnpopulatedData(EstimID, i + 1, tmpArySht)
        End If
        
    
        '3.2.2 Function to Merge Data from different Arrays respectively
        If IsArray(tmpAry) = False And IsArray(tmpArySht) = True Then
            tmpAry = tmpArySht
        Else
            If IsArray(tmpAry) = True And IsArray(tmpArySht) = True Then
                tmpAry = merge2DArrayData(tmpAry, tmpArySht)
            End If
        End If
       
    Next i
    
    DoCmd.SetWarnings True
    'close excel
    xlsWrkBk.Application.Quit
    Set xls = Nothing
    
    getImportSheetData = tmpAry
End Function


'3.2.2 Function to Merge Data from different Arrays respectively
Private Function merge2DArrayData(tmpAryA, tmpAryB)
    Dim tmpAryC As Variant
    Dim SzC As Long
    Dim i, ib, j As Long
    
    Dim lb1a, ub1a, lb1b, ub1b, lb2a, ub2a, lb2b, ub2b As Long
    
    If IsArray(tmpAryA) = False Then GoTo errLabel2
    If IsArray(tmpAryB) = False Then GoTo errLabel2
    
    
    lb1a = LBound(tmpAryA, 1)
    ub1a = UBound(tmpAryA, 1)
    lb1b = LBound(tmpAryB, 1)
    ub1b = UBound(tmpAryB, 1)
    lb2a = LBound(tmpAryA, 2)
    ub2a = UBound(tmpAryA, 2)
    lb2b = LBound(tmpAryB, 2)
    ub2b = UBound(tmpAryB, 2)
    
    'Make sure that Array Second Dimensions match otherwise exit Function
    If (ub2a - lb2a) <> (ub2b - lb2b) Then GoTo errLabel1
    
    'Determine required Size of Merged Array
    SzC = (ub1a - lb1a) + (ub1b - lb1b) + 1
    
    ReDim tmpAryC(lb1a To (lb1a + SzC), lb2a To ub2a)
    
    'First Nested Loop to append 1st Array to the new array
    For i = lb1a To ub1a
        For j = lb2a To ub2a
            tmpAryC(i, j) = tmpAryA(i, j)
        Next j
    Next i
    
    ib = 0 'counter that will be used for second array
    'Second Nested Loop to append 1st Array to the new array
    For i = (ub1a + 1) To (lb1a + SzC)
        ib = ib + 1
        For j = lb2a To ub2a
            tmpAryC(i, j) = tmpAryB(ib, j)
        Next j
    Next i
    
    merge2DArrayData = tmpAryC
    Exit Function
    
errLabel1:
        MsgBox "Error in merge2DArrayData Function caused because Array Sizes Did not match"
        Exit Function
errLabel2:
        MsgBox "Error in merge2DArrayData Function caused because empty Array"
        Exit Function
End Function


'3.3  Function to Delete Data that is Zero Value or Does not Contain "X"
Private Function deleteUnpopulatedData(EstimID, TabID, tmpAry1)
    Dim i, j, k, m As Long
    Dim tmpAry2 As Variant
    Dim Flag As Boolean
    
    Flag = False
    'First Determine the Count of Applicable lines to keep for Array Sizing Purpose
    m = 0
    For i = LBound(tmpAry1, 1) To UBound(tmpAry1, 1)
        Flag = False 'reset Flag for each i
        If tmpAry1(i, 13) <> 0 Then Flag = True  'Get all record Data if the Value is non Zero
        If tmpAry1(i, 1) = "X" And tmpAry1(i, 4) <> "" Then Flag = True 'Get other Titles that are Marked X if Desc is Not Blank
            If Flag = True Then
                m = m + 1
            End If
    Next i
    
    j = 0
    
    If m > 0 Then
        ReDim tmpAry2(1 To m + 1, 1 To 16)
        
        For i = LBound(tmpAry1, 1) To UBound(tmpAry1, 1)
            Flag = False 'reset Flag for each i
            If tmpAry1(i, 13) <> 0 Then Flag = True  'Get all record Data if the Value is non Zero
            If tmpAry1(i, 1) = "X" And tmpAry1(i, 4) <> "" Then Flag = True 'Get other Titles that are Marked X if Desc is Not Blank
            
            If Flag = True Then
                j = j + 1
                'Record EstimNo, TabNo, and Row Number for Reference
                tmpAry2(j, 1) = EstimID
                tmpAry2(j, 2) = TabID
                tmpAry2(j, 3) = i + 1
                    
                For k = 1 To 13
                    tmpAry2(j, k + 3) = tmpAry1(i, k)
                Next k
            End If
        Next i
    Else: Exit Function
    
    End If
    
    deleteUnpopulatedData = tmpAry2
End Function


'3.4 Sub that will update records in table
Sub updateTableRecords(tblName As String, aryData As Variant)
    Dim rs As DAO.Recordset
    Dim i, j As Long
    Dim fldNameAry, fldNameType As Variant
    Dim temp As Variant
    
    'open table
    'Listing of all Fields to be updated
    fldNameAry = Split("EstimID,TabID,Row,X,Qty,UOM,Desc,unitClientMatl,TotalClientMatl,unitContrMatl,TotalContrMatl,UnitMh,TotalMh,LabRate,TotalLab,Total", ",")
    fldNameType = Split("lg,lg,lg,txt,dbl,txt,txt,dbl,dbl,dbl,dbl,dbl,dbl,dbl,dbl,dbl", ",")
    
    'Update all fields in the Array
    Set rs = CurrentDb.OpenRecordset("tbl1EstimData")
    
    For i = LBound(aryData, 1) To UBound(aryData, 1)
        rs.AddNew
        For j = LBound(fldNameAry, 1) To UBound(fldNameAry, 1)
            
            Select Case fldNameType(j)
                Case "lg"
                    If VarType(aryData(i, j + 1)) = vbLong Then
                        rs.Fields(fldNameAry(j)) = aryData(i, j + 1)
                    Else
                        rs.Fields(fldNameAry(j)) = Null
                    End If
                    
                Case "txt"
                    If VarType(aryData(i, j + 1)) = vbString Then
                        rs.Fields(fldNameAry(j)) = aryData(i, j + 1)
                    Else
                        rs.Fields(fldNameAry(j)) = ""
                    End If
            
                Case "dbl"
                    If IsNumeric(aryData(i, j + 1)) = True Then
                        rs.Fields(fldNameAry(j)) = aryData(i, j + 1)
                    Else
                        rs.Fields(fldNameAry(j)) = Null
                    End If
            End Select
        Next j
        rs.Update
    Next i
End Sub


'3.3.1 Function to return Excel Column Letter given Numeric Designation
Private Function colmNoLetter(colmNo As Variant)
    Dim Ltr1, Ltr2 As String
    Dim AlphaChars As Variant
    
    AlphaChars = Split("A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z", ",")
    
    If colmNo < 27 Then colmNoLetter = AlphaChars(colmNo - 1)
    
    If colmNo > 26 And colmNo < 703 Then
        Ltr1 = AlphaChars(Int(-1 * colmNo / 26) / -1 - 2)
        Ltr2 = AlphaChars(26 * ((colmNo / 26) - Int((colmNo / 26))) - 1)
        colmNoLetter = Ltr1 & Ltr2
    End If
End Function
