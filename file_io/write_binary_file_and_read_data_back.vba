Option Explicit
' filenme:      write_binary_file_and_read_data_back.xlsm
'
' Purpose:      Example of how to write data to a binary file from excel
'               Example of how to read data back from the binary file and place it back in excel sheet
'
' Dependencies: Microsoft Scripting Runtime for early binding dictionary objects
'
' by T.Sciple, scipele@yahoo.com 8/3/2024
'
' Table 1 in the example is defined as the following excluding the header row to exclude 
' writing the header data to the binary file
''+-----+---------+---------+-----------+----------------------------------------------------------+
'| row | ex_integ| ex_long | ex_double | ex_text                                                  |
'+-----+---------+---------+-----------+----------------------------------------------------------+
'|  1  |   354   | 354545  | 3.1415515 | one                                                      |
'+-----+---------+---------+-----------+----------------------------------------------------------+
'|  2  |  32063  |  5454   | 560.653   | two                                                      |
'+-----+---------+---------+-----------+----------------------------------------------------------+
'|  3  | -25701  |    0    | 875.4544  | note that this example text must be less than 255        |
'|     |         |         |           | characters since the cbyte conversion of the len() is    |
'|     |         |         |           | used                                                     |
'+-----+---------+---------+-----------+----------------------------------------------------------+

Type myType
    row As Integer
    ex_integ As Integer     'Vba Integer is a 2 Byte
    ex_long As Long         'Vba Long is a 4 Byte
    ex_double As Double
    ex_text As String
End Type

'Declare a dynamic array of the User Defined Type
Sub main()
    'Dim d() as a user defined type and read table data
    Dim d() As myType
    Call ReadTableData(d())
    
    'Write the Array of User Defined Type to a Binary File
    Call WriteFromUserDefinedTypeToBinaryFile(d())
    
    'Now erase the array since the contents were written to the binary file
    Erase d
    
    'now read the data back from the binary file to the user defined type
    Call ReadFromBinaryFileBackIntoUserDefinedType(d())
    
    'Place the data read back on the sheet
    Call WriteDataFromUdtBackToSheet(d())
End Sub


Sub ReadTableData(ByRef d() As myType)
    'set the worksheet and table (change "sheet1" and "table1" to your sheet and table names)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    Set tbl = ws.ListObjects("Table1")
    
    'create a dictionary to map table headers to udt fields
    Dim fieldMap As Scripting.Dictionary
    Set fieldMap = New Scripting.Dictionary
    fieldMap.CompareMode = TextCompare 'Ignore case
    
    'get the header row
    Dim headerRow As Range
    Set headerRow = tbl.HeaderRowRange
    
    'populate the field map
    Dim colIndex As Long
    For colIndex = 1 To headerRow.Columns.Count
        Select Case headerRow.Cells(1, colIndex).Value
            Case "row"
                fieldMap.Add "row", colIndex
            Case "ex_integ"
                fieldMap.Add "ex_integ", colIndex
            Case "ex_long"
                fieldMap.Add "ex_long", colIndex
            Case "ex_double"
                fieldMap.Add "ex_double", colIndex
            Case "ex_text"
                fieldMap.Add "ex_text", colIndex
        End Select
    Next colIndex
    
    'resize the array to hold the table data
    ReDim d(1 To tbl.ListRows.Count)
    
    'loop through the table rows and populate the udt array
    Dim tblRow As ListRow
    Dim i As Integer
    i = 1
    For Each tblRow In tbl.ListRows
        With d(i)
            If fieldMap.Exists("row") Then .row = tblRow.Range(1, fieldMap("row")).Value
            If fieldMap.Exists("ex_integ") Then .ex_integ = tblRow.Range(1, fieldMap("ex_integ")).Value
            If fieldMap.Exists("ex_long") Then .ex_long = tblRow.Range(1, fieldMap("ex_long")).Value
            If fieldMap.Exists("ex_double") Then .ex_double = tblRow.Range(1, fieldMap("ex_double")).Value
            If fieldMap.Exists("ex_text") Then .ex_text = tblRow.Range(1, fieldMap("ex_text")).Value
        End With
        i = i + 1
    Next tblRow
    
End Sub


Sub WriteFromUserDefinedTypeToBinaryFile(ByRef d() As myType)
    'set the file path for the binary file
    Dim filePath As String
    filePath = "C:\t\test.bin"
    
    'delete any previous file
    If Dir(filePath) <> "" Then
        Kill filePath
        MsgBox "Deleted Previous File: " & filePath, vbInformation
    End If
    
    Dim folderPath As String
    folderPath = "c:\t"
    'Check if directory exists
    If Dir(folderPath, vbDirectory) = "" Then
	'Directory does not exist, so create it
    MkDir folderPath
    End If
    
    'write the number of rows that are written
    Dim row_cnt As Integer
    row_cnt = UBound(d) - LBound(d) + 1
    
    'open binary file for writing
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Binary As #fileNum
    
    'write the number of rows written
    Put #fileNum, , row_cnt

    'write the user defined type variables individually for each row
    Dim i As Long
    For i = LBound(d) To UBound(d)
        Put #fileNum, , d(i).row
        Put #fileNum, , d(i).ex_integ
        Put #fileNum, , d(i).ex_long
        Put #fileNum, , d(i).ex_double
        
        'write the length of the string first, then read the entire string
        Put #fileNum, , CByte(Len(d(i).ex_text))
        Put #fileNum, , d(i).ex_text
    Next i
    
    'Close the file
    Close #fileNum
    
    MsgBox ("Binary File written and closed")
End Sub


Sub ReadFromBinaryFileBackIntoUserDefinedType(ByRef d() As myType)
    Dim i As Long
    Dim fileNum As Integer
    Dim tempStr As String
    Dim strL As Byte
    
    'Define the binary file path
    Dim filePath As String
    filePath = "c:\t\test.bin"
    
    'Open the binary file
    fileNum = FreeFile
    Open filePath For Binary As fileNum
    
    'Get the number of rows written
    Dim row_cnt As Integer
    
    Get #fileNum, , row_cnt
    
    'Write the User Defined Type Variables individually for each Row
    ReDim d(1 To row_cnt)
    For i = LBound(d) To UBound(d)
        Get #fileNum, , d(i).row
        Get #fileNum, , d(i).ex_integ
        Get #fileNum, , d(i).ex_long
        Get #fileNum, , d(i).ex_double
        Call getBinaryStrLenAndReturnStringToUdt(fileNum, d(i).ex_text)
    Next i
    
    'Close the file
    Close #fileNum
    
    'Notify the user
    MsgBox "Data imported successfully."
End Sub


Sub getBinaryStrLenAndReturnStringToUdt(ByRef fileNum As Integer, ByRef udtStg As String)
        Dim strL As Byte    'Assumes one byte unsigned digit is used for the string length < 255
        Dim tempStr As String
        
        Get #fileNum, , strL
        tempStr = String$(strL, vbNullChar)    'initializes a string to given length
        Get #fileNum, , tempStr 'reads the string
        udtStg = tempStr 'sets the string to your UDT variable
End Sub


Sub WriteDataFromUdtBackToSheet(ByRef d() As myType)
    Dim ws As Worksheet
    Dim i As Integer
    Dim startRow As Integer
    Dim startCol As Integer
    
    'Set the worksheet (change "Sheet1" to your sheet name)
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    
    'Specify the starting cell for writing the data
    '********** Change as needed ***********
    startRow = 4
    startCol = 9
    
   'write header names
    ws.Cells(startRow, startCol).Value = "row"
    ws.Cells(startRow, startCol + 1).Value = "ex_integ"
    ws.Cells(startRow, startCol + 2).Value = "ex_long"
    ws.Cells(startRow, startCol + 3).Value = "ex_double"
    ws.Cells(startRow, startCol + 4).Value = "ex_text"
    
    'Write data
    For i = LBound(d) To UBound(d)
        With ws
            .Cells(startRow + i, startCol).Value = d(i).row
            .Cells(startRow + i, startCol + 1).Value = d(i).ex_integ
            .Cells(startRow + i, startCol + 2).Value = d(i).ex_long
            .Cells(startRow + i, startCol + 3).Value = d(i).ex_double
            .Cells(startRow + i, startCol + 4).Value = d(i).ex_text
        End With
    Next i
End Sub