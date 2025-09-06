Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | mExportlTblToBinary.vba                                     |
'| EntryPoint   | Sub ExportlTblToBinary                                      |
'| Purpose      | Write data from an Excel Table to a Binary File             |
'| Inputs       | Instructions Sheet Definitions                              |
'| Outputs      | Creates a binary file named per rngFilePathName             |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 9/6/2025                                          |

' - Binary format:
'   - Long: number of rows
'   - Integer: number of columns
'   - For each column: Long (field name length) + byte array (field name in ANSI)
'   - For each column: Byte (data type code from dtDataType)
'   - For each row and column:
'       enShortText: Long (length) + byte array (ANSI, truncate at 255 chars, nulls as '_')
'       enLongText: Long (length) + byte array (ANSI, no truncation, nulls as space)
'       enDouble: 8-byte Double
'       enLongInt: 4-byte Long
'       enByte: 1-byte Byte
'       enInteger: 2-byte Integer
'       enBoolean: 1-byte (0 for False, 1 for True)
'       enDate: 8-byte Double
'       enCurrency: 8-byte Currency
' - Handles large datasets in chunks.


' Define an Enum for data types with a "dt" prefix for clarity
Private Enum dtDataType
    enShortText = 0    ' Short text (e.g., up to 255 characters)
    enLongText = 1     ' Long text (e.g., memo fields or large strings)
    enDouble = 2       ' Double-precision floating-point number
    enLongInt = 3      ' Long integer
    enByte = 4         ' Byte
    enInteger = 5      ' Standard integer
    enBoolean = 6      ' True/False
    enDate = 7         ' Date/Time
    enCurrency = 8     ' Currency (fixed-point number)
End Enum

' UDT to store general export data
Private Type GeneralData
    ShtName As String             ' Worksheet name
    TableName As String           ' Table name
    FilePathAndName As String     ' Output File Path and name
    FieldNames() As Variant
    FieldTableIndex() As Variant
    DataType() As Variant
    DataTypeCode() As Variant
    NumExportFields As Integer
    NumExportRows As Long         ' Number of rows to export
End Type


Public Sub ExportTableToBinary()
    
    Dim StartTime As Double
    StartTime = Timer
    
    Dim gd As GeneralData
    Call GetGeneralData(gd)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(gd.ShtName)

    ' Explicitly check if table exists
    Dim tableExists As Boolean
    Dim lo As ListObject
    tableExists = False
    Dim tbl As ListObject
    
    For Each lo In ws.ListObjects
        If LCase(lo.Name) = LCase(gd.TableName) Then
            Set tbl = lo
            tableExists = True
            Exit For
        End If
    Next lo
    
    If Not tableExists Then
        MsgBox "Error: Table '" & gd.TableName & "' not found on the worksheet."
        Exit Sub
    End If
    
    If tbl Is Nothing Then
        MsgBox "Table '" & gd.TableName & "' could not be set."
        Exit Sub
    End If

    Dim headers As Range
    Set headers = tbl.HeaderRowRange
    
    ' Dynamically size colIndices based on gd.NumExportFields
    Dim colIndices() As Integer
    ReDim colIndices(1 To gd.NumExportFields)
    
    Dim found As Boolean
    Dim i As Integer, j As Integer

    For j = 1 To gd.NumExportFields
        found = False
        For i = 1 To headers.Columns.Count
            If headers.Cells(1, i).value = gd.FieldNames(j - 1) Then
                colIndices(j) = i
                found = True
                Exit For
            End If
        Next i
        If Not found Then
            MsgBox "Column '" & gd.FieldNames(j - 1) & "' not found in the table."
            Exit Sub
        End If
    Next j

    Dim data As Range
    Set data = tbl.DataBodyRange

    If data Is Nothing Then
        MsgBox "No data in the table."
        Exit Sub
    End If

    gd.NumExportRows = data.Rows.Count

    Call CreateSubDirectories(gd.FilePathAndName)
    
    Dim f As Integer
    f = FreeFile
    Open gd.FilePathAndName For Binary Access Write As #f

    ' Write header info
    Put #f, , gd.NumExportRows  ' Long: number of rows
    Put #f, , gd.NumExportFields  ' Integer: number of columns

    ' Write field names
    Dim bHeader() As Byte, lenH As Long
    For j = 1 To gd.NumExportFields
        bHeader = StrConv(gd.FieldNames(j - 1), vbFromUnicode)
        lenH = UBound(bHeader) + 1
        Put #f, , lenH
        If lenH > 0 Then Put #f, , bHeader
    Next j

    ' Write data type codes
    Dim bType As Byte
    For j = 1 To gd.NumExportFields
        bType = gd.DataTypeCode(j - 1)
        Put #f, , bType
    Next j

    ' Process data in chunks
    Dim chunkSize As Long: chunkSize = 5000
    Dim startRow As Long: startRow = 1
    Dim endRow As Long, chunkRows As Long
    
    ' Dynamically size colRanges based on gd.NumExportFields
    Dim colRanges() As Range
    ReDim colRanges(1 To gd.NumExportFields)
    For j = 1 To gd.NumExportFields
        Set colRanges(j) = data.Columns(colIndices(j))
    Next j
    
    ' Dynamically size chunks based on gd.NumExportFields
    Dim chunks() As Variant
    ReDim chunks(1 To gd.NumExportFields)
    
    Dim r As Long, len_ As Long
    Dim b() As Byte
    Dim value As String
    Dim dblValue As Double
    Dim lngValue As Long
    Dim bytValue As Byte
    Dim intValue As Integer
    Dim boolValue As Byte
    Dim curValue As Currency

    While startRow <= gd.NumExportRows
        endRow = Application.Min(startRow + chunkSize - 1, gd.NumExportRows)
        chunkRows = endRow - startRow + 1
        
        ' Load chunk for each column
        For j = 1 To gd.NumExportFields
            chunks(j) = colRanges(j).Rows(startRow & ":" & endRow).value
        Next j
        
        ' Write each row in the chunk
        For r = 1 To chunkRows
            For j = 1 To gd.NumExportFields
                Select Case gd.DataTypeCode(j - 1)
                    Case enShortText
                        value = CStr(chunks(j)(r, 1))
                        If value = "" Then value = "_"
                        If Len(value) > 255 Then value = Left(value, 255)
                        b = StrConv(value, vbFromUnicode)
                        len_ = UBound(b) + 1
                        Put #f, , len_
                        If len_ > 0 Then Put #f, , b
                    Case enLongText
                        value = CStr(chunks(j)(r, 1))
                        If value = "" Then value = "_"
                        b = StrConv(value, vbFromUnicode)
                        len_ = UBound(b) + 1
                        Put #f, , len_
                        If len_ > 0 Then Put #f, , b
                    Case enDouble
                        If IsEmpty(chunks(j)(r, 1)) Or Not IsNumeric(chunks(j)(r, 1)) Then
                            dblValue = 0#
                        Else
                            dblValue = CDbl(chunks(j)(r, 1))
                        End If
                        Put #f, , dblValue
                    Case enLongInt
                        If IsEmpty(chunks(j)(r, 1)) Or Not IsNumeric(chunks(j)(r, 1)) Then
                            lngValue = 0
                        Else
                            lngValue = CLng(chunks(j)(r, 1))
                        End If
                        Put #f, , lngValue
                    Case enByte
                        If IsEmpty(chunks(j)(r, 1)) Or Not IsNumeric(chunks(j)(r, 1)) Then
                            bytValue = 0
                        Else
                            bytValue = CByte(chunks(j)(r, 1))
                        End If
                        Put #f, , bytValue
                    Case enInteger
                        If IsEmpty(chunks(j)(r, 1)) Or Not IsNumeric(chunks(j)(r, 1)) Then
                            intValue = 0
                        Else
                            intValue = CInt(chunks(j)(r, 1))
                        End If
                        Put #f, , intValue
                    Case enBoolean
                        If IsEmpty(chunks(j)(r, 1)) Then
                            boolValue = 0
                        Else
                            boolValue = IIf(chunks(j)(r, 1), 1, 0)
                        End If
                        Put #f, , boolValue
                    Case enDate
                        If IsEmpty(chunks(j)(r, 1)) Or Not IsDate(chunks(j)(r, 1)) Then
                            dblValue = 0#
                        Else
                            dblValue = CDbl(chunks(j)(r, 1))
                        End If
                        Put #f, , dblValue
                    Case enCurrency
                        If IsEmpty(chunks(j)(r, 1)) Or Not IsNumeric(chunks(j)(r, 1)) Then
                            curValue = 0@
                        Else
                            curValue = CCur(chunks(j)(r, 1))
                        End If
                        Put #f, , curValue
                End Select
            Next j
        Next r
        
        startRow = endRow + 1
    Wend
    
    Close #f
    
    'Determine how many seconds code took to run
    Dim SecondsElapsed As Double
    SecondsElapsed = Round(Timer - StartTime, 2)
        
    MsgBox "Export to " & gd.FilePathAndName & " completed in " & SecondsElapsed & " Seconds"
End Sub


Private Function rng_to_ary_1d(ShtName As String, _
                              rng_str As String, _
                              base_num As Integer) _
                              As Variant
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets(ShtName).Range(rng_str)
    
    ' Dimension and resize a temporary one-dimensional array
    Dim tmp_ary As Variant
    ReDim tmp_ary(base_num To rng.Count + base_num - 1)

    ' Read the range into the array
    Dim item As Variant
    Dim i As Long
    i = base_num
    For Each item In rng
        tmp_ary(i) = item.value
        i = i + 1
    Next
    
    rng_to_ary_1d = tmp_ary
End Function


Sub GetGeneralData(ByRef gd As GeneralData)
    gd.ShtName = Range("rngSheetName")
    gd.TableName = Range("rngTableName")
    gd.FilePathAndName = Range("rngFilePathName")
    gd.FieldNames() = rng_to_ary_1d("instructions", "rngFieldNames", 0)
    gd.DataType() = rng_to_ary_1d("instructions", "rngDataTypes", 0)
    gd.NumExportFields = UBound(gd.FieldNames) - LBound(gd.FieldNames) + 1
    
    ReDim gd.FieldTableIndex(LBound(gd.FieldNames) To UBound(gd.FieldNames))
    ReDim gd.DataTypeCode(LBound(gd.DataType) To UBound(gd.DataType))
    
    Dim elem As Variant
    Dim i As Integer
    i = 0
    For Each elem In gd.DataType
        gd.DataTypeCode(i) = Switch( _
            gd.DataType(i) = "ShortText", enShortText, _
            gd.DataType(i) = "LongText", enLongText, _
            gd.DataType(i) = "Double", enDouble, _
            gd.DataType(i) = "LongInt", enLongInt, _
            gd.DataType(i) = "Byte", enByte, _
            gd.DataType(i) = "Integer", enInteger, _
            gd.DataType(i) = "Boolean", enBoolean, _
            gd.DataType(i) = "Date", enDate, _
            gd.DataType(i) = "Currency", enCurrency _
        )
        If IsEmpty(gd.DataTypeCode(i)) Then
            MsgBox "Error: Invalid data type '" & gd.DataType(i) & "' at position " & i + 1 & "."
            Exit Sub
        End If
        i = i + 1
    Next elem
End Sub


Private Sub CreateSubDirectories(ByVal fileNameAndPath As String)
    ' Purpose: Check and create subdirectories for the given file path
    Dim path As String
    Dim directories() As String
    Dim currentPath As String
    Dim i As Integer
    
    ' Extract the directory path from dg.FilePathAndName
    path = Left(fileNameAndPath, InStrRev(fileNameAndPath, "\") - 1)
    
    ' Split the path into individual directories
    directories = Split(path, "\")
    
    ' Build and check/create each directory level
    currentPath = directories(0) ' Start with drive letter (e.g., C:)
    For i = 1 To UBound(directories)
        currentPath = currentPath & "\" & directories(i)
        If Len(Dir(currentPath, vbDirectory)) = 0 Then
            MkDir currentPath
        End If
    Next i
End Sub