Option Explicit
Option Compare Database

'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | mImportBinaryToTable.vba                                    |
'| EntryPoint   | Sub ImportBinaryToTable                                     |
'| Purpose      | Import data from a binary file to an MS Access table        |
'| Inputs       | File path and table name                                    |
'| Outputs      | New table in MS Access database                             |
'| Dependencies | DAO library (Microsoft Office X.0 Access Database Engine)   |
'| By Name,Date | T.Sciple, 9/6/2025

Dim start_time As Double

' Enum for data types, matching the Excel VBA code
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

' UDT to store import data, matching the Excel VBA code
Private Type GeneralData
    tableName As String           ' Table name
    FilePathAndName As String     ' Input File Path and name
    FieldNames() As Variant
    DataTypeCode() As Variant
    NumExportFields As Integer
    NumExportRows As Long         ' Number of rows to import
End Type


' Example usage with FileDialog
Public Sub ImportBinaryData()

    Dim filePath As String
    
    ' Set table name
    Dim tableName As String
    tableName = "t_ImportedData"
    
    ' Create FileDialog object
    Dim fileDialog As Object
    Set fileDialog = Application.fileDialog(3) ' 3 = msoFileDialogFilePicker
    
    With fileDialog
        .Title = "Select Binary File to Import"
        .Filters.Add "Binary Files", "*.bin"
        .AllowMultiSelect = False
        If .Show = True Then
            filePath = .SelectedItems(1)
            start_time = Timer
            Call ImportBinaryFileToTable(filePath, tableName)
        End If
    End With
    
    ' Refresh navigation pane to ensure table is visible
    Application.RefreshDatabaseWindow

    
    Set fileDialog = Nothing
End Sub


' Function to create a table in Access based on binary file metadata
Private Sub CreateImportTable(gd As GeneralData)
    Dim db As DAO.Database
    Dim tbl As DAO.TableDef
    Dim fld As DAO.Field
    Dim i As Integer
    
    Set db = CurrentDb
    
    ' Check if table exists; if so, delete it (optional: could prompt user)
    On Error Resume Next
    db.TableDefs.Delete gd.tableName
    On Error GoTo 0
    
    ' Create new table
    Set tbl = db.CreateTableDef(gd.tableName)
    
    ' Add fields based on FieldNames and DataTypeCode
    With tbl
        ' Add ID field (AutoNumber, Primary Key)
        Set fld = .CreateField("ID", dbLong)
        fld.Attributes = dbAutoIncrField
        .Fields.Append fld
        
        ' Add data fields
        For i = 0 To gd.NumExportFields - 1
            Dim fieldName As String
            Dim fieldType As Integer
            
            fieldName = gd.FieldNames(i)
            Select Case gd.DataTypeCode(i)
                Case enShortText
                    fieldType = dbText
                    Set fld = .CreateField(fieldName, fieldType, 255)
                Case enLongText
                    fieldType = dbMemo
                    Set fld = .CreateField(fieldName, fieldType)
                Case enDouble
                    fieldType = dbDouble
                    Set fld = .CreateField(fieldName, fieldType)
                Case enLongInt
                    fieldType = dbLong
                    Set fld = .CreateField(fieldName, fieldType)
                Case enByte
                    fieldType = dbByte
                    Set fld = .CreateField(fieldName, fieldType)
                Case enInteger
                    fieldType = dbInteger
                    Set fld = .CreateField(fieldName, fieldType)
                Case enBoolean
                    fieldType = dbBoolean
                    Set fld = .CreateField(fieldName, fieldType)
                Case enDate
                    fieldType = dbDate
                    Set fld = .CreateField(fieldName, fieldType)
                Case enCurrency
                    fieldType = dbCurrency
                    Set fld = .CreateField(fieldName, fieldType)
                Case Else
                    MsgBox "Error: Invalid data type code " & gd.DataTypeCode(i) & " for field " & fieldName
                    Exit Sub
            End Select
            .Fields.Append fld
        Next i
    End With
    
    
    ' Add table to database
    db.TableDefs.Append tbl
    Debug.Print "Table '" & gd.tableName & "' created."
    
    
    Set fld = Nothing
    Set tbl = Nothing
    Set db = Nothing
End Sub


' Function to import binary file into Access table
Private Sub ImportBinaryFileToTable(filePath As String, tableName As String)
    Dim gd As GeneralData
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim f As Integer
    Dim i As Long, j As Integer
    Dim lenH As Long
    Dim bHeader() As Byte
    Dim bType As Byte
    Dim bData() As Byte
    Dim lenData As Long
    Dim dblValue As Double
    Dim lngValue As Long
    Dim bytValue As Byte
    Dim intValue As Integer
    Dim boolValue As Byte
    Dim curValue As Currency
    Dim strValue As String
    
    On Error GoTo ErrorHandler
    
    ' Initialize GeneralData
    gd.FilePathAndName = filePath
    gd.tableName = tableName
    
    ' Open binary file
    f = FreeFile
    Open gd.FilePathAndName For Binary Access Read As #f
    
    ' Read header info
    Get #f, , gd.NumExportRows
    Get #f, , gd.NumExportFields
    
    ' Read field names
    ReDim gd.FieldNames(0 To gd.NumExportFields - 1)
    For j = 0 To gd.NumExportFields - 1
        Get #f, , lenH
        If lenH > 0 Then
            ReDim bHeader(0 To lenH - 1)
            Get #f, , bHeader
            gd.FieldNames(j) = StrConv(bHeader, vbUnicode)
        Else
            gd.FieldNames(j) = ""
        End If
    Next j
    
    ' Read data type codes
    ReDim gd.DataTypeCode(0 To gd.NumExportFields - 1)
    For j = 0 To gd.NumExportFields - 1
        Get #f, , bType
        gd.DataTypeCode(j) = bType
    Next j
    
    ' Create table based on metadata
    Call CreateImportTable(gd)
    
    ' Open recordset for inserting data
    Set db = CurrentDb
    Set rst = db.OpenRecordset(gd.tableName, dbOpenDynaset)
    
    ' Read and insert data
    For i = 1 To gd.NumExportRows
        
        rst.AddNew
        For j = 0 To gd.NumExportFields - 1
            Select Case gd.DataTypeCode(j)
                Case enShortText, enLongText
                    Get #f, , lenData
                    If lenData > 0 Then
                        ReDim bData(0 To lenData - 1)
                        Get #f, , bData
                        strValue = StrConv(bData, vbUnicode)
                        rst.Fields(gd.FieldNames(j)).Value = strValue
                    Else
                        rst.Fields(gd.FieldNames(j)).Value = ""
                    End If
                Case enDouble
                    Get #f, , dblValue
                    If dblValue <> 0# Then
                        rst.Fields(gd.FieldNames(j)).Value = dblValue
                    End If
                Case enLongInt
                    Get #f, , lngValue
                    If lngValue <> 0 Then
                        rst.Fields(gd.FieldNames(j)).Value = lngValue
                    End If
                Case enByte
                    Get #f, , bytValue
                    If bytValue <> 0 Then
                        rst.Fields(gd.FieldNames(j)).Value = bytValue
                    End If
                Case enInteger
                    Get #f, , intValue
                    If intValue <> 0 Then
                        rst.Fields(gd.FieldNames(j)).Value = intValue
                    End If
                Case enBoolean
                    Get #f, , boolValue
                    rst.Fields(gd.FieldNames(j)).Value = (boolValue = 1)
                Case enDate
                    Get #f, , dblValue
                    If dblValue <> 0# Then
                        rst.Fields(gd.FieldNames(j)).Value = CDate(dblValue)
                    End If
                Case enCurrency
                    Get #f, , curValue
                    If curValue <> 0@ Then
                        rst.Fields(gd.FieldNames(j)).Value = curValue
                    End If
            End Select
        Next j
        rst.Update
    Next i
    
    Close #f
    rst.Close
    
    Dim elapsed_time As Double
    elapsed_time = Round(Timer - start_time, 2)
    
    MsgBox "Import completed from " & gd.FilePathAndName & " to table '" & gd.tableName & "' " & elapsed_time & " seconds."
    
    
    
    ' Clean up
    Set rst = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error importing file: " & Err.Description, vbCritical
    If Not rst Is Nothing Then rst.Close
    Close #f
    Set rst = Nothing
    Set db = Nothing
End Sub