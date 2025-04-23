Option Explicit
' filename:         fileFullPathNameFolderPicker.vba
'
' purpose:          template for file and folder pickers
'
' usage:            run Sub selectFile() or
'                   Sub folderPicker()
'
' dependencies:     OLE Automation
'                   Microsoft Office 16.0 Object Library
'                   ctv OLE Control module
'                   Microsoft Forms 2.0 Object Library
'                   Ref Edit Control
'                   Microsoft Windows Common Controls 6.0 (SP6)
'
' By:               T.Sciple, 2/24/2025
    
    
Sub selectFile()
    'Create and set dialog box as variabble
    Dim dialogBox As FileDialog
    Set dialogBox = Application.FileDialog(msoFileDialogOpen)
    
    'do not allow multiple files to be selected
    dialogBox.AllowMultiSelect = False
    
    'set the title of the Dialog Box
    dialogBox.Title = "Select a File"
    
    'Set default directory to the location of the Excel file containing the macro
    Dim defaultPath As String
    defaultPath = ThisWorkbook.Path
    
    'Ensure there is a valid path (not empty, which could happen with unsaved workbooks)
    If defaultPath <> "" Then
        dialogBox.InitialFileName = defaultPath & "\"
    End If
    
    'Set the defaultFolder to Open
    'not used
    'clear the dialog box filters
    dialogBox.Filters.Clear
    
    'apply file filters - use ; to separate filters for the same name
    dialogBox.Filters.Add "Excel Workbooks", "*.csv"
    
    'show the fialog box and output the fill file name
    Dim fileFullPathName As String
    
    If dialogBox.Show = -1 Then
        ActiveSheet.Range("fileFullPathName").Value = dialogBox.SelectedItems(1)
        fileFullPathName = dialogBox.SelectedItems(1)
    End If
    
    Call ClearAllCellsAndTables
    Call FastImportTextFile(fileFullPathName)
    Call formatSheet
End Sub


Sub folderPicker()
    'List files and files in Subfolders
     Dim strPath As String _
     , strFileSpec As String _
     , booIncludeSubfolders As Boolean

     strPath = SelectFolder() & "\"

     'If user clicks cancel then the strPath will only be "\" then exit the Sub
     If strPath = "\" Then Exit Sub

     strFileSpec = "*.*"
     booIncludeSubfolders = True

     ActiveSheet.Range("folderpath").Value = strPath
     
     
End Sub


Public Function SelectFolder()   'or Sub SelectFolder()
    Dim Fd As FileDialog
    Set Fd = Application.FileDialog(msoFileDialogFolderPicker)
    With Fd
        .AllowMultiSelect = False
        .Title = "Please select folder"
        If .Show = True Then      'if OK is pressed
             SelectFolder = .SelectedItems(1)
        Else
           Exit Function  'click Cancel or X box to close the dialog
        End If
    End With
    Set Fd = Nothing
End Function


Sub FastImportTextFile(ByVal fileFullPathName As String)
    Dim fileNumber As Integer
    Dim lineContent As String
    Dim rowNumber As Long
    Dim colData As Variant
    Dim ws As Worksheet
    Dim dataArray() As Variant
    Dim rowCounter As Long
    Dim maxColumns As Integer
    Dim tempArr() As String
    
    fileNumber = FreeFile
    
    ' Set worksheet
    Set ws = ThisWorkbook.Sheets("results") ' Change to desired sheet
    
    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Open the file
    Open fileFullPathName For Input As #fileNumber
    
    ' Read file line by line and store in an array
    rowCounter = 0
    maxColumns = 0
    
    ' First, count rows to predefine array size
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, lineContent
        rowCounter = rowCounter + 1
        tempArr = Split(lineContent, "|")
        If UBound(tempArr) > maxColumns Then maxColumns = UBound(tempArr)
    Loop
    
    ' Reset file pointer to start
    Close #fileNumber
    Open fileFullPathName For Input As #fileNumber
    
    ' Redefine the array with correct size
    ReDim dataArray(1 To rowCounter, 1 To maxColumns + 1)
    
    rowNumber = 1
    ' Read again, filling the array
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, lineContent
        colData = Split(lineContent, "|")
        
        ' Store in array
        Dim colNumber As Integer
        For colNumber = LBound(colData) To UBound(colData)
            dataArray(rowNumber, colNumber + 1) = Trim(colData(colNumber))
        Next colNumber
        
        rowNumber = rowNumber + 1
    Loop
    
    ' Close the file
    Close #fileNumber
    
    ' Write all data to the worksheet in one go
    ws.Range(ws.Cells(1, 1), ws.Cells(rowCounter, maxColumns + 1)).Value = dataArray
    
    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Import completed successfully!", vbInformation
End Sub


Sub formatSheet()

    Sheets("results").Activate
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$4:$C$165"), , xlYes).Name = _
        "Table1"
    Range("Table1[#All]").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("F8").Select
    Columns("C:C").ColumnWidth = 100
    Cells.Select
    Range("Table1[[#Headers],[Correct Answer]]").Activate
    Cells.EntireRow.AutoFit
    ActiveWindow.SmallScroll Down:=-36
    Range("B5").Select
    ActiveWindow.FreezePanes = True
End Sub


Sub ClearAllCellsAndTables()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("results") ' Change to your sheet name
    
    ' Delete all tables (ListObjects)
    Dim lo As ListObject
    Do While ws.ListObjects.Count > 0
        ws.ListObjects(1).delete
    Loop
    
    ' Clear all cells
    ws.Cells.Clear
End Sub
