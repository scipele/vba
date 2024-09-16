Option Explicit
' filename:         FilePathFolderPicker.vba
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
' By:               T.Sciple, 09/16/2024
    
Sub selectFile()
    'Create and set dialog box as variabble
    Dim dialogBox As FileDialog
    Set dialogBox = Application.FileDialog(msoFileDialogOpen)
    
    'do not allow multiple files to be selected
    dialogBox.AllowMultiSelect = False
    
    'set the title of the Dialog Box
    dialogBox.Title = "Select a File"
    
    'Set the defaultFolder to Open
    'not used
    
    'clear the dialog box filters
    dialogBox.Filters.Clear
    
    'apply file filters - use ; to separate filters for the same name
    'dialogBox.Filters.Add "Excel Workbooks", "*.xlsl;*.xls;*.xlsm"
    
    'show the fialog box and output the fill file name
    If dialogBox.Show = -1 Then
        ActiveSheet.Range("filepath").Value = dialogBox.SelectedItems(1)
    End If
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