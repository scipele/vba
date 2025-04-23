Option Explicit
' filename:         DeleteEmptyFolder.vba
'
' purpose:          recursively delete empty folders given a starting path from
'                   a named range 'startPath'
'
' usage:            run Sub DeleteEmptyFolder()
'
' dependencies:     none
'
' By:               T.Sciple, 09/16/2024

Private oFSO As Object

Sub DeleteEmptyFolder()
    Dim oRootFDR As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oRootFDR = oFSO.GetFolder(Range("startPath").Value) '<--- Change to your root folder
 
    If DeleteEmptyFolderOnly(oRootFDR) Then
        oRootFDR.Delete
    End If
    Set oRootFDR = Nothing
    Set oFSO = Nothing
End Sub


Private Function DeleteEmptyFolderOnly(ByRef oFDR As Object) As Boolean
    Dim bDeleteFolder As Boolean, oSubFDR As Object
    bDeleteFolder = False
    ' Recurse into SubFolders
    For Each oSubFDR In oFDR.SubFolders
        If DeleteEmptyFolderOnly(oSubFDR) Then
            Debug.Print "Delete", oSubFDR.Path ' Comment for production use
            oSubFDR.Delete
        End If
    Next
    ' Mark ok to delete when no files and subfolders
    If oFDR.Files.Count = 0 And oFDR.SubFolders.Count = 0 Then
        bDeleteFolder = True
    End If
    DeleteEmptyFolderOnly = bDeleteFolder
End Function
