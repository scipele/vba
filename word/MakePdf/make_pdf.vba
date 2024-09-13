Option Explicit

Sub MakePdf()
' Purpose: Save a PDF Document From Active Word Document in same folder
'	Insert the following macro into a module in the "Normal" Document Template
'	Create shortcut key as follows:
'		1. Rightclick on blank portion of ribbon area
'		2. Customize the ribbon
'		3. Choose commands from: select MakePdf
'		4. Click on Keyboard shortcuts: [Customize] command button
'   5. Scroll down to Macros - in the Categories combo box
'   6. Slect 'MakePdf' Macro in the Macros: combo box
'		7. Place cursor in the Press new shortcut key textbox
'		8. Press the 'Alt-P' keys together or some other key combination
'		9. Click the [Assign] Command Button
'		10. Close
' By: T.Sciple, 9/13/2024

  Dim UniqueName As Boolean
  UniqueName = False
  Dim myPath As String
  myPath = ActiveDocument.FullName
  Dim CurrentFolder As String
  CurrentFolder = ActiveDocument.Path & "\"
  Dim FileName As String
  FileName = Mid(myPath, InStrRev(myPath, "\") + 1, _
  InStrRev(myPath, ".") - InStrRev(myPath, "\") - 1)
    
  'Check if file exist?
  Do While UniqueName = False
  Dim DirFile As String
    DirFile = CurrentFolder & FileName & ".pdf"
    If Len(Dir(DirFile)) <> 0 Then
      Dim UserAnswer As Variant
      UserAnswer = MsgBox("File Already Exists! Click " & _
        "[Yes] to override. Click [No] to Rename.", vbYesNoCancel)
      
      If UserAnswer = vbYes Then
        UniqueName = True
      ElseIf UserAnswer = vbNo Then
        Do
          'Retrieve New File Name
            FileName = InputBox("Provide New File Name " & _
              "(will ask again if you provide an invalid file name)", _
              "Enter File Name", FileName)
          
          'Exit if User Wants To
            If FileName = "False" Or FileName = "" Then Exit Sub
        Loop While ValidFileName(FileName) = False
      Else
        Exit Sub 'Cancel
      End If
    Else
      UniqueName = True
    End If
  Loop
      
    'Save As PDF Document
      On Error GoTo ErrorTryingToSave
        ActiveDocument.ExportAsFixedFormat _
         OutputFileName:=CurrentFolder & FileName & ".pdf", _
         ExportFormat:=wdExportFormatPDF
      On Error GoTo 0 'diables the on error
    
    'Confirm Save To User
      With ActiveDocument
        Dim FolderName As String
        FolderName = Mid(.Path, InStrRev(.Path, "\") + 1, Len(.Path) - InStrRev(.Path, "\"))
      End With
      
    ActiveDocument.Close SaveChanges:=wdSaveChanges
    Exit Sub
    
    'Error Handlers
ErrorTryingToSave:
      MsgBox "There was a problem saving your PDF. This is most commonly caused" & _
       " by the original PDF file already being open."
      Exit Sub
End Sub


Function ValidFileName(FileName As String) As Boolean
  'Purpose: Determine If A Given Word Document File Name Is Valid

  'Determine Folder Where Temporary Files Are Stored
  Dim TempPath As String
  TempPath = Environ("TEMP")

  'Create a Temporary XLS file (XLS in case there are macros)
  On Error GoTo InvalidFileName
  Dim doc As Document
  Set doc = ActiveDocument.SaveAs2(ActiveDocument.TempPath & _
  "\" & FileName & ".doc", wdFormatDocument)
  On Error Resume Next

  'Delete Temp File
  Kill doc.FullName

  'File Name is Valid
  ValidFileName = True
  'Exit if no error is found
  Exit Function

  'Error Handler Label
  InvalidFileName:
  'File Name is Invalid
  ValidFileName = False
End Function