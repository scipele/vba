Option Compare Database
Option Explicit

'Create a user defined type that can be utilized passed as a single variable
Type est_data
    BidDate As String
    curClient As String
    curYr As String
    estimateNo As String
    folderName As String
    folderType As String
    folderTypeNo As Integer
    jobLoc As String
    pad As String
    pathSource As String
    estimatorNameID As Long
    estimatorName As String
    estimFileName As String
    shortClient As String
    titleClean As String
    titleOrig As String
End Type


Sub CreateFolderStructure()
    'define an instance of UDT
    Dim ed As est_data
        
    On Error GoTo ErrorHandler
    
    ed.folderType = getFolderType()
    ed.folderName = getFolderName(ed)
    If ed.folderName = "" Then
        MsgBox ("Missing Data To Create Folder Name")
        GoTo ErrorHandler
    Else
        Call make_folder_and_subs(ed)
    End If
    
     Exit Sub 'Exit if there are no errors
  
ErrorHandler:
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Err.Clear
    
End Sub


'1.0 Get Folder Name
Private Function getFolderType() As String
    
    On Error GoTo ErrorHandler
    
    Dim tmpfolderTypeNo As Integer
    tmpfolderTypeNo = InputBox("Enter Number for type of Folder that you want to Make/Open" & vbNewLine _
                & "1 - Mechanical" & vbNewLine _
                & "2 - Shop Fabrication" & vbNewLine _
                & "3 - Shop Paint" & vbNewLine _
                & "4 - Soft Craft" & vbNewLine _
                & "5 - Architectural" & vbNewLine _
                & "6 - Electrical" & vbNewLine _
                & "7 - Cancel" & vbNewLine, "Folder Type")
                
    Dim tmp_type As String
    Select Case tmpfolderTypeNo
        Case 1: tmp_type = "1.ME"
        Case 2: tmp_type = "2.SF"
        Case 3: tmp_type = "3.SP"
        Case 4: tmp_type = "4.SC"
        Case 5: tmp_type = "5.AR"
        Case 6: tmp_type = "6.EL"
        Case Else
            MsgBox ("Correct entry to a valid number indiated above")
    End Select
                
    Forms("frm03EstimData").Controls("tboxFolderType") = tmp_type
    
    getFolderType = tmp_type
     
     Exit Function 'Exit if there are no errors
  
ErrorHandler:
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Err.Clear
End Function


Public Sub OpenFolderStructure()
    'define an instance of UDT
    Dim ed As est_data
    On Error GoTo ErrorHandler
    ed.folderType = getFolderType()
    ed.folderName = getFolderName(ed)
        
    ' Create the command to open the folder
    Dim folderCommand As String
    folderCommand = "explorer.exe """ & ed.folderName & """"

    ' Use the Shell function to open the folder
    Shell folderCommand, vbNormalFocus
     
    Exit Sub 'Exit if there are no errors
  
ErrorHandler:
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Err.Clear
End Sub


Private Function getFolderName(ByRef ed As est_data)
    On Error GoTo ErrorHandler
    ed.titleOrig = Form_frm03EstimData.Title
    'Check to make sure that Title does not contain any charcaters that are not allowed for folder names
    ed.titleClean = CleanStringForFolderName(ed.titleOrig)
    If ed.titleClean <> ed.titleOrig Then
        Form_frm03EstimData.Title.Value = ed.titleClean
    End If
    
    ed.estimateNo = Form_frm03EstimData.EstimNo
    Form_frm03EstimData.cboClientCityState.SetFocus
    ed.curClient = Form_frm03EstimData.cboClientCityState.text
    If ed.curClient = "" Then
        MsgBox ("Missing Client Name")
        GoTo ErrorHandler
    Else
        ed.shortClient = Misc.getTextLft(1, ed.curClient, ",", 1)
    End If
    
    ed.curYr = "20" & Misc.getTextLft(1, ed.estimateNo, "-", 1)
    ed.folderName = "\\Rds\root\T\Estimates\" & ed.curYr & "\" & ed.folderType & "\" & ed.shortClient & "\" & ed.estimateNo & " " & ed.titleClean
    getFolderName = ed.folderName
    Exit Function 'Exit if there are no errors
  
ErrorHandler:
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Err.Clear
End Function


Sub make_folder_and_subs(ByRef ed As est_data)

    Dim i, j, Count, SubCount As Long
    On Error GoTo ErrorHandler
    
    'Pad Path names with "\" if its not there
    If Not Right(ed.folderName, 1) = "\" Then
        ed.folderName = ed.folderName & "\"
    End If
    
    'Check if folder exists
    If Dir(ed.folderName, vbDirectory) = "" Then

        'If folder does not exist, create it
        Dim arrFolders() As String
        Dim strFolder As String
        
        arrFolders = Split(ed.folderName, "\")
        'Put together initial folder
        strFolder = "\\" & arrFolders(2) & "\" & arrFolders(3) & "\" & arrFolders(4)
        
        For j = 5 To UBound(arrFolders)
            strFolder = strFolder & "\" & arrFolders(j)
            If Dir(strFolder, vbDirectory) = "" Then
                MkDir strFolder
                If j = 1 Then
                    'Display Message 1
                    MsgBox ("Created One Main Folder")
                Else
                    SubCount = SubCount + 1
                End If
            End If
        Next j
        
        'Display Message 2
        If j > 1 Then MsgBox ("Created " & SubCount + 1 & " Subfolder(s)")
        
        'If Mechanical Folder then Copy over Template
        If ed.folderType = "1.ME" Or ed.folderType = "3.SP" Then
            Call CopyFilesAndFolders(ed)
        End If
        
    Else
        MsgBox ("Folder already Exist")
    End If

    Exit Sub 'Exit if there are no errors
  
ErrorHandler:
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Err.Clear
End Sub


Sub CopyFilesAndFolders(ByRef ed As est_data)
'copy over template folder structure

    On Error GoTo ErrorHandler
    
    Dim pathSource As String
    Select Case ed.folderType
    
        Case "1.ME"
            ' Set the source path
            pathSource = "\\Rds\root\T\Estimates\FolderStructure\1.ME\2X-XXXX Short Title\*.*" ' Replace with your source folder path
        
        Case "3.SP"
            ' Set the source path
            pathSource = "\\Rds\root\T\Estimates\FolderStructure\3.SP\2X-XXXX Short Title\*.*" ' Replace with your source folder path
    End Select
    
    ' Build the xcopy command
    Dim shellCommand As String
    shellCommand = "xcopy """ & pathSource & """ """ & ed.folderName & """ /E /I /C /K /Y"
    
    ' Execute the command through the Shell
    Call Shell(shellCommand, vbNormalFocus)
    
    If ed.folderType = "1.ME" Then
        Call UpdEstimData.rename_estim_sht(ed)
        Call UpdEstimData.prep_data_then_update(ed)
    End If
    
    Exit Sub 'Exit if there are no errors
  
ErrorHandler:
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Err.Clear
    
End Sub


Function CleanStringForFolderName(ByVal inputString As String) As String
    Dim invalidChars As String
    Dim resultString As String
    Dim i As Integer
    Dim currentChar As String

    On Error GoTo ErrorHandler

    ' List of characters not allowed in Windows folder names
    invalidChars = "\/:*?""<>|"

    ' Initialize the result string
    resultString = ""

    ' Loop through each character in the input string
    For i = 1 To Len(inputString)
        ' Get the current character
        currentChar = Mid(inputString, i, 1)

        ' Check if the character is not allowed
        If InStr(1, invalidChars, currentChar) > 0 Then
            ' Replace invalid character with an underscore
            resultString = resultString & "_"
        Else
            ' Keep the valid character
            resultString = resultString & currentChar
        End If
    Next i

    ' Return the cleaned string
    CleanStringForFolderName = resultString
    
    Exit Function ' Exit the subroutine if there are no errors
  
ErrorHandler:
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Err.Clear
    
End Function