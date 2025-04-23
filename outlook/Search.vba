' filename:         Search.vba
'
' Purpose:          search the subject line in current mailbox
'                   for a particular searched subject
'
' usage:            run file_emails() Sub
'
' Dependencies:     Uses Reference Library Microsoft Scripting Runtime to facilitate use of
'                   dictionary object early binding
'
' By:               T.Sciple, 09/16/2024


Sub SearchBySubject()
    Dim myExplorer As Outlook.Explorer
    Dim mySearch As String
    Dim subjectInput As String
    
    ' Prompt the user for the subject to search
    subjectInput = InputBox("Enter the subject to search for:", "Search Mailbox by Subject")
    
    ' If the user clicks Cancel or enters an empty string, exit the subroutine
    If subjectInput = "" Then Exit Sub
    
    ' Define the search query
    mySearch = "subject: " & subjectInput
    
    ' Get the current active explorer
    Set myExplorer = Application.ActiveExplorer
    
    ' Set the focus on the search box and perform the search
    myExplorer.search mySearch, olSearchScopeCurrentMailbox
    'myExplorer.search mySearch, olSearchScopeAllFolders
End Sub
