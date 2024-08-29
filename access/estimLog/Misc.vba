Option Compare Database
Option Explicit


' Remove non-numeric characters from the input text, excluding the opening parenthesis
Function GetNumericText(inputText As String) As String
    Dim result As String
    Dim i As Integer

    For i = 1 To Len(inputText)
        If IsNumeric(Mid(inputText, i, 1)) Then
            result = result & Mid(inputText, i, 1)
        ElseIf Mid(inputText, i, 1) = "(" And i > 1 Then
            ' Allow the opening parenthesis if it is not the first character
            result = result & "("
        End If
    Next i

    GetNumericText = result
End Function


'Various Functions
Function getTextLft(startLoc, text, matchTxt, noDelimeters)
    Dim i, location

    On Error GoTo ErrorHandler

    For i = 1 To noDelimeters
        location = InStr(startLoc, text, matchTxt, vbTextCompare)
        startLoc = location + 1
    Next i
    getTextLft = Trim(Left(text, location - 1))
    
    Exit Function 'Exit if there are no errors
  
ErrorHandler:
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Err.Clear
End Function

