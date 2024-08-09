' Subs:     alpha_counting_function
'
'
' Purpose:  This code looks at the values in a particular estimate sheet and sets whether the visibility of the sheets depending on the values
'           found within the workbook
'
'
' Dependencies:  None
'
' By:  T. Sciple, 8/8/2024
'
sub test_example

dim i as Long'
'
for i = 1 to 26
    debug.print i, alpha_ounting_function(i)
next i

'Now start with two letters
debug.print alpha_ounting_function(i)
for i = 27 to 100
    debug.print i, alpha_ounting_function(i)
next i


end sub


Function alpha_ounting_function(ByVal n As Long) As String
    Dim result As String
    Dim ascA As Integer
    
    If n > 0 And n < 16385 Then
        ascA = 64  ' ASCII value of 'A' is 65, so we use 64 to adjust for the 1-based index
        
        result = ""  ' Initialize result as an empty string
        
        Do While n > 0
            n = n - 1  ' Adjust for 1-based index of Excel columns
            result = Chr(ascA + (n Mod 26) + 1) & result  ' Prepend the letter
            n = n \ 26  ' Integer division to reduce n
        Loop
    Else
        result = "Error - Invalid column number"
    End If
    
    colm_no_to_letter = result
End Function