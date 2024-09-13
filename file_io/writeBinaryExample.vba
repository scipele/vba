'filename:  writeBinaryExample.vba
'Purpose:   Simple Example to Write a Binary File
'By:        T. Sciple, 9/13/2024

Option Explicit

Type miscData
    strg1 As String
    numByte As Byte
    numbInt As Integer
    numDbl As Double
    numbLong As Long
End Type

Sub WriteBinaryExample()
    ' Set the file path for the binary file
    Dim fileNum As Integer
    Dim filepath As String
    Dim md As miscData
    
    filepath = "C:\t\test.bin"
    fileNum = FreeFile

    'delete any previous file
    On Error Resume Next
    Kill filepath
    MsgBox "Deleted Previous File: " & filepath, vbInformation

    'hard code some numbers in structure variable for the test
    md.strg1 = "Test String that is 38 characters long"
    md.numByte = 234
    md.numbInt = 1364
    md.numDbl = 3.14159
    md.numbLong = 128634

    ' Open binary file for writing
    Open filepath For Binary Access Write As #fileNum
    ' Write each structure variable individually
        ' Write the fixed part of the structure
        Put #fileNum, , CInt(Len(md.strg1))  'Put the length of string in
        Put #fileNum, , md.strg1
        Put #fileNum, , md.numByte
        Put #fileNum, , md.numbInt
        Put #fileNum, , md.numDbl
        Put #fileNum, , md.numbLong
    ' Close the file
    Close #fileNum
    
    MsgBox ("Binary File written and closed")
End Sub