Option Explicit

'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | outlineNumbering.vba                                        |
'| EntryPoint   | GetNextOutlineNumber                                        |
'| Purpose      | Generate Next Outline Number                                |
'| Inputs       | prevOutlineNo, levelNo(1,2,3...)                            |
'| Output       | outline Number String                                       |
'| Dependencies | none                                                        |
'| By Name/Date | T.Sciple, 11/27/2024                                        |

Enum outlineType
    otStartingNumber
    otErrorPrevious
    otErrorMissingLevel
    otErrorSkippedLevel
    otSameLevel
    otLowerLevel
    otHigherLevel
End Enum


Public Function GetNextOutlineNumber(ByVal prevOutlineNo As String, _
                                     ByVal levelNo As Integer, _
                                     Optional ByVal formatCode As String = "") As String
    
    'Cleanup Previous String for Processing Purposes
    prevOutlineNo = CleanPrevString(prevOutlineNo)
    
    Dim start_flag As Boolean
    start_flag = CheckIfStartingNumber(prevOutlineNo, levelNo)
    
    'Split the outline number into parts based on the "." separator
    Dim parts() As String
    parts = Split(prevOutlineNo, ".")
    
    Dim prev_num_level As Integer
    prev_num_level = UBound(parts) + 1
    
    Dim outl_type As Integer
    outl_type = Switch( _
        start_flag, otStartingNumber, _
        IIf(Len(prevOutlineNo) > 3, Left(prevOutlineNo, 3), "") = "Err", otErrorPrevious, _
        levelNo = 0, otErrorMissingLevel, _
        (levelNo - prev_num_level) > 1, otErrorSkippedLevel, _
        levelNo = prev_num_level, otSameLevel, _
        levelNo < prev_num_level, otLowerLevel, _
        (levelNo - prev_num_level) = 1, otHigherLevel)
    
    GetNextOutlineNumber = GetStrFromParts(outl_type, parts(), prevOutlineNo, prev_num_level, levelNo, formatCode)
End Function


Private Function GetStrFromParts(ByVal outl_type As Integer, _
                                 ByRef parts() As String, _
                                 ByVal prevOutlineNo As String, _
                                 ByVal prev_num_level As Integer, _
                                 ByVal levelNo As Integer, _
                                 ByVal formatCode As String) As String
                                     
    
    Dim format_code As String
    format_code = GetFormatCode(levelNo, formatCode)
    
    Dim pad_lead As String
    pad_lead = GetPadding(levelNo)
    
    Dim str As String
    str = ""
    
    Select Case outl_type
        
        Case otStartingNumber
            str = "1."
        
        Case otErrorPrevious
            str = "Err-previous no error"
        
        Case otErrorMissingLevel
            str = "Err-missing level"
        
        Case otErrorSkippedLevel
            str = "Err-skipped level"
            
        Case otSameLevel
            If IsNumeric(parts(levelNo - 1)) Then
                'Increment the specified level
                parts(levelNo - 1) = Format(CInt(parts(levelNo - 1)) + 1, format_code)
            End If
        
        Case otLowerLevel
            'Truncate any excess levels if the number of levels exceeds the desired depth
            If UBound(parts) > levelNo - 1 Then
                ReDim Preserve parts(0 To levelNo - 1)
            End If
            'Increment the specified level
            parts(levelNo - 1) = Format(CInt(parts(levelNo - 1)) + 1, format_code)
        
        Case otHigherLevel
            'Add element to array if needed
            If UBound(parts) < levelNo - 1 Then
                ReDim Preserve parts(0 To prev_num_level)
            End If
            parts(UBound(parts)) = Format(1, format_code)
    
    End Select
    
    If str = "" Then str = pad_lead & IIf(levelNo = 1, parts(0) & ".", Join(parts, "."))
    GetStrFromParts = str
End Function


Private Function GetFormatCode(ByVal levelNo As Integer, ByVal formatCode As String) As String
    
    If levelNo = 0 Then
        GetFormatCode = "00"
        Exit Function
    End If
    
    ' Determine the format code for the given level
    If formatCode <> "" And levelNo <= Len(formatCode) Then
        GetFormatCode = String(CInt(Mid(formatCode, levelNo, 1)), "0")
    Else
        ' Default to "00" for levels greater than 1 and "0" for level 1
        GetFormatCode = IIf(levelNo > 1, "00", "0")
    End If
End Function


Private Function CheckIfStartingNumber(ByVal prevOutlineNo As String, _
                                      ByVal levelNo As Integer) As Boolean
    
    'Assume starting number if previous is non numberic, levelNo=1, and also not a Previous Error
    If Not IsNumeric(Left(prevOutlineNo, 1)) And levelNo = 1 Then
        If Len(prevOutlineNo) > 2 Then
            If Not Left(prevOutlineNo, 3) = "Err" Then
                CheckIfStartingNumber = True
            End If
        End If
    End If
End Function


Private Function GetPadding(ByVal lvlNo As Integer)
    'set to empty string
    Dim pad As String
    pad = ""
    Dim i As Integer
    If lvlNo > 1 Then
        For i = 1 To lvlNo
            pad = pad & "  "    'Used two spaces from indent per level
        Next i
    End If
    
    GetPadding = pad
End Function


Private Function CleanPrevString(ByVal prevOutlineNo)
    
    'Strip any leading spaces from previous outline number
    Dim str As String
    str = LTrim(prevOutlineNo)
    
    'Strip trailing period from end if it exists
    If Len(prevOutlineNo) > 1 Then
        If Right(prevOutlineNo, 1) = "." Then
            str = Left(prevOutlineNo, Len(prevOutlineNo) - 1)
        End If
    End If

    CleanPrevString = str
End Function