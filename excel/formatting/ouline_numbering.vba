Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | outlineNumbering.vba                                        |
'| EntryPoint   | GetNextOutlineNumber                                        |
'| Purpose      | Generate Next Outline Number                                |
'| Inputs       | prevOutlineNo, levelNo(1,2,3...)                            |
'| Outputs      | outline Number String                                       |
'| Dependencies | none                                                        |
'| By Name/Date | T.Sciple, 11/26/2024                                        |

Enum outlineType
    otSameLevelAsPrevious = 0
    otLowerLevelThanPrevious
    otHigherLevelThanPrevious
End Enum

Public Function GetNextOutlineNumber(ByVal prevOutlineNo As String, _
                                     ByVal levelNo As Integer) As String
    
    Dim parts() As String
    Dim result As String
    Dim i As Integer
    
    'Assume 1. If previous is empty and levelNo = 1
    If Not IsNumeric(Left(prevOutlineNo, 1)) And levelNo = 1 Then
        GetNextOutlineNumber = "1."
        Exit Function
    End If
    
    
    ' Split the outline number into parts based on the "." separator
    parts = Split(prevOutlineNo, ".")
    
    Dim prev_num_level As Integer
    prev_num_level = UBound(parts) + 1
    'Decrease the count if the part is empty string i.e. dot at end
    If parts(UBound(parts)) = "" Then prev_num_level = prev_num_level - 1
        
    Dim outl_type As Integer
    
    outl_type = Switch( _
                            levelNo = prev_num_level, otSameLevelAsPrevious, _
                            levelNo < prev_num_level, otLowerLevelThanPrevious, _
                            levelNo > prev_num_level, otHigherLevelThanPrevious)
    
    Dim format_code As String
    format_code = IIf(levelNo > 1, "00", "0")
    
    Select Case outl_type
        Case otSameLevelAsPrevious
            If IsNumeric(parts(levelNo - 1)) Then
                ' Increment the specified level
                parts(levelNo - 1) = Format(CInt(parts(levelNo - 1)) + 1, format_code)
            End If
        
        Case otLowerLevelThanPrevious
            ' Truncate any excess levels if the number of levels exceeds the desired depth
            If UBound(parts) > levelNo - 1 Then
                ReDim Preserve parts(0 To levelNo - 1)
            End If
            ' Increment the specified level
            parts(levelNo - 1) = Format(CInt(parts(levelNo - 1)) + 1, format_code)
        
        Case otHigherLevelThanPrevious
            'Add element to array if needed
            If UBound(parts) < levelNo - 1 Then
                ReDim Preserve parts(0 To prev_num_level)
            End If
            parts(UBound(parts)) = "01"
    
    End Select
    
    ' Return the updated outline number
    GetNextOutlineNumber = IIf(levelNo = 1, parts(0) & ".", Join(parts, "."))
End Function