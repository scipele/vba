'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | mFormatting.vba                                             |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple / BJR, 10/1/2025                                   |

Option Compare Database
Option Explicit


' Constants for padding characters
Private Const PAD_ZERO As String = "0"
Private Const PAD_UNDERSCORE As String = "_"
Private Const DEFAULT_PAD_CHAR As String = PAD_UNDERSCORE
Private Const DEFAULT_PAD_AT_BEGINNING As Boolean = True


Function PadStr(inputNo As Variant, numChars As Integer) As String
    ' Pads non-decimal numeric values and non-numeric/null values with underscores at the beginning
    ' For decimal numbers, rounds to numChars characters and pads with zeros at the end if needed
    ' Truncates to numChars if result is too long
    
    Dim str As String
    Dim padChar As String
    Dim padAtBeginning As Boolean
    
    ' Initialize defaults
    padChar = DEFAULT_PAD_CHAR
    padAtBeginning = DEFAULT_PAD_AT_BEGINNING
    
    ' Handle invalid numChars
    If numChars <= 0 Then
        PadStr = ""
        Exit Function
    End If
    
    ' Determine input type and padding behavior
    Select Case True
        Case IsNull(inputNo):
            str = ""
            padChar = PAD_UNDERSCORE
            padAtBeginning = True
            
        Case IsNumeric(inputNo):
            str = CStr(inputNo) ' Preserve original input format
            If InStr(1, str, ".", vbTextCompare) > 0 Then
                str = round_dec_str_to_overal_len(inputNo, numChars)
                padChar = PAD_ZERO
                padAtBeginning = False
            Else
                padChar = PAD_UNDERSCORE
                padAtBeginning = True
            End If
            
        Case Else:
            str = Trim(CStr(inputNo))
            padChar = PAD_UNDERSCORE
            padAtBeginning = True
    End Select
    
    ' Pad the string only if needed
    If Len(str) < numChars Then
        If padAtBeginning Then
            str = String(numChars - Len(str), padChar) & str
        Else
            str = str & String(numChars - Len(str), padChar)
        End If
    End If
    
    ' Truncate if too long
    If Len(str) > numChars Then
        str = Left(str, numChars)
    End If
    
    PadStr = str
End Function


Function round_dec_str_to_overal_len(dec_str As Variant, trucated_str_len As Integer) As String
    ' Rounds a decimal number (as string or numeric) to produce a string of trucated_str_len characters
    ' Includes sign in length calculation for negative numbers
    ' Returns input string for non-decimal numbers, empty string for invalid inputs
    ' Example: "2.0" with length 4 -> "2.00", "20.375" with length 5 -> "20.38"
    
    Dim num As Double
    Dim str As String
    Dim integerPartLength As Integer
    Dim decimalPlaces As Integer
    Dim isNegative As Boolean
    
    ' Handle invalid inputs
    If IsNull(dec_str) Or Not IsNumeric(dec_str) Or trucated_str_len <= 0 Then
        round_dec_str_to_overal_len = ""
        Exit Function
    End If
    
    ' Convert to Double and preserve original string
    num = CDbl(dec_str)
    str = CStr(dec_str) ' Use original input to preserve decimal point
    
    ' Return non-decimal numbers as-is
    If InStr(1, str, ".", vbTextCompare) = 0 Then
        round_dec_str_to_overal_len = str
        Exit Function
    End If
    
    ' Check for negative sign
    isNegative = (Left(str, 1) = "-")
    If isNegative Then
        integerPartLength = Len(Split(str, ".")(0)) - 1 ' Exclude sign
    Else
        integerPartLength = Len(Split(str, ".")(0))
    End If
    
    ' Calculate decimal places for target length, accounting for sign
    decimalPlaces = trucated_str_len - integerPartLength - 1 - IIf(isNegative, 1, 0) ' Subtract 1 for decimal point, 1 for sign if negative
    
    ' Ensure at least 1 decimal place for decimal numbers
    If decimalPlaces < 1 Then decimalPlaces = 1
    
    ' Round and format
    num = Round(num, decimalPlaces)
    round_dec_str_to_overal_len = Format(num, IIf(isNegative, "-0.", "0.") & String(decimalPlaces, "0"))
    
    ' Truncate to target length if needed
    If Len(round_dec_str_to_overal_len) > trucated_str_len Then
        round_dec_str_to_overal_len = Left(round_dec_str_to_overal_len, trucated_str_len)
    End If
End Function


