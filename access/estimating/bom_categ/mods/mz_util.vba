Option Compare Database
Option Explicit

' =============================================================================
' Module:       mz_util
' Purpose:      General-purpose utility functions shared across modules.
' By:           T.Sciple, 03/01/2026
' =============================================================================


' =============================================================================
' Token-matching helpers
' =============================================================================

' Returns True when AT LEAST ONE element in tokens() is found in srcStr.
Public Function AnyTokenFound(ByVal srcStr As String, _
                              ByRef tokens() As String) As Boolean
    Dim j As Long
    For j = LBound(tokens) To UBound(tokens)
        If Len(Trim(tokens(j))) > 0 Then
            If InStr(1, srcStr, Trim(tokens(j)), vbTextCompare) > 0 Then
                AnyTokenFound = True
                Exit Function
            End If
        End If
    Next j
    AnyTokenFound = False
End Function


' Returns True when EVERY element in tokens() is found in srcStr.
Public Function AllTokensFound(ByVal srcStr As String, _
                               ByRef tokens() As String) As Boolean
    Dim j As Long
    For j = LBound(tokens) To UBound(tokens)
        If InStr(1, srcStr, Trim(tokens(j)), vbTextCompare) = 0 Then
            AllTokensFound = False
            Exit Function
        End If
    Next j
    AllTokensFound = True
End Function


' =============================================================================
' String Utilities
' =============================================================================

' Strips characters outside printable ASCII range (32-126).
Public Function RemoveNonPrintableASCII(ByVal desc_str As String) As String
    Dim i As Integer
    Dim result As String
    result = ""
    For i = 1 To Len(desc_str)
        Dim charCode As Integer
        charCode = AscW(Mid(desc_str, i, 1))
        If charCode >= 32 And charCode <= 126 Then
            result = result & Mid(desc_str, i, 1)
        End If
    Next i
    RemoveNonPrintableASCII = result
End Function


' Wraps a string value for safe SQL insertion (single-quote escaping).
Public Function QuoteStr(ByVal s As String) As String
    QuoteStr = "'" & Replace(s, "'", "''") & "'"
End Function
