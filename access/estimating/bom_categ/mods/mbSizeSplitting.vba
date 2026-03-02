Option Compare Database
Option Explicit

' =============================================================================
' Module:       mSizeSplitting
' Purpose:      Parse size dimensions from BOM description strings and convert
'               fractional inch notation to decimal values.
' Inputs:       Table: d_bom (desc_w_size column)
' Outputs:      Updates sz_1, sz_2, desc in d_bom
' Dependencies: maMain module (RemoveNonPrintableASCII, TBL_BOM_RAW), DAO
' By:           T.Sciple, 02/28/2026
' =============================================================================


' --- UDT returned by StripAnyNonSizeRelatedText ------------------------------
Private Type StripResult
    size_text       As String   ' numeric / size characters  (0-9 . / -)
    non_size_text   As String   ' everything else  (text tokens for desc)
End Type

' --- UDT returned by ParseSizes ----------------------------------------------
Public Type SizeResult
    sz(2)       As String       ' primary size (decimal inches or "FLAT")
    col_marker  As Integer      ' position of last inch-mark consumed
    txt_saved   As String       ' stripped text returned to desc
End Type


' #############################################################################
'                         TEST HELPER
' #############################################################################

Sub TestSizeParse()

    Const num_parts = 4#
    Dim strgs(num_parts) As String
    
    strgs(0) = ".75"" x .5"" CON-SWAGE S/160 A234-WPB SMLS CS (PBE)"
    strgs(1) = ".5"" FNPT x .5"" TUBING ADAPTER 316SS"
    strgs(2) = ".5"" x 3"" LONG NIPPLE A312-TP316/316L SS DUAL MARKED (TBE)"
    strgs(3) = ".5"" 2000# THRD 90D ELL 316/316L SS DUAL MARKED"
    
    Dim s As String
    Dim i As Integer
    
    For i = 0 To num_parts - 1
        s = RemoveNonPrintableASCII(strgs(i))
        s = Replace(s, "''", """")
        Dim sr As SizeResult
        sr = ParseSizes(s)
        Debug.Print "indx:" & i & " " & _
                    "sz(1)=", sr.sz(0), _
                    "sz(2)=", sr.sz(1), _
                    "col_marker=", sr.col_marker, _
                    "string:", strgs(i)
    Next i

End Sub


' #############################################################################
'                         ENTRY POINT
' #############################################################################

' =============================================================================
' Parse sizes from desc_w_size and update sz_1, sz_2, desc in d_bom
' =============================================================================
Public Sub ParseAndUpdateBomSizes()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM " & TBL_BOM_RAW, dbOpenDynaset)
    Dim desc_str As String

    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do While Not rs.EOF
            desc_str = Nz(rs!desc_w_size, "")
            desc_str = RemoveNonPrintableASCII(desc_str)
            Dim sr As SizeResult
            sr = ParseSizes(desc_str)

            rs.Edit
            rs!sz_1 = sr.sz(0)
            rs!sz_2 = sr.sz(1)
            If sr.col_marker > 0 Then
                rs!desc = Right(desc_str, Len(desc_str) - sr.col_marker - 1)
            Else
                rs!desc = sr.txt_saved & " " & desc_str
            End If
            rs.Update
            rs.MoveNext
        Loop
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing
    MsgBox "Size parsing and update complete.", vbInformation
End Sub


' #############################################################################
'                         SIZE HELPERS
' #############################################################################

' Parses both sz1 and sz2 from a description string in one pass.
' Returns a SizeResult UDT with sz1, sz2, and col_marker.
' Handles all preprocessing internally (trimming, quote normalization, ASCII cleanup).
Public Function ParseSizes(ByVal strg As String) As SizeResult
    Dim result As SizeResult
    Dim str_parts() As String
    Dim partCount As Integer
    Dim sz2Raw As String, locX As Integer

    ' --- Preprocessing: limit length, normalize quotes, clean ASCII ---
    If Len(strg) > 14 Then strg = Left(strg, 14)
    strg = Replace(strg, "''", """")
    str_parts = Split(Replace(strg, """", "|"), "|")

    ' --- split the string---
    partCount = UBound(str_parts)   ' 0=no marks, 1=one mark, 2+=two marks
    
    Dim i As Integer
    Dim special_case As Integer
    For i = 0 To partCount

        'Check for special cases
        If InStr(1, str_parts(i), "lo", vbTextCompare) > 0 Then special_case = 1
        If InStr(1, str_parts(i), "lg", vbTextCompare) > 0 Then special_case = 1
        If InStr(1, str_parts(i), "flat", vbTextCompare) > 0 Then special_case = 2
        
        Select Case special_case
            Case 1
                ' If length type measurement is found then reset the sz(i) to an empty string for the previous element
                If i > 0 Then result.sz(i - 1) = ""     'Length indicators are not actual size dimensions, so return empty string for sz1 and sz2.
            Case 2
                result.sz(i) = "FLAT"   'If "flat" is found anywhere in the description, set sz1 to "FLAT" and ignore any inch marks.
            Case Else
                ' dont process the last part as a numeric part
                If i < partCount Then
                    Dim strip As StripResult
                    strip = StripAnyNonSizeRelatedText(str_parts(i))
                    result.txt_saved = result.txt_saved & strip.non_size_text
                    result.sz(i) = ConvToDecIn(strip.size_text)
                    result.col_marker = result.col_marker + Len(str_parts(i)) + 1
                End If
        End Select
    Next i

    ParseSizes = result
End Function

Private Function StripAnyNonSizeRelatedText(ByVal strg As String) As StripResult
    Dim sr As StripResult
    Dim i As Integer
    sr.size_text = ""
    sr.non_size_text = ""
    For i = 1 To Len(strg)
        If Mid(strg, i, 1) Like "[0-9./-]" Then
            sr.size_text = sr.size_text & Mid(strg, i, 1)
        Else
            sr.non_size_text = sr.non_size_text & Mid(strg, i, 1)
        End If
    Next i
    StripAnyNonSizeRelatedText = sr
End Function


' Converts an inch string (e.g. 4", 1-1/2") to a decimal inch value.
' Expects '' already normalized to " by the caller.
Public Function ConvToDecIn(ByVal measStr As String) As Variant
    measStr = Replace(measStr, """", "")
    measStr = Trim(measStr)
    If Len(measStr) = 0 Then
        ConvToDecIn = ""
        Exit Function
    End If
    ConvToDecIn = InchStrToDec(measStr)
End Function


Private Function InchStrToDec(InchPartStr As String)
    Dim InchPartAry() As String

    InchPartStr = Replace(InchPartStr, " ", "-", 1, -1, vbTextCompare)
    InchPartStr = Replace(InchPartStr, "-", ",", 1, -1, vbTextCompare)
    InchPartStr = Replace(InchPartStr, "/", ",", 1, -1, vbTextCompare)

    InchPartAry = Split(InchPartStr, ",")

    Select Case UBound(InchPartAry)
        Case 0
            InchStrToDec = Val(InchPartStr)
        Case 1
            InchStrToDec = Val(InchPartAry(0) / InchPartAry(1))
        Case 2
            InchStrToDec = Val(InchPartAry(0) + InchPartAry(1) / InchPartAry(2))
    End Select
End Function
