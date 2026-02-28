Option Compare Database
Option Explicit

' =============================================================================
' Module:       mBomProcessing
' Purpose:      Combined BOM processing – size parsing, categorization, and
'               material identification.  Shared utilities live in maMain.
' Inputs:       Tables: d_bom, parse_def_indx_code, material_parse_def
' Outputs:      Updates sz_1, sz_2, desc, categ_id, matl_id in d_bom
' Dependencies: maMain module, DAO
' By:           T.Sciple, 02/28/2026
' =============================================================================


' --- Module-level Constants --------------------------------------------------
Private Const TBL_PARSE_DEF As String = "parse_def_indx_code"
Private Const TBL_MATL_DEF As String = "parse_matl_def"


' --- Unified UDT for any rule-based matching ---------------------------------
' Used by both category and material processing.  Fields not applicable
' to a given rule set are simply left at their defaults (False / empty).
Private Type BomRule
    rule_id      As Long
    inclAll()    As String       ' every token must appear   (AND)
    inclAny()    As String       ' at least one token        (OR)
    exclAny()    As String       ' none of these may appear  (exclusion)
    has_inclAll  As Boolean
    has_inclAny  As Boolean
    has_exclAny  As Boolean
End Type

' --- UDT returned by StripAnyNonSizeRelatedText ------------------------------
Private Type StripResult
    size_text       As String   ' numeric / size characters  (0-9 . / -)
    non_size_text   As String   ' everything else  (text tokens for desc)
End Type

' --- UDT returned by ParseSizes ----------------------------------------------
Private Type SizeResult
    sz(2)       As String       ' primary size (decimal inches or "FLAT")
    col_marker  As Integer      ' position of last inch-mark consumed
    txt_saved   As String       ' stripped text returned to desc
End Type


' #############################################################################
'                         SECTION 1 – SIZE PARSING
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



' =============================================================================
' Entry Point: Parse sizes from desc_w_size and update sz_1, sz_2, desc
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


' --- Size helpers ------------------------------------------------------------

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


' #############################################################################
'            SECTION 2 – CATEGORIZATION & MATERIAL IDENTIFICATION
' #############################################################################


' =============================================================================
' Entry Point: Categorize each BOM line by matching desc against parse rules
' =============================================================================
Public Sub CategorizeBom()
    Dim rules() As BomRule
    Dim rule_count As Long
    LoadCategRules rules, rule_count
    ApplyRulesToBom rules, rule_count, "categ_id"
End Sub


' =============================================================================
' Entry Point: Assign material IDs to BOM rows by matching desc against rules
' =============================================================================
Public Sub IdentifyBomMaterialIds()
    Dim rules() As BomRule
    Dim rule_count As Long
    LoadMaterialRules rules, rule_count

    If rule_count = 0 Then
        MsgBox "No active material rules found in table '" & TBL_MATL_DEF & "'.", _
               vbExclamation
        Exit Sub
    End If

    ApplyRulesToBom rules, rule_count, "matl_id"
End Sub


' =============================================================================
' Shared: Apply a loaded rule set to d_bom, updating targetField
' =============================================================================
Private Sub ApplyRulesToBom(ByRef rules() As BomRule, _
                            ByVal ruleCount As Long, _
                            ByVal targetField As String)
    Dim descs() As String
    Dim ids() As Long
    Dim desc_count As Long
    LoadBomDescs descs, ids, desc_count

    Dim db As DAO.Database
    Set db = CurrentDb

    Dim i As Long
    Dim matched_id As Long
    Dim sql As String

    For i = 0 To desc_count - 1
        matched_id = FindMatchingRuleId(descs(i), rules, ruleCount)
        If matched_id > 0 Then
            sql = "UPDATE [" & TBL_BOM_RAW & "]" & _
                  " SET [" & targetField & "] = " & matched_id & _
                  " WHERE [id_raw] = " & ids(i)
            db.Execute sql, dbFailOnError
        End If
    Next i

    Set db = Nothing
    MsgBox targetField & " update complete.  " & desc_count & _
           " rows processed.", vbInformation
End Sub


' =============================================================================
' Shared: Find the first matching rule id for a description string
' =============================================================================
Private Function FindMatchingRuleId(ByVal descStr As String, _
                                    ByRef rules() As BomRule, _
                                    ByVal ruleCount As Long) As Long
    Dim i As Long
    For i = 0 To ruleCount - 1
        If RuleMatches(descStr, rules(i)) Then
            FindMatchingRuleId = rules(i).rule_id
            Exit Function
        End If
    Next i
    FindMatchingRuleId = 0
End Function


' =============================================================================
' Shared: Evaluate one BomRule against a description string
'   1. inclAll  – every token must appear          (AND)
'   2. inclAny  – at least one token must appear    (OR)
'   3. exclAny  – none of these tokens may appear   (exclusion)
' =============================================================================
Private Function RuleMatches(ByVal descStr As String, _
                             ByRef rule As BomRule) As Boolean
    RuleMatches = False

    If rule.has_inclAll Then
        If Not AllTokensFound(descStr, rule.inclAll) Then Exit Function
    End If

    If rule.has_inclAny Then
        If Not AnyTokenFound(descStr, rule.inclAny) Then Exit Function
    End If

    If rule.has_exclAny Then
        If AnyTokenFound(descStr, rule.exclAny) Then Exit Function
    End If

    ' Require at least one inclusive condition to have been defined
    If (Not rule.has_inclAll) And (Not rule.has_inclAny) Then Exit Function

    RuleMatches = True
End Function


' --- Rule loaders (separate because source tables have different schemas) ----

Private Sub LoadCategRules(ByRef rules() As BomRule, _
                           ByRef ruleCount As Long)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
             "SELECT * FROM " & TBL_PARSE_DEF & " ORDER BY id_pd_code", _
             dbOpenSnapshot)

    If rs.EOF And rs.BOF Then
        ruleCount = 0
        rs.Close
        Set rs = Nothing
        Set db = Nothing
        Exit Sub
    End If

    rs.MoveLast
    ruleCount = rs.RecordCount
    rs.MoveFirst
    ReDim rules(0 To ruleCount - 1)

    Dim idx As Long: idx = 0
    Dim raw_inclAll As String
    Dim raw_inclAny As String
    Dim raw_exclAny As String

    Do While Not rs.EOF
        With rules(idx)
            .rule_id = Nz(rs!id_pd_code, 0)

            raw_inclAll = Trim(Nz(rs!desc_incl_all, ""))
            raw_inclAny = Trim(Nz(rs!desc_incl_any, ""))
            raw_exclAny = Trim(Nz(rs!desc_not_any, ""))

            If Len(raw_inclAll) > 0 Then
                .inclAll = Split(raw_inclAll, DELIM)
                .has_inclAll = True
            Else
                .has_inclAll = False
            End If

            If Len(raw_inclAny) > 0 Then
                .inclAny = Split(raw_inclAny, DELIM)
                .has_inclAny = True
            Else
                .has_inclAny = False
            End If

            If Len(raw_exclAny) > 0 Then
                .exclAny = Split(raw_exclAny, DELIM)
                .has_exclAny = True
            Else
                .has_exclAny = False
            End If
        End With

        idx = idx + 1
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub


Private Sub LoadMaterialRules(ByRef rules() As BomRule, _
                              ByRef ruleCount As Long)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
             "SELECT [id_mat_rule], [match_token], [exclude_token], " & _
             "[priority], [is_active] " & _
             "FROM [" & TBL_MATL_DEF & "] " & _
             "ORDER BY [priority] DESC, [id_mat_rule]", _
             dbOpenSnapshot)

    If rs.EOF And rs.BOF Then
        ruleCount = 0
        rs.Close
        Set rs = Nothing
        Set db = Nothing
        Exit Sub
    End If

    rs.MoveLast
    ReDim rules(0 To rs.RecordCount - 1)
    rs.MoveFirst

    Dim idx As Long
    idx = 0

    Dim raw_match As String
    Dim raw_excl As String
    Dim is_active_value As Variant

    Do While Not rs.EOF
        is_active_value = Nz(rs!is_active, 1)

        If CLng(is_active_value) <> 0 Then
            With rules(idx)
                .rule_id = Nz(rs!id_mat_rule, 0)

                raw_match = Trim(Nz(rs!match_token, ""))
                raw_excl = Trim(Nz(rs!exclude_token, ""))

                If Len(raw_match) > 0 Then
                    .inclAny = Split(raw_match, DELIM)
                    .has_inclAny = True
                Else
                    .has_inclAny = False
                End If

                If Len(raw_excl) > 0 Then
                    .exclAny = Split(raw_excl, DELIM)
                    .has_exclAny = True
                Else
                    .has_exclAny = False
                End If
            End With

            idx = idx + 1
        End If

        rs.MoveNext
    Loop

    ruleCount = idx

    If ruleCount = 0 Then
        Erase rules
    ElseIf ruleCount < rs.RecordCount Then
        ReDim Preserve rules(0 To ruleCount - 1)
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub


' =============================================================================
' Immediate-window test helper
' =============================================================================
Public Sub TestMaterialMatch(ByVal descText As String)
    Dim rules() As BomRule
    Dim rule_count As Long
    LoadMaterialRules rules, rule_count

    Dim matched_id As Long
    matched_id = FindMatchingRuleId(descText, rules, rule_count)

    Debug.Print "Desc: " & descText
    Debug.Print "Matched matl_id: " & matched_id
End Sub