Option Compare Database
Option Explicit


' =============================================================================
' Module:       mBomProcessing
' Purpose:      Combined BOM processing – size parsing, categorization, and
'               material identification.  Shared utilities live in maMain.
' Inputs:       Tables: d_bom_raw, parse_def_indx_code, material_parse_def
' Outputs:      Updates sz_1, sz_2, desc, categ_id, matl_id in d_bom_raw
' Dependencies: maMain module, DAO
' By:           T.Sciple, 02/28/2026
' =============================================================================


' --- Module-level Constants --------------------------------------------------
Private Const TBL_PARSE_DEF As String = "parse_def_indx_code"
Private Const TBL_MATL_DEF As String = "material_parse_def"


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

' --- UDT returned by ParseSizes ----------------------------------------------
Private Type SizeResult
    sz1         As String       ' primary size (decimal inches or "FLAT")
    sz2         As String       ' secondary size (decimal inches) or ""
    col_marker  As Integer      ' position of last inch-mark consumed
End Type


' #############################################################################
'                         SECTION 1 – SIZE PARSING
' #############################################################################


Sub TestSizeParse()
    Dim strg As String
    strg = "8"" S/STD BORE A106B SMLS CS PIPE"
    strg = RemoveNonPrintableASCII(strg)
    strg = Replace(strg, "''", """")

    Dim sr As SizeResult
    sr = ParseSizes(strg)

    Debug.Print "sz1=" & sr.sz1 & "  sz2=" & sr.sz2 & _
                "  col_marker=" & sr.col_marker
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
            desc_str = Replace(desc_str, "''", """")

            Dim sr As SizeResult
            sr = ParseSizes(desc_str)

            rs.Edit
            rs!sz_1 = sr.sz1
            rs!sz_2 = sr.sz2
            If sr.col_marker > 0 Then
                rs!desc = Right(desc_str, Len(desc_str) - sr.col_marker - 1)
            Else
                rs!desc = desc_str
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
Public Function ParseSizes(ByVal strg As String) As SizeResult
    Dim result As SizeResult
    Dim shortStr As String
    Dim inchLoc1 As Integer, inchLoc2 As Integer
    Dim lenLoc As Integer, locX As Integer
    Dim tmpSize2 As String

    shortStr = IIf(Len(strg) > 7, Left(strg, 7), strg)

    ' --- Check for "flat" (no numeric size) ---
    If InStr(1, shortStr, "flat", vbTextCompare) > 0 Then
        result.sz1 = "FLAT"
        result.col_marker = InStr(1, shortStr, "flat", vbTextCompare)
        ParseSizes = result
        Exit Function
    End If

    ' --- Locate first inch mark (sz1) ---
    inchLoc1 = InStr(1, shortStr, """", vbTextCompare)
    If inchLoc1 = 0 Then GoTo Done

    result.sz1 = ConvToDecIn(Trim(Left(shortStr, inchLoc1)))
    result.col_marker = inchLoc1

    ' --- Locate & validate second inch mark (sz2) ---
    inchLoc2 = InStr(inchLoc1 + 1, strg, """", vbTextCompare)
    lenLoc = InStr(inchLoc1 + 1, strg, """ long", vbTextCompare)
    If lenLoc = 0 Then lenLoc = InStr(inchLoc1 + 1, strg, """ lg", vbTextCompare)

    Select Case True
        Case inchLoc2 = 0, inchLoc2 > 13       ' no second mark or too far
        Case inchLoc2 = lenLoc                  ' actually a length dimension
        Case Not IsNumeric(Mid(strg, inchLoc2 - 1, 1))  ' non-numeric before mark
        Case Else
            tmpSize2 = Mid(strg, inchLoc1, inchLoc2 - inchLoc1 + 1)
            locX = InStr(1, tmpSize2, "x", vbTextCompare)
            If locX > 0 Then tmpSize2 = Mid(tmpSize2, locX + 1)
            result.sz2 = ConvToDecIn(tmpSize2)
            result.col_marker = inchLoc2
    End Select

Done:
    ParseSizes = result
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
' Shared: Apply a loaded rule set to d_bom_raw, updating targetField
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
