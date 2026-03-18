Option Compare Database
Option Explicit

' =============================================================================
' Module:       mBomProcessing
' Purpose:      BOM categorization and material identification via
'               rule-based matching.  Shared utilities live in maMain.
'               Size parsing has been moved to mSizeSplitting.
' Inputs:       Tables: d_bom, parse_indx_code, parse_matl
' Outputs:      Updates id_categ, id_matl in d_bom
' Dependencies: maMain module, DAO
' By:           T.Sciple, 02/28/2026
' =============================================================================


' --- Module-level Constants --------------------------------------------------
'Private Const TBL_PARSE_DEF As String = "parse_indx_code"
'Private Const TBL_MATL_DEF As String = "parse_matl"


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

' --- UDT describing a rule-source table and its column mapping ---------------
' Pass one of these to LoadRules() so the same loader works for category,
' material, ASTM designation, or any future rule table.
' All rule tables must use standardized column names for matching:
'   incl_all  – AND-match tokens
'   incl_any  – OR-match tokens
'   not_any   – exclusion tokens
Private Type RuleSourceDef
    tbl_name      As String      ' source table name
    fld_id        As String      ' PK / rule-id column
    fld_active    As String      ' is_active column  ("" = no filter)
    order_by      As String      ' ORDER BY clause   (without "ORDER BY")
    target_field  As String      ' column to update in d_bom (e.g. "id_categ")
End Type


' #############################################################################
'            CATEGORIZATION & MATERIAL IDENTIFICATION
' #############################################################################


' =============================================================================
' Entry Point: Classify all BOM fields (categ, material, etc.) in one pass.
' Iterates through every rule-source definition returned by BuildRuleSourceDefs.
' =============================================================================
Public Sub ClassifyAllBomFields()
    Dim defs() As RuleSourceDef
    Dim def_count As Long
    Call BuildRuleSourceDefs(defs, def_count)

    Dim i As Long
    Dim rules() As BomRule
    Dim rule_count As Long

    For i = 0 To def_count - 1
        LoadRules defs(i), rules, rule_count

        If rule_count > 0 Then
            ApplyRulesToBom rules, rule_count, defs(i).target_field
        Else
            MsgBox "No active rules found in table '" & defs(i).tbl_name & "'.", _
                   vbExclamation
        End If
    Next i
End Sub


' =============================================================================
' BuildRuleSourceDefs – single place that lists every rule table.
' Add new entries here when new classification tables are created.
' =============================================================================
Private Sub BuildRuleSourceDefs(ByRef defs() As RuleSourceDef, _
                                ByRef defCount As Long)
    defCount = 2
    
    ReDim defs(0 To defCount - 1)

    ' --- Category rules ------------------------------------------------------
    With defs(0)
        .tbl_name = "parse_indx_code"
        .fld_id = "id_pd_code"
        .fld_active = ""
        .order_by = "[id_pd_code]"
        .target_field = "id_categ"
    End With

    ' --- Material rules ------------------------------------------------------
    With defs(1)
        .tbl_name = "parse_matl"
        .fld_id = "id_matl"
        .fld_active = ""
        .order_by = "[id_matl]"
        .target_field = "id_matl"
    End With
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
                  " WHERE [id_bom] = " & ids(i)
            db.Execute sql, dbFailOnError
        End If
    Next i

    Set db = Nothing
    'MsgBox targetField & " update complete.  " & desc_count & _
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
        If Not mz_util.AllTokensFound(descStr, rule.inclAll) Then Exit Function
    End If

    If rule.has_inclAny Then
        If Not mz_util.AnyTokenFound(descStr, rule.inclAny) Then Exit Function
    End If

    If rule.has_exclAny Then
        If mz_util.AnyTokenFound(descStr, rule.exclAny) Then Exit Function
    End If

    ' Require at least one inclusive condition to have been defined
    If (Not rule.has_inclAll) And (Not rule.has_inclAny) Then Exit Function

    RuleMatches = True
End Function


' --- Generalized rule loader --------------------------------------------------

' =============================================================================
' LoadRules – single loader for any rule-source table.
' Reads from the table / columns described in src, populates rules() and
' ruleCount.  Rows where fld_active = 0 are skipped (if fld_active is set).
' Uses the same robust pattern as the original LoadCategRules.
' =============================================================================
Private Sub LoadRules(ByRef src As RuleSourceDef, _
                      ByRef rules() As BomRule, _
                      ByRef ruleCount As Long)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
             "SELECT * FROM [" & src.tbl_name & "] ORDER BY " & src.order_by, _
             dbOpenSnapshot)

    ' --- empty table ---------------------------------------------------------
    If rs.EOF And rs.BOF Then
        ruleCount = 0
        rs.Close
        Set rs = Nothing
        Set db = Nothing
        Exit Sub
    End If

    rs.MoveLast
    Dim totalRows As Long
    totalRows = rs.RecordCount
    rs.MoveFirst
    ReDim rules(0 To totalRows - 1)

    Dim idx As Long: idx = 0
    Dim raw_inclAll As String
    Dim raw_inclAny As String
    Dim raw_exclAny As String
    Dim has_active_filter As Boolean
    has_active_filter = (Len(src.fld_active) > 0)

    Do While Not rs.EOF

        ' --- skip inactive rows (when an active-flag column is defined) ------
        If has_active_filter Then
            If CLng(Nz(rs.Fields(src.fld_active), 1)) = 0 Then
                rs.MoveNext
                GoTo NextRow
            End If
        End If

        With rules(idx)
            .rule_id = Nz(rs.Fields(src.fld_id), 0)

            ' --- incl_all (AND) ----------------------------------------------
            raw_inclAll = Trim(Nz(rs!incl_all, ""))
            If Len(raw_inclAll) > 0 Then
                .inclAll = Split(raw_inclAll, DELIM)
                .has_inclAll = True
            Else
                .has_inclAll = False
            End If

            ' --- incl_any (OR) -----------------------------------------------
            raw_inclAny = Trim(Nz(rs!incl_any, ""))
            If Len(raw_inclAny) > 0 Then
                .inclAny = Split(raw_inclAny, DELIM)
                .has_inclAny = True
            Else
                .has_inclAny = False
            End If

            ' --- not_any (exclusion) -----------------------------------------
            raw_exclAny = Trim(Nz(rs!not_any, ""))
            If Len(raw_exclAny) > 0 Then
                .exclAny = Split(raw_exclAny, DELIM)
                .has_exclAny = True
            Else
                .has_exclAny = False
            End If
        End With

        idx = idx + 1
        rs.MoveNext
NextRow:
    Loop

    ruleCount = idx

    ' --- trim array if inactive rows were skipped ----------------------------
    If ruleCount = 0 Then
        Erase rules
    ElseIf ruleCount < totalRows Then
        ReDim Preserve rules(0 To ruleCount - 1)
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub


' =============================================================================
' Immediate-window test helper
' =============================================================================
Public Sub TestRuleMatch(ByVal descText As String)
    Dim defs() As RuleSourceDef
    Dim def_count As Long
    BuildRuleSourceDefs defs, def_count

    Dim i As Long
    Dim rules() As BomRule
    Dim rule_count As Long
    Dim matched_id As Long

    Debug.Print "Desc: " & descText
    For i = 0 To def_count - 1
        LoadRules defs(i), rules, rule_count
        matched_id = FindMatchingRuleId(descText, rules, rule_count)
        Debug.Print "  " & defs(i).target_field & ": " & matched_id
    Next i
End Sub