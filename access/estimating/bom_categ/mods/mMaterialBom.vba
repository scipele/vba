Option Compare Database
Option Explicit


' =============================================================================
' Module:       mMaterialBom.vba
' EntryPoint:   IdentifyBomMaterialIds
' Purpose:      Reads material parse rules from 'material_parse_def',
'               evaluates d_bom_raw.desc text against those rules, and updates
'               d_bom_raw.matl_id using numeric rule id.
' Inputs:       Tables: material_parse_def, d_bom_raw
' Outputs:      Updates matl_id field in d_bom_raw
' Dependencies: DAO (Microsoft Office xx.0 Access Database Engine Object Library)
' By:           T.Sciple, 02/27/2026
' =============================================================================


' --- Constants ---------------------------------------------------------------
Private Const DELIM As String = "|"
Private Const TBL_MATL_DEF As String = "material_parse_def"
Private Const TBL_BOM_RAW As String = "d_bom_raw"


' --- UDT for one material rule row ------------------------------------------
Private Type MaterialRule
    id_mat_rule      As Long
    mat_general      As String
    astm_designation As String
    priority         As Long
    inclAny()        As String
    exclAny()        As String
    has_inclAny      As Boolean
    has_exclAny      As Boolean
End Type


' =============================================================================
' Entry Point
' =============================================================================
Public Sub IdentifyBomMaterialIds()
    Dim rules() As MaterialRule
    Dim rule_count As Long
    LoadMaterialRules rules, rule_count

    If rule_count = 0 Then
        MsgBox "No active material rules found in table '" & TBL_MATL_DEF & "'.", _
               vbExclamation
        Exit Sub
    End If

    Dim descs() As String
    Dim raw_ids() As Long
    Dim desc_count As Long
    LoadBomDescs descs, raw_ids, desc_count

    Dim db As DAO.Database
    Set db = CurrentDb

    Dim i As Long
    Dim matched_matl_id As Long
    Dim sql As String

    For i = 0 To desc_count - 1
        matched_matl_id = FindMaterialRuleId(descs(i), rules, rule_count)
        If matched_matl_id > 0 Then
            sql = "UPDATE [" & TBL_BOM_RAW & "]" & _
                  " SET [matl_id] = " & matched_matl_id & _
                  " WHERE [id_raw] = " & raw_ids(i)
            db.Execute sql, dbFailOnError
        End If
    Next i

    Set db = Nothing
    MsgBox "Material ID pass complete. " & desc_count & " rows processed.", vbInformation
End Sub


' =============================================================================
' Load material rules from table
' =============================================================================
Private Sub LoadMaterialRules(ByRef rules() As MaterialRule, _
                              ByRef ruleCount As Long)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
             "SELECT [id_mat_rule], [mat_general], [astm_designation], " & _
             "[match_token], [exclude_token], [priority], [is_active] " & _
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
                .id_mat_rule = Nz(rs!id_mat_rule, 0)
                .mat_general = Nz(rs!mat_general, "")
                .astm_designation = Nz(rs!astm_designation, "")
                .priority = Nz(rs!priority, 0)

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
' Load d_bom_raw descriptions
' =============================================================================
Private Sub LoadBomDescs(ByRef descs() As String, _
                         ByRef rawIds() As Long, _
                         ByRef descCount As Long)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
             "SELECT [id_raw], [desc] FROM [" & TBL_BOM_RAW & "]", _
             dbOpenSnapshot)

    If rs.EOF And rs.BOF Then
        descCount = 0
        rs.Close
        Set rs = Nothing
        Set db = Nothing
        Exit Sub
    End If

    rs.MoveLast
    descCount = rs.RecordCount
    rs.MoveFirst

    ReDim descs(0 To descCount - 1)
    ReDim rawIds(0 To descCount - 1)

    Dim idx As Long
    idx = 0

    Do While Not rs.EOF
        rawIds(idx) = Nz(rs!id_raw, 0)
        descs(idx) = Trim(Nz(rs!desc, ""))

        idx = idx + 1
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub


' =============================================================================
' Matching
' =============================================================================
Private Function FindMaterialRuleId(ByVal descStr As String, _
                                    ByRef rules() As MaterialRule, _
                                    ByVal ruleCount As Long) As Long
    Dim i As Long
    For i = 0 To ruleCount - 1
        If MaterialRuleMatches(descStr, rules(i)) Then
            FindMaterialRuleId = rules(i).id_mat_rule
            Exit Function
        End If
    Next i

    FindMaterialRuleId = 0
End Function


Private Function MaterialRuleMatches(ByVal descStr As String, _
                                     ByRef rule As MaterialRule) As Boolean
    MaterialRuleMatches = False

    If Not rule.has_inclAny Then Exit Function
    If Not AnyTokenFound(descStr, rule.inclAny) Then Exit Function

    If rule.has_exclAny Then
        If AnyTokenFound(descStr, rule.exclAny) Then Exit Function
    End If

    MaterialRuleMatches = True
End Function


Private Function AnyTokenFound(ByVal srcStr As String, _
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


' =============================================================================
' Optional helper for quickly testing one line in Immediate window
' =============================================================================
Public Sub TestMaterialMatch(ByVal descText As String)
    Dim rules() As MaterialRule
    Dim rule_count As Long
    LoadMaterialRules rules, rule_count

    Dim matched_id As Long
    matched_id = FindMaterialRuleId(descText, rules, rule_count)

    Debug.Print "Desc: " & descText
    Debug.Print "Matched matl_id: " & matched_id
End Sub
