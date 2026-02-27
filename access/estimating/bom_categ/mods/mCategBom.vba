Option Compare Database
Option Explicit


' =============================================================================
' Module:       mCategBom.vba
' EntryPoint:   CategorizeBom
' Purpose:      Reads category parse definitions from 'parse_def_indx_code',
'               reads desc strings from 'd_bom_raw', and assigns a category
'               code to each BOM line item based on pattern matching rules.
' Inputs:       Tables: parse_def_indx_code, d_bom_raw
' Outputs:      Updates categ_id field in d_bom_raw
' Dependencies: DAO (Microsoft Office xx.0 Access Database Engine Object Library)
' By:           T.Sciple, 02/27/2026
' =============================================================================


' --- Constants ---------------------------------------------------------------
Private Const DELIM As String = "|"
Private Const TBL_PARSE_DEF As String = "parse_def_indx_code"
Private Const TBL_BOM_RAW As String = "d_bom_raw"


' --- UDT for one parse-definition row ----------------------------------------
' Each string-array field is "jagged" because pipe-delimited source values
' produce varying element counts after Split().
Private Type ParseDef
    id_pd_code  As Long
    categ       As String
    short_desc  As String
    inclAll()   As String   ' every token must appear (AND logic)
    inclAny()   As String   ' at least one token must appear (OR logic)
    notAny()    As String   ' none of these may appear (exclusion)
    has_inclAll As Boolean
    has_inclAny As Boolean
    has_notAny  As Boolean
End Type


' =============================================================================
' Entry Point
' =============================================================================
Public Sub CategorizeBom()
    ' --- Step 1 : load parse definitions into jagged array of UDTs -----------
    Dim defs() As ParseDef
    Dim def_count As Long
    LoadParseDefs defs, def_count

    ' --- Step 2 : load desc strings from d_bom_raw --------------------------
    Dim descs() As String
    Dim ids() As Long
    Dim desc_count As Long
    LoadBomDescs descs, ids, desc_count

    ' --- Step 3 : match each desc to a category and write back ---------------
    Dim db As DAO.Database
    Set db = CurrentDb

    Dim i As Long
    Dim matched_pd_code As Long
    Dim sql As String

    For i = 0 To desc_count - 1
        matched_pd_code = FindCategoryId(descs(i), defs, def_count)
        If matched_pd_code > 0 Then
            sql = "UPDATE [" & TBL_BOM_RAW & "]" & _
                  " SET [categ_id] = " & matched_pd_code & _
                  " WHERE [id_raw] = " & ids(i)
            db.Execute sql, dbFailOnError
        End If
    Next i

    Set db = Nothing
    MsgBox "BOM categorization complete.  " & desc_count & _
           " rows processed.", vbInformation
End Sub


' =============================================================================
' Step 1 – Load parse_def_indx_code into a jagged UDT array
' =============================================================================
Private Sub LoadParseDefs(ByRef defs() As ParseDef, _
                          ByRef defCount As Long)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
             "SELECT * FROM " & TBL_PARSE_DEF & " ORDER BY id_pd_code", _
             dbOpenSnapshot)

    ' Count rows so we can size the array once
    If rs.EOF And rs.BOF Then
        defCount = 0
        rs.Close
        Set rs = Nothing
        Set db = Nothing
        Exit Sub
    End If

    rs.MoveLast
    defCount = rs.RecordCount
    rs.MoveFirst
    ReDim defs(0 To defCount - 1)

    Dim idx As Long: idx = 0
    Dim raw_inclAll As String
    Dim raw_inclAny As String
    Dim raw_notAny As String

    Do While Not rs.EOF
        With defs(idx)
            .id_pd_code = Nz(rs!id_pd_code, 0)
            .categ = Nz(rs!categ, "")
            .short_desc = Nz(rs!short_desc, "")

            ' --- split pipe-delimited fields into jagged arrays --------------
            raw_inclAll = Trim(Nz(rs!desc_incl_all, ""))
            raw_inclAny = Trim(Nz(rs!desc_incl_any, ""))
            raw_notAny = Trim(Nz(rs!desc_not_any, ""))

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

            If Len(raw_notAny) > 0 Then
                .notAny = Split(raw_notAny, DELIM)
                .has_notAny = True
            Else
                .has_notAny = False
            End If
        End With

        idx = idx + 1
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub


' =============================================================================
' Step 2 – Load d_bom_raw.desc into a dynamic String array
'           (sized at runtime from the record count)
' =============================================================================
Private Sub LoadBomDescs(ByRef descs() As String, _
                         ByRef ids() As Long, _
                         ByRef descCount As Long)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
             "SELECT [id_raw], [desc] FROM [" & TBL_BOM_RAW & "]", dbOpenSnapshot)

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
    ReDim ids(0 To descCount - 1)

    Dim idx As Long: idx = 0
    Do While Not rs.EOF
        ids(idx) = Nz(rs!id_raw, 0)
        descs(idx) = Trim(Nz(rs!desc, ""))
        idx = idx + 1
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub


' =============================================================================
' Step 3 – Category matching algorithm
' =============================================================================
' Evaluation order for each ParseDef row against a description string:
'   1. desc_incl_all  – ALL tokens must be found          (AND)
'   2. desc_incl_any  – at least ONE token must be found  (OR)
'   3. desc_not_any   – NONE of these tokens may be found (exclusion)
'
' A row "matches" when:
'   (incl_all is empty  OR  every incl_all token is found)
'   AND
'   (incl_any is empty  OR  at least one incl_any token is found)
'   AND
'   (not_any  is empty  OR  none of the not_any tokens are found)
'
' First match wins (rows are ordered by id_pd_code).
' Returns the id_pd_code of the first matching row, or 0 if none match.
' -----------------------------------------------------------------------------
Private Function FindCategoryId(ByVal descStr As String, _
                                ByRef defs() As ParseDef, _
                                ByVal defCount As Long) As Long
    Dim i As Long
    For i = 0 To defCount - 1
        If RowMatches(descStr, defs(i)) Then
            FindCategoryId = defs(i).id_pd_code
            Exit Function
        End If
    Next i

    FindCategoryId = 0
End Function


' -----------------------------------------------------------------------------
' Evaluate a single ParseDef row against the upper-cased description.
' -----------------------------------------------------------------------------
Private Function RowMatches(ByVal descStr As String, _
                            ByRef def As ParseDef) As Boolean
    RowMatches = False

    ' --- 1. desc_incl_all : every token must appear --------------------------
    If def.has_inclAll Then
        If Not AllTokensFound(descStr, def.inclAll) Then Exit Function
    End If

    ' --- 2. desc_incl_any : at least one token must appear -------------------
    If def.has_inclAny Then
        If Not AnyTokenFound(descStr, def.inclAny) Then Exit Function
    End If

    ' --- 3. desc_not_any : none of these tokens may appear -------------------
    If def.has_notAny Then
        If AnyTokenFound(descStr, def.notAny) Then Exit Function
    End If

    ' A rule with all three fields empty should NOT auto-match everything;
    ' require at least one inclusive condition to have been defined.
    If (Not def.has_inclAll) And (Not def.has_inclAny) Then Exit Function

    RowMatches = True
End Function


' -----------------------------------------------------------------------------
' Returns True when EVERY element in tokens() is found in srcStr.
' -----------------------------------------------------------------------------
Private Function AllTokensFound(ByVal srcStr As String, _
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


' -----------------------------------------------------------------------------
' Returns True when AT LEAST ONE element in tokens() is found in srcStr.
' -----------------------------------------------------------------------------
Private Function AnyTokenFound(ByVal srcStr As String, _
                               ByRef tokens() As String) As Boolean
    Dim j As Long
    For j = LBound(tokens) To UBound(tokens)
        If InStr(1, srcStr, Trim(tokens(j)), vbTextCompare) > 0 Then
            AnyTokenFound = True
            Exit Function
        End If
    Next j
    AnyTokenFound = False
End Function


' =============================================================================
' Utility – wrap a string value for SQL
' =============================================================================
Private Function QuoteStr(ByVal s As String) As String
    QuoteStr = "'" & Replace(s, "'", "''") & "'"
End Function