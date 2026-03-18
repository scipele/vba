Option Compare Database
Option Explicit

' =============================================================================
' Module:       maMain
' Purpose:      Shared constants, utilities, and master entry point for
'               BOM processing (size parsing, categorization, material ID).
' Dependencies: DAO (Microsoft Office xx.0 Access Database Engine Object Library)
' By:           T.Sciple, 02/28/2026
' =============================================================================


' --- Shared Constants --------------------------------------------------------
Public Const DELIM As String = "|"
Public Const TBL_BOM_RAW As String = "d_bom"


' =============================================================================
' Master Entry Point – runs all BOM processing steps in sequence
' ===================================================4==========================
Public Sub RunAllBomProcessing()
    Call mbSizeSplitting.ParseAndUpdateBomSizes
    Call mcBomProcessing.ClassifyAllBomFields
    MsgBox "All BOM processing steps complete.", vbInformation
End Sub


' =============================================================================
' Clear all data in specified fields of d_bom table
' =============================================================================
Public Sub ClearBomProcessedData()
    Dim db As DAO.Database
    Dim sql As String
    
    On Error GoTo ErrorHandler
    
    Set db = CurrentDb
    
    ' Update all records in d_bom, setting specified fields to NULL
    sql = "UPDATE [" & TBL_BOM_RAW & "] " & _
          "SET [sz_1] = Null, " & _
          "[sz_2] = Null, " & _
          "[desc] = Null, " & _
          "[id_categ] = Null, " & _
          "[id_matl] = Null"
    
    db.Execute sql, dbFailOnError
    
    MsgBox "All data cleared from d_bom table.", vbInformation
    
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error clearing d_bom data: " & Err.Description, vbCritical
    Set db = Nothing
End Sub


' =============================================================================
' Shared: Load d_bom descriptions into parallel arrays
' =============================================================================
Public Sub LoadBomDescs(ByRef descs() As String, _
                        ByRef rawIds() As Long, _
                        ByRef descCount As Long)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
             "SELECT [id_bom], [desc] FROM [" & TBL_BOM_RAW & "]", _
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
        rawIds(idx) = Nz(rs!id_bom, 0)
        descs(idx) = Trim(Nz(rs!desc, ""))
        idx = idx + 1
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub


' =============================================================================
' Shared: Token-matching helpers
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
' Shared: String Utilities
' =============================================================================

' Strips characters outside printable ASCII range (32–126).
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