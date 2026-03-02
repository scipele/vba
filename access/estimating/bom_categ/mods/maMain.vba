Option Compare Database
Option Explicit

' =============================================================================
' Module:       maMain
' Purpose:      Shared constants, BOM-specific utilities, and master entry
'               point for BOM processing (size parsing, categorization,
'               material ID).
' Dependencies: mz_util, mbSizeSplitting, mcBomProcessing, DAO
' By:           T.Sciple, 02/28/2026
' =============================================================================


' --- Shared Constants --------------------------------------------------------
Public Const DELIM As String = "|"
Public Const TBL_BOM_RAW As String = "d_bom"


' =============================================================================
' Master Entry Point – runs all BOM processing steps in sequence
' =============================================================================
Public Sub RunAllBomProcessing()
    ParseAndUpdateBomSizes
    ClassifyAllBomFields
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
          "[categ_id] = Null, " & _
          "[matl_id] = Null"
    
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