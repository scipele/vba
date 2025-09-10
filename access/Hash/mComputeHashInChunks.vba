Option Compare Database
Option Explicit

'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | mComputeHashesInChunks.vba                                  |
'| EntryPoint   | ComputeHashForTableArray                                    |
'| Purpose      | Compute Sha1 Hash for text strings in a Table               |
'| Inputs       | MSAccess Table 't_tbl1'  input field 'Description'          |
'| Outputs      | MSAccess Table 't_tbl1' output field 'hash'                 |
'| Dependencies | other module 'mSha1Hash'                                    |
'| By Name,Date | T.Sciple, 9/10/2025                                         |

Sub ComputeHashForTableArray()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim totalRecords As Long
    Dim recordsProcessed As Long
    Dim chunk_size As Long
    Dim i As Long

    ' --- Configuration ---
    Const TABLE_NAME As String = "t_tbl1"
    Const DESCRIPTION_FIELD As String = "Description"
    Const HASH_FIELD As String = "hash"
    chunk_size = 500 ' Process 500 records at a time
    ' ---------------------

    Dim start_time As Double
    start_time = Timer

    On Error GoTo ErrorHandler
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM " & TABLE_NAME, dbOpenDynaset, dbOptimistic)

    If rs.recordCount = 0 Then
        MsgBox "The table is empty.", vbInformation
        GoTo ExitSub
    End If

    rs.MoveLast
    totalRecords = rs.recordCount
    rs.MoveFirst

    recordsProcessed = 0
    Do While Not rs.EOF
        For i = 1 To chunk_size
            If Not rs.EOF Then
                rs.Edit
                rs.Fields(HASH_FIELD).Value = mSha1Hash.GetSha1Hash(rs.Fields(DESCRIPTION_FIELD).Value)
                rs.Update
                rs.MoveNext
                recordsProcessed = recordsProcessed + 1
            Else
                Exit For
            End If
        Next i
        Debug.Print "Processed " & recordsProcessed & " of " & totalRecords & " records."
        DoEvents
    Loop

    Dim elapsed_time As Double
    elapsed_time = Round(Timer - start_time, 2)
    MsgBox "Hashing complete. Updated " & recordsProcessed & " records in " & elapsed_time & " seconds", vbInformation

ExitSub:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    MsgBox "Error Number: " & Err.Number & "Description: " & Err.Description, vbCritical
    GoTo ExitSub
End Sub