Option Compare Database
Option Explicit

Sub ComputeHashForTableArray()
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim totalRecords As Long
    Dim batchSize As Long
    Dim minID As Long
    Dim maxID As Long
    Dim startID As Long
    Dim endID As Long
    Dim idArray() As Long
    Dim descArray() As String
    Dim hashArray() As String
    Dim i As Long
    Dim sql As String
    Dim batchCount As Long
    Dim recordCount As Long
    
    ' Set batch size (5000 is safe; adjust to 2000 or 10000 based on testing)
    batchSize = 5000
    
    ' Open database
    Set db = CurrentDb
    
    ' Get total records and ID range
    totalRecords = DCount("*", "t_tbl1")
    minID = DMin("ID", "t_tbl1")
    maxID = DMax("ID", "t_tbl1")
    
    If totalRecords = 0 Or IsNull(minID) Or IsNull(maxID) Then
        MsgBox "No records to process or invalid ID range.", vbExclamation
        Exit Sub
    End If
    
    ' Performance optimizations
    Application.Echo False
    DoCmd.SetWarnings False
    
    ' Initialize batch tracking
    startID = minID
    batchCount = 0
    
    ' Loop through batches using ID ranges
    Do While startID <= maxID
        endID = startID + batchSize - 1
        If endID > maxID Then endID = maxID
        
        ' Fetch batch of records
        sql = "SELECT ID, Description FROM t_tbl1 " & _
              "WHERE ID BETWEEN " & startID & " AND " & endID & " " & _
              "AND (hash IS NULL OR hash = '') AND Description IS NOT NULL"
        Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
        
        ' Count records in batch
        recordCount = 0
        If Not (rs.EOF And rs.BOF) Then
            rs.MoveLast
            recordCount = rs.recordCount
            rs.MoveFirst
        End If
        
        If recordCount > 0 Then
            ' Resize arrays
            ReDim idArray(0 To recordCount - 1)
            ReDim descArray(0 To recordCount - 1)
            ReDim hashArray(0 To recordCount - 1)
            
            ' Populate ID and Description arrays
            i = 0
            Do Until rs.EOF
                idArray(i) = rs!ID
                descArray(i) = rs!Description
                i = i + 1
                rs.MoveNext
            Loop
            
            ' Compute hashes for the batch
            For i = 0 To recordCount - 1
                On Error Resume Next ' Handle potential errors in GetSha1Hash
                hashArray(i) = GetSha1Hash(descArray(i))
                If Err.Number <> 0 Then
                    hashArray(i) = "ERROR"
                    Debug.Print "Hash error for ID " & idArray(i) & ": " & Err.Description
                    Err.Clear
                End If
                On Error GoTo ErrHandler
            Next i
            
            ' Update using temporary table
            Call UpdateBatchWithTempTable(db, idArray, hashArray)
        End If
        
        ' Clean up recordset
        rs.Close
        Set rs = Nothing
        
        ' Update batch start and counter
        startID = endID + 1
        batchCount = batchCount + 1
        
        ' Progress update
        Debug.Print "Completed batch " & batchCount & " (IDs " & (startID - batchSize) & " to " & endID & ", " & recordCount & " records)"
        DoEvents
    Loop
    
    ' Clean up and restore settings
    DoCmd.SetWarnings True
    Application.Echo True
    Set db = Nothing
    
    MsgBox "Hash computation completed. Processed " & totalRecords & " records in " & batchCount & " batches.", vbInformation
    Exit Sub

ErrHandler:
    DoCmd.SetWarnings True
    Application.Echo True
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Sub UpdateBatchWithTempTable(db As DAO.Database, idArray() As Long, hashArray() As String)
    Dim i As Long
    Dim sql As String
    
    ' Create temporary table
    On Error Resume Next ' Handle if table exists
    db.Execute "DROP TABLE tempHashUpdate"
    On Error GoTo 0
    sql = "CREATE TABLE tempHashUpdate (ID Long, hash Text(40))"
    db.Execute sql
    
    ' Insert ID and hash into temp table
    For i = LBound(idArray) To UBound(idArray)
        sql = "INSERT INTO tempHashUpdate (ID, hash) VALUES (" & idArray(i) & ", '" & Replace(hashArray(i), "'", "''") & "')"
        db.Execute sql, dbFailOnError
    Next i
    
    ' Update main table from temp table
    sql = "UPDATE t_tbl1 INNER JOIN tempHashUpdate ON t_tbl1.ID = tempHashUpdate.ID " & _
          "SET t_tbl1.hash = tempHashUpdate.hash"
    db.Execute sql, dbFailOnError
    
    ' Drop temp table
    db.Execute "DROP TABLE tempHashUpdate"
End Sub