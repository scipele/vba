Option Compare Database
Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | EstimatePendingBacklog.vba                                  |
'| EntryPoint   | cmdGetEstimProjections_Click() calls Public Main()          |
'| Purpose      | Comput Backlog for Jobs and Estimates based on Probabilities|
'| Inputs       | Various Table Data in Access Database                       |
'| Outputs      | code updates Tables -> tbl3Backlog, tbl4BacklogSum          |
'| Dependencies | Microsoft Office 16.0 Access Database endgine Object Library|
'| By Name,Date | T.Sciple, 12/10/2024                                        |


'0. Main Sub
Public Sub Main()

    Dim estBacklog As Variant
    
    '1. This Function Reads the query "qry02Projections" Estimate Projection Values and Separates the Records into by Division ID
    estBacklog = getEstPendingItems()
        
    '2.  This Function separates the estimate Records by Division ID
    Dim estBacklogByDiv As Variant
    estBacklogByDiv = getSeparateRecordsByDivision(estBacklog)
    Erase estBacklog
    
    '3. This Sub deletes the "tbl2BacklogEstim" so that it can be redefined
    Call DeleteIfExists("tbl2BacklogEstim")
    
    '4. This Sub creates a New Blank Table "tbl2BacklogEstim" where the Values will be placed
    Call createNewTable1
    
    '5. This Sub adds records and data to "tbl2BacklogEstim" based on the data contained in the Array "estBacklogByDiv" passed to the Sub
    Call AddRecordsTotbl2(estBacklogByDiv)
    Erase estBacklogByDiv
    
    '6. This Function Reads the query "qry05CurJobEstimBacklogRecap" Estimate/Job Backlog Values into a temporary Arry then saves it to the Function Name
    Dim BacklogData As Variant
    BacklogData = getEstJobBacklog()
    
    '7. This Sub deletes the "tbl3Backlog" so that it can be redefined
    Call DeleteIfExists("tbl3Backlog")
    
    '8. This Sub creates a New Blank Table "tbl3Backlog" where the Values will be placed
    Dim MaxMonths As Integer
    MaxMonths = 24
    Call createNewTable2(MaxMonths)
    
    '9. This Sub adds records and data to "tbl3Backlog" based on the data contained in the temporary Array passed to the Sub
    Call AddRecords(BacklogData, MaxMonths)
    
    '10. This Sub Calculates the Spread of the Projected Values throughout the months of the backlog
    Call DistributeMonths(BacklogData, MaxMonths)
    
    '11. This Function iterates through each estimate/job record and then each of the 24 months and Sums up the Projected Values for each month.  The Sum Values are recorded for Job and Estimates Separately
    Dim sumMoValues As Variant
    sumMoValues = getSummaryData(BacklogData, MaxMonths)
    
    '12. This Sub Updates "tbl2BacklogSum" based on passed tmpSumValues Array
    Call UpdTbl1(sumMoValues, MaxMonths)
    
    MsgBox ("Completed all Calculations")
End Sub


'1. This Function Reads the linked table "tblEstimData" data into an array
Private Function getEstPendingItems()

    Dim numRecords, i As Long
    Dim rs As Object
    Set rs = CurrentDb.OpenRecordset("qry01Pending")
    Dim tempAry As Variant
    Dim cnt As Long
    numRecords = rs.RecordCount
        
    ReDim tempAry(1 To numRecords, 1 To 14)
    
    'first count the records have a status of "Won" which is id 4
    cnt = 0
    For i = 1 To numRecords
        If rs!Status = 4 Then cnt = cnt + 1
    Next i

    For i = 1 To numRecords
        tempAry(i, 1) = "E"
        tempAry(i, 2) = rs!EstimNo
        tempAry(i, 3) = rs!ClientCityState
        tempAry(i, 4) = rs!Title
        tempAry(i, 5) = rs!Pct
        tempAry(i, 6) = rs!BidDate
        tempAry(i, 7) = rs!StartDate
        tempAry(i, 8) = rs!FinDate
        tempAry(i, 9) = rs!ShopAmt
        tempAry(i, 10) = rs!PaintAmt
        tempAry(i, 11) = rs!MechAmt
        tempAry(i, 12) = rs!EIAmt
        tempAry(i, 13) = rs!ArchAmt
        tempAry(i, 14) = rs!SoftCraftAmt
        
        rs.MoveNext
    Next i

    getEstPendingItems = tempAry
    Erase tempAry
End Function


'2. This Sub creates and Array of Estimate Projection Values that are separated by Division / Cost Center
Private Function getSeparateRecordsByDivision(ByRef tempAry As Variant) As Variant

    Dim tmpArySeparated As Variant
    Dim i, j, k, m, cnt As Long
    Dim tempCnt, tempCntB  As Integer
    
    'Hold off until non Zero Values are Counted
    
    'Run through the Array and Count non Zero Values so that the required Array Size Can be determined
    For i = 1 To UBound(tempAry)
        For j = 9 To 14
            If tempAry(i, j) > 0 Then cnt = cnt + 1
        Next j
    Next i

    ReDim tmpArySeparated(1 To cnt, 1 To 10)

    'i counter is the counter associated with the tmpAry
    For i = 1 To UBound(tempAry)
        
    'First determine the quantity of fields that contain nonZero Estimate Values
        tempCnt = 0
        For j = 9 To 14
            If tempAry(i, j) > 0 Then tempCnt = tempCnt + 1
        Next j
        
        If tempCnt > 0 Then
            For k = 1 To tempCnt
                 m = m + 1
                'The first 8 array elements are simplied copied to each of the records that will be separated
                For j = 1 To 8
                    tmpArySeparated(m, j) = tempAry(i, j)
                Next j
                'This one is tricky because we want to assign the nonZero Value Depending on which count we are on
                tempCntB = 0
                For j = 9 To 14
                    If tempAry(i, j) > 0 Then
                        tempCntB = tempCntB + 1
                        If k = tempCntB Then
                            tmpArySeparated(m, 9) = tempAry(i, j)
                            'Now assign the Division ID as j less 8, which will be stored in the 10th array element
                            tmpArySeparated(m, 10) = j - 8
                        End If
                        
                    End If
                Next j
            Next k
        End If
    Next i

    getSeparateRecordsByDivision = tmpArySeparated
    Erase tmpArySeparated
End Function


'3.  Next Delete the "tbl2BacklogEstim" if it Exists (Generic Function below is Called from the Form Code
'4. This Sub creates a New Blank Table "tbl2BacklogEstim" where the Values will be placed
Sub createNewTable1()

    'This Sub creates a New Blank Table tbl4BacklogSum where the Values will be placed
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fldID, fldEorJ, fldEstJobNo, fldClientLoc, fldJobEstDesc, fldPct, fldBidDate, fldStartDate, fldFinDate, fldTotalAmt, fldProjAmt, fldDurMo, fldDivID As DAO.Field
    Dim moYr As Variant
    Dim i As Integer
    
    Set db = CurrentDb
    
    'create the table definition
    Set tdf = db.CreateTableDef("tbl2BacklogEstim")
    
    'create the field definitions
    Set fldID = tdf.CreateField("ID", dbLong): fldID.Attributes = dbAutoIncrField: fldID.Required = True
    Set fldEorJ = tdf.CreateField("EorJ", dbText): fldEorJ.Required = True
    Set fldEstJobNo = tdf.CreateField("EstJobNo", dbText): fldEstJobNo.Required = True
    Set fldDivID = tdf.CreateField("DivID", dbLong): fldDivID.Required = True
    Set fldClientLoc = tdf.CreateField("ClientLoc", dbText): fldClientLoc.Required = True
    Set fldJobEstDesc = tdf.CreateField("JobEstDesc", dbText): fldJobEstDesc.Required = True
    Set fldPct = tdf.CreateField("Pct", dbDouble): fldPct.Required = True
    Set fldBidDate = tdf.CreateField("BidDate", dbText): fldBidDate.Required = True
    Set fldStartDate = tdf.CreateField("StartDate", dbDate): fldStartDate.Required = True
    Set fldFinDate = tdf.CreateField("FinDate", dbDate): fldFinDate.Required = True
    Set fldTotalAmt = tdf.CreateField("TotalAmt", dbDouble): fldTotalAmt.Required = True
    Set fldProjAmt = tdf.CreateField("ProjAmt", dbDouble): fldProjAmt.Required = True
    Set fldDurMo = tdf.CreateField("DurMo", dbDouble): fldDurMo.Required = True
    
    'Add the fields to the table
    tdf.Fields.Append fldID
    tdf.Fields.Append fldEorJ
    tdf.Fields.Append fldEstJobNo
    tdf.Fields.Append fldDivID
    tdf.Fields.Append fldClientLoc
    tdf.Fields.Append fldJobEstDesc
    tdf.Fields.Append fldPct
    tdf.Fields.Append fldBidDate
    tdf.Fields.Append fldStartDate
    tdf.Fields.Append fldFinDate
    tdf.Fields.Append fldTotalAmt
    tdf.Fields.Append fldProjAmt
    tdf.Fields.Append fldDurMo
     
    'add the table to the database
    db.TableDefs.Append tdf
    
    'refresh the tables and the database
    db.TableDefs.Refresh
    Application.RefreshDatabaseWindow
    
    'unload variables
    Set fldID = Nothing
    Set fldEstJobNo = Nothing
    Set fldClientLoc = Nothing
    Set fldJobEstDesc = Nothing
    Set fldPct = Nothing
    Set fldBidDate = Nothing
    Set fldStartDate = Nothing
    Set fldFinDate = Nothing
    Set fldTotalAmt = Nothing
    Set fldProjAmt = Nothing
    Set fldDurMo = Nothing
    Set fldDivID = Nothing
End Sub


'5. This Sub adds records and data to "tbl2BacklogEstim" based on the data contained in the tempAry passed to the Sub
Sub AddRecordsTotbl2(ByRef tempAry As Variant)

    Dim db As DAO.Database
    Dim rs As Object
    Dim i, j As Integer
    
    Set db = CurrentDb
    Set rs = CurrentDb.OpenRecordset("tbl2BacklogEstim")
    
    On Error Resume Next
        
    With rs
        For i = LBound(tempAry, 1) To UBound(tempAry, 1)
            .AddNew
            If IsNull(tempAry(i, 1)) = True Then !EorJ = "" Else !EorJ = tempAry(i, 1)
            If IsNull(tempAry(i, 2)) = True Then !EstJobNo = "" Else !EstJobNo = tempAry(i, 2)
            If IsNull(tempAry(i, 10)) = True Then !DivID = 0 Else !DivID = tempAry(i, 10)
            If IsNull(tempAry(i, 3)) = True Then !ClientLoc = "" Else !ClientLoc = tempAry(i, 3)
            If IsNull(tempAry(i, 4)) = True Then !JobEstDesc = "" Else !JobEstDesc = tempAry(i, 4)
            If IsNull(tempAry(i, 5)) = True Then !Pct = "0" Else !Pct = tempAry(i, 5)
            If IsNull(tempAry(i, 6)) = True Then !BidDate = VBA.Date Else !BidDate = tempAry(i, 6)
            If IsNull(tempAry(i, 7)) = True Then !StartDate = VBA.Date Else !StartDate = tempAry(i, 7)
            If IsNull(tempAry(i, 8)) = True Then !FinDate = VBA.Date Else !FinDate = tempAry(i, 8)
            If IsNull(tempAry(i, 9)) = True Then !TotalAmt = 0 Else !TotalAmt = tempAry(i, 9)
            If IsNull(tempAry(i, 9)) = True Then !ProjAmt = 0 Else !ProjAmt = tempAry(i, 5) * tempAry(i, 9)
            !DurMo = 0
            .Update
        Next i
    End With
    
    Erase tempAry
End Sub


'6. This Function Reads the query "qry05CurJobEstimBacklogRecap" Estimate/Job Backlog Values into a temporary Arry then saves it to the Function Name
Private Function getEstJobBacklog()
    Dim numRecords, i As Long
    Dim rs As Object
    Set rs = CurrentDb.OpenRecordset("qry05CurJobEstimBacklogRecap")
    Dim tempAry As Variant
    numRecords = rs.RecordCount
        
    ReDim tempAry(1 To numRecords, 1 To 11)
    
    For i = 1 To numRecords
        tempAry(i, 1) = rs!EorJ
        tempAry(i, 2) = rs!JobEstNo
        tempAry(i, 3) = rs!DivID
        tempAry(i, 4) = rs!ClientLoc
        tempAry(i, 5) = rs!JobEstDesc
        tempAry(i, 6) = rs!Pct
        tempAry(i, 7) = rs!BidDate
        tempAry(i, 8) = rs!Start
        tempAry(i, 9) = rs!Finish
        tempAry(i, 10) = rs!TotalAmt
        tempAry(i, 11) = rs!BackLog
        rs.MoveNext
    Next i

    getEstJobBacklog = tempAry
    
    Erase tempAry
End Function


'7.  Next Delete the "tbl3Backlog" if it Exists (Generic Function below is Called from the Form Code
'8. This Sub creates a New Blank Table "tbl3Backlog" where the Values will be placed
Sub createNewTable2(ByVal MaxMonths As Integer)

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fldID, fldEorJ, fldEstJobNo, fldDivID, fldClientLoc, fldJobEstDesc, fldPct, fldBidDate, fldStartDate, fldFinDate, fldTotalAmt, fldProjAmt, fldDurMo, fldMo1, fldMo2, fldMo3, fldMo4, fldMo5, fldMo6, fldMo7, fldMo8, fldMo9, fldMo10, fldMo11, fldMo12, fldMo13, fldMo14, fldMo15, fldMo16, fldMo17, fldMo18, fldMo19, fldMo20, fldMo21, fldMo22, fldMo23, fldMo24 As DAO.Field
    Dim moYr As Variant
    Dim i As Integer
    
    Set db = CurrentDb
    
    'create the table definition
    Set tdf = db.CreateTableDef("tbl3Backlog")
    
    'create the field definitions
    Set fldID = tdf.CreateField("ID", dbLong): fldID.Attributes = dbAutoIncrField: fldID.Required = True
    Set fldEorJ = tdf.CreateField("EorJ", dbText): fldEorJ.Required = True
    Set fldEstJobNo = tdf.CreateField("EstJobNo", dbText): fldEstJobNo.Required = True
    Set fldDivID = tdf.CreateField("DivID", dbLong): fldDivID.Required = True
    Set fldClientLoc = tdf.CreateField("ClientLoc", dbText): fldClientLoc.Required = True
    Set fldJobEstDesc = tdf.CreateField("JobEstDesc", dbText): fldJobEstDesc.Required = True
    Set fldPct = tdf.CreateField("Pct", dbDouble): fldPct.Required = True
    Set fldBidDate = tdf.CreateField("BidDate", dbText): fldBidDate.Required = True
    Set fldStartDate = tdf.CreateField("StartDate", dbDate): fldStartDate.Required = True
    Set fldFinDate = tdf.CreateField("FinDate", dbDate): fldFinDate.Required = True
    Set fldTotalAmt = tdf.CreateField("TotalAmt", dbDouble): fldTotalAmt.Required = True
    Set fldProjAmt = tdf.CreateField("ProjAmt", dbDouble): fldProjAmt.Required = True
    Set fldDurMo = tdf.CreateField("DurMo", dbDouble): fldDurMo.Required = True
    
    'get next 24 months txt name from Function
    moYr = getMoYr(MaxMonths)
   
    Set fldMo1 = tdf.CreateField(moYr(1), dbDouble): fldMo1.Required = True
    Set fldMo2 = tdf.CreateField(moYr(2), dbDouble): fldMo2.Required = True
    Set fldMo3 = tdf.CreateField(moYr(3), dbDouble): fldMo3.Required = True
    Set fldMo4 = tdf.CreateField(moYr(4), dbDouble): fldMo4.Required = True
    Set fldMo5 = tdf.CreateField(moYr(5), dbDouble): fldMo5.Required = True
    Set fldMo6 = tdf.CreateField(moYr(6), dbDouble): fldMo6.Required = True
    Set fldMo7 = tdf.CreateField(moYr(7), dbDouble): fldMo7.Required = True
    Set fldMo8 = tdf.CreateField(moYr(8), dbDouble): fldMo8.Required = True
    Set fldMo9 = tdf.CreateField(moYr(9), dbDouble): fldMo9.Required = True
    Set fldMo10 = tdf.CreateField(moYr(10), dbDouble): fldMo10.Required = True
    Set fldMo11 = tdf.CreateField(moYr(11), dbDouble): fldMo11.Required = True
    Set fldMo12 = tdf.CreateField(moYr(12), dbDouble): fldMo12.Required = True
    Set fldMo13 = tdf.CreateField(moYr(13), dbDouble): fldMo13.Required = True
    Set fldMo14 = tdf.CreateField(moYr(14), dbDouble): fldMo14.Required = True
    Set fldMo15 = tdf.CreateField(moYr(15), dbDouble): fldMo15.Required = True
    Set fldMo16 = tdf.CreateField(moYr(16), dbDouble): fldMo16.Required = True
    Set fldMo17 = tdf.CreateField(moYr(17), dbDouble): fldMo17.Required = True
    Set fldMo18 = tdf.CreateField(moYr(18), dbDouble): fldMo18.Required = True
    Set fldMo19 = tdf.CreateField(moYr(19), dbDouble): fldMo19.Required = True
    Set fldMo20 = tdf.CreateField(moYr(20), dbDouble): fldMo20.Required = True
    Set fldMo21 = tdf.CreateField(moYr(21), dbDouble): fldMo21.Required = True
    Set fldMo22 = tdf.CreateField(moYr(22), dbDouble): fldMo22.Required = True
    Set fldMo23 = tdf.CreateField(moYr(23), dbDouble): fldMo23.Required = True
    Set fldMo24 = tdf.CreateField(moYr(24), dbDouble): fldMo24.Required = True
    
    'Add the fields to the table
    tdf.Fields.Append fldID
    tdf.Fields.Append fldEorJ
    tdf.Fields.Append fldEstJobNo
    tdf.Fields.Append fldDivID
    tdf.Fields.Append fldClientLoc
    tdf.Fields.Append fldJobEstDesc
    tdf.Fields.Append fldPct
    tdf.Fields.Append fldBidDate
    tdf.Fields.Append fldStartDate
    tdf.Fields.Append fldFinDate
    tdf.Fields.Append fldTotalAmt
    tdf.Fields.Append fldProjAmt
    tdf.Fields.Append fldDurMo
    tdf.Fields.Append fldMo1
    tdf.Fields.Append fldMo2
    tdf.Fields.Append fldMo3
    tdf.Fields.Append fldMo4
    tdf.Fields.Append fldMo5
    tdf.Fields.Append fldMo6
    tdf.Fields.Append fldMo7
    tdf.Fields.Append fldMo8
    tdf.Fields.Append fldMo9
    tdf.Fields.Append fldMo10
    tdf.Fields.Append fldMo11
    tdf.Fields.Append fldMo12
    tdf.Fields.Append fldMo13
    tdf.Fields.Append fldMo14
    tdf.Fields.Append fldMo15
    tdf.Fields.Append fldMo16
    tdf.Fields.Append fldMo17
    tdf.Fields.Append fldMo18
    tdf.Fields.Append fldMo19
    tdf.Fields.Append fldMo20
    tdf.Fields.Append fldMo21
    tdf.Fields.Append fldMo22
    tdf.Fields.Append fldMo23
    tdf.Fields.Append fldMo24
     
    'add the table to the database
    db.TableDefs.Append tdf
    
    'refresh the tables and the database
    db.TableDefs.Refresh
    Application.RefreshDatabaseWindow
    
    'unload variables
    Set fldID = Nothing
    Set fldEstJobNo = Nothing
    Set fldDivID = Nothing
    Set fldClientLoc = Nothing
    Set fldJobEstDesc = Nothing
    Set fldPct = Nothing
    Set fldBidDate = Nothing
    Set fldStartDate = Nothing
    Set fldFinDate = Nothing
    Set fldTotalAmt = Nothing
    Set fldProjAmt = Nothing
    Set fldDurMo = Nothing
    
End Sub


'9. This Sub adds records and data to "tbl3Backlog" based on the data contained in the tempAry passed to the Sub
Sub AddRecords(ByRef tempAry As Variant, _
               ByVal MaxMonths As Integer)

    Dim db As DAO.Database
    Dim rs As Object
    Dim i, j As Integer
    Dim moYr As Variant
    
    Set db = CurrentDb
    Set rs = CurrentDb.OpenRecordset("tbl3Backlog")
    
    On Error Resume Next
        
    With rs
        For i = LBound(tempAry, 1) To UBound(tempAry, 1)
            .AddNew
            If IsNull(tempAry(i, 1)) = True Then !EorJ = "" Else !EorJ = tempAry(i, 1)
            If IsNull(tempAry(i, 2)) = True Then !EstJobNo = "" Else !EstJobNo = tempAry(i, 2)
            If IsNull(tempAry(i, 3)) = True Then !DivID = "" Else !DivID = tempAry(i, 3)
            If IsNull(tempAry(i, 4)) = True Then !ClientLoc = "" Else rs!ClientLoc = tempAry(i, 4)
            If IsNull(tempAry(i, 5)) = True Then !JobEstDesc = "" Else rs!JobEstDesc = tempAry(i, 5)
            If IsNull(tempAry(i, 6)) = True Then !Pct = "0" Else rs!Pct = tempAry(i, 6)
            If IsNull(tempAry(i, 7)) = True Then !BidDate = VBA.Date Else rs!BidDate = tempAry(i, 7)
            If IsNull(tempAry(i, 8)) = True Then !StartDate = VBA.Date Else rs!StartDate = tempAry(i, 8)
            If IsNull(tempAry(i, 9)) = True Then !FinDate = VBA.Date Else rs!FinDate = tempAry(i, 9)
            If IsNull(tempAry(i, 10)) = True Then !TotalAmt = 0 Else rs!TotalAmt = tempAry(i, 10)
            If IsNull(tempAry(i, 11)) = True Then !ProjAmt = 0 Else rs!ProjAmt = tempAry(i, 11)
            !DurMo = 0
            'rerun Mo Yr Function to get the Field names
            moYr = getMoYr(MaxMonths)
            For j = 1 To MaxMonths
                rs.Fields(moYr(j)) = 0
            Next j
            .Update
        Next i
    End With
End Sub


'10. This Sub Calculates the Spread of the Projected Values throughout the months of the backlog
Private Sub DistributeMonths(ByRef tempAry As Variant, _
                             ByVal MaxMonths As Integer)
    
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim rs As Object
    Set rs = CurrentDb.OpenRecordset("tbl3Backlog")
    
    'rerun Mo Yr Function to get the Field names
    Dim moYr As Variant
    moYr = getMoYr(MaxMonths)
    
    On Error Resume Next
    
    Dim i As Long
    Dim j As Long
    Dim NoMonths As Integer
    Dim startMo As Integer
    Dim maxMo As Integer
    
    With rs
        .MoveFirst
        For i = LBound(tempAry, 1) To UBound(tempAry, 1)
            
            'Compute the duration to spread the estimate Projections
            If IsNull(tempAry(i, 8)) = True Or IsNull(tempAry(i, 9)) = True Then
                NoMonths = Int(-(0.0181 * (Nz(tempAry(i, 9))) ^ 0.3598)) / -1
            Else
                NoMonths = DateDiff("m", tempAry(i, 8), tempAry(i, 9))
            End If
            
            If NoMonths < 1 Then NoMonths = 1
            
            'Determine start Month-Yr as integer
            startMo = DateDiff("m", VBA.Date, Nz(tempAry(i, 8))) + 1
            If startMo < 1 Then
                'reduce number of months for projects that have already started
                NoMonths = NoMonths + startMo
                'If computed Number of Months is less than 1 then set it equal to 1 month
                If NoMonths < 1 Then NoMonths = 1
                startMo = 1
            End If
            
            'update records
            .Edit
            
            maxMo = startMo + NoMonths - 1
            If maxMo > NoMonths Then
                maxMo = NoMonths
            End If
            
            For j = startMo To maxMo
                rs.Fields(moYr(j)) = Int(-Nz(tempAry(i, 11)) / NoMonths) / -1
            Next j
            !DurMo = NoMonths
            .Update
            .MoveNext
        Next i
    End With
End Sub


'11. This Function iterates through each estimate/job record and then each of the 24 months and Sums up the Projected Values for each month.  The Sum Values are recorded for Job and Estimates Separately for each Company Division
Private Function getSummaryData(ByRef tempAry As Variant, _
                                ByVal MaxMonths As Integer)

    Dim db As DAO.Database
    Dim rs As Object
    Dim monthSum As Variant
    Dim moYr As Variant
    Dim i As Long, j As Long, k As Long
    
    Set db = CurrentDb
    Set rs = CurrentDb.OpenRecordset("tbl3Backlog")
    
    ReDim monthSum(1 To 24, 1 To 14)
    'First Array Dimension is used to Store the Month 1 to 24 is for the twenty four month periods
    'Second Array Dimension is used to Store the Following
    '  1 to 7 is for job data by Company DivID
    '  8 to 14 is to Store Estimate Data by Company DivID
    
    'rerun Mo Yr Function to get the Field names in an array that is next 24 months in "YY-MM" Format
    moYr = getMoYr(MaxMonths)
    
    On Error Resume Next
    
    With rs
        .MoveFirst
        For i = LBound(tempAry, 1) To UBound(tempAry, 1)
            For j = 1 To 24
                If !EorJ = "J" Then
                    'Now Compute the running Sum of the Month Values
                    If tempAry(i, 3) = 1 Then monthSum(j, 1) = monthSum(j, 1) + rs.Fields(moYr(j))
                    If tempAry(i, 3) = 2 Then monthSum(j, 2) = monthSum(j, 2) + rs.Fields(moYr(j))
                    If tempAry(i, 3) = 3 Then monthSum(j, 3) = monthSum(j, 3) + rs.Fields(moYr(j))
                    If tempAry(i, 3) = 4 Then monthSum(j, 4) = monthSum(j, 4) + rs.Fields(moYr(j))
                    If tempAry(i, 3) = 5 Then monthSum(j, 5) = monthSum(j, 5) + rs.Fields(moYr(j))
                    If tempAry(i, 3) = 6 Then monthSum(j, 6) = monthSum(j, 6) + rs.Fields(moYr(j))
                End If
                
                If !EorJ = "E" Then
                    If tempAry(i, 3) = 1 Then monthSum(j, 8) = monthSum(j, 8) + rs.Fields(moYr(j))
                    If tempAry(i, 3) = 2 Then monthSum(j, 9) = monthSum(j, 9) + rs.Fields(moYr(j))
                    If tempAry(i, 3) = 3 Then monthSum(j, 10) = monthSum(j, 10) + rs.Fields(moYr(j))
                    If tempAry(i, 3) = 4 Then monthSum(j, 11) = monthSum(j, 11) + rs.Fields(moYr(j))
                    If tempAry(i, 3) = 5 Then monthSum(j, 12) = monthSum(j, 12) + rs.Fields(moYr(j))
                    If tempAry(i, 3) = 6 Then monthSum(j, 13) = monthSum(j, 13) + rs.Fields(moYr(j))
                End If
            Next j
            .MoveNext
        Next i
        
        'next roundup all the values to whole numbers
        For j = 1 To 24
            For k = 1 To 14
                monthSum(j, k) = Int(-monthSum(j, k)) / -1
            Next k
        Next j
        
    End With
    
    getSummaryData = monthSum

End Function


'12. This Sub Updates "tbl4BacklogSum" based on passed tmpSumValues Array
'tmpSumValues sumMoValues
Sub UpdTbl1(ByRef sumMoValues As Variant, _
            ByVal MaxMonths As Integer)

    Dim db As DAO.Database
    Dim rs As Object
    Dim moYr As Variant
    Dim i As Integer
    
    Set db = CurrentDb
    Set rs = CurrentDb.OpenRecordset("tbl4BacklogSum")
    
    moYr = getMoYr(MaxMonths)
    
    With rs
        .MoveFirst
        For i = 1 To MaxMonths
            .Edit
            !MonthYr = moYr(i)
            
            'Active Job Backlog Values
            If sumMoValues(i, 1) = Null Then !JobVal1 = 0 Else !JobVal1 = sumMoValues(i, 1)
            If sumMoValues(i, 2) = Null Then !JobVal2 = 0 Else !JobVal2 = sumMoValues(i, 2)
            If sumMoValues(i, 3) = Null Then !JobVal3 = 0 Else !JobVal3 = sumMoValues(i, 3)
            If sumMoValues(i, 4) = Null Then !JobVal4 = 0 Else !JobVal4 = sumMoValues(i, 4)
            If sumMoValues(i, 5) = Null Then !JobVal5 = 0 Else !JobVal5 = sumMoValues(i, 5)
            If sumMoValues(i, 6) = Null Then !JobVal6 = 0 Else !JobVal6 = sumMoValues(i, 6)
            'Computed Overall Job Totals
            !JobVal7 = sumMoValues(i, 1) + sumMoValues(i, 2) + sumMoValues(i, 3) + sumMoValues(i, 4) + sumMoValues(i, 5) + sumMoValues(i, 6)
            
            'Pending Estimate Backlog Values
            If sumMoValues(i, 8) = Null Then !EstVal1 = 0 Else !EstVal1 = sumMoValues(i, 8)
            If sumMoValues(i, 9) = Null Then !EstVal2 = 0 Else !EstVal2 = sumMoValues(i, 9)
            If sumMoValues(i, 10) = Null Then !EstVal3 = 0 Else !EstVal3 = sumMoValues(i, 10)
            If sumMoValues(i, 11) = Null Then !EstVal4 = 0 Else !EstVal4 = sumMoValues(i, 11)
            If sumMoValues(i, 12) = Null Then !EstVal5 = 0 Else !EstVal5 = sumMoValues(i, 12)
            If sumMoValues(i, 13) = Null Then !EstVal6 = 0 Else !EstVal6 = sumMoValues(i, 13)
            
            'Computed Overall Estimate Totals
            !EstVal7 = sumMoValues(i, 8) + sumMoValues(i, 9) + sumMoValues(i, 10) + sumMoValues(i, 11) + sumMoValues(i, 12) + sumMoValues(i, 13)
            
            .Update
            .MoveNext
        Next i
    End With

End Sub

'.20 This Function is called by Various Subs above and is used to return an Array of 24 Months from Current Date
'    The Date Format is set as "YY-MM" so that it will Sort in Order

Private Function getMoYr(ByVal MaxMonths As Integer) As Variant

    Dim aryMoYr As Variant
    Dim i As Long
    Dim tmpMonthNo, tmpYr As Integer
    Dim moStr, tmpMonthTxt As String
    Dim moTxtAry  As Variant
    
    moStr = ("01,02,03,04,05,06,07,08,09,10,11,12")
    moTxtAry = Split(moStr, ",")
        
    ReDim aryMoYr(1 To MaxMonths)
    tmpMonthNo = Month(VBA.Date)
    tmpMonthTxt = moTxtAry(tmpMonthNo - 1)
    tmpYr = CInt(Right(Str(Year(VBA.Date)), 2))
    
    For i = 1 To MaxMonths
        aryMoYr(i) = Trim(Str(tmpYr)) & "-" & tmpMonthTxt
        tmpMonthNo = tmpMonthNo + 1
        If tmpMonthNo = 13 Then
            tmpMonthNo = 1  ' reset month no back to january
            tmpYr = tmpYr + 1
        End If
        tmpMonthTxt = moTxtAry(tmpMonthNo - 1)
    Next i
    
    getMoYr = aryMoYr
End Function


'21 This Function is called from the Form Code directly and deletes tables if they exist
Private Sub DeleteIfExists(tableName)

    If Not IsNull(DLookup("Name", "MSysObjects", "Name='" & tableName & "'")) Then
        DoCmd.SetWarnings False
        DoCmd.Close acTable, tableName, acSaveYes
        DoCmd.DeleteObject acTable = acDefault, tableName
        Debug.Print "Table" & tableName & "deleted..."
        DoCmd.SetWarnings True
    End If

End Sub
