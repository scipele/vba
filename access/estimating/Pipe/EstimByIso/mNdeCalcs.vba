'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | mNdeCalcs.vba                                               |
'| EntryPoint   | GetRtTrips and GetPwhtTrips                                 |
'| Purpose      | calculate the number of XRay and PWHT Trips                 |
'| Inputs       | user inputs                                                 |
'| Outputs      | various                                                     |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple / BJR, 10/1/2025                                   |

Option Compare Database

Function GetRtTrips() As String
    
    ' Use the current database session
    Dim db As DAO.Database
    Set db = CurrentDb
    
    ' Define the SQL query with dynamic est_id
    Dim sql As String
    sql = "SELECT " & _
        "Sum(" & _
        "(Nz([butt_wld_qty])+" & _
        "Nz([sw_qty]))" & _
        "*" & _
        "Switch([ta_data].[rt_pct]<=0.05,0.2," & _
        "[ta_data].[rt_pct]<=0.1,0.3," & _
        "[ta_data].[rt_pct]<=0.2,0.4," & _
        "[ta_data].[rt_pct]<=0.5,0.6," & _
        "[ta_data].[rt_pct]<=1,1)/" & _
        "Switch(" & _
        "[tb_qtys].[size_id]<=8,8," & _
        "[tb_qtys].[size_id]<=13,6," & _
        "[tb_qtys].[size_id]<=15,5," & _
        "[tb_qtys].[size_id]<=18,3," & _
        "[tb_qtys].[size_id]<=22,2," & _
        "True,1)) " & _
        "AS rt_trips " & _
        "FROM ta_data " & _
        "INNER JOIN tb_qtys " & _
        "ON ta_data.est_id = tb_qtys.est_id;"
    
    ' Execute the query
    Debug.Print sql
    On Error GoTo ErrorHandler
    
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset(sql)
    
    ' Check if records are returned
    If Not rs.EOF Then
        Dim trips As Double
        trips = rs("rt_trips")
    End If
    
    'Round the number
    trips = Int(-1 * trips) / -1
    
    Dim trips_str As String
    trips_str = CStr(trips)
    
    GetRtTrips = trips_str
    
    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function


Function GetPwhtTrips() As String
    
    On Error GoTo ErrorHandler
    ' Use the current database session
    Dim db As DAO.Database
    Set db = CurrentDb
    
    ' Define the SQL query with dynamic est_id
    Dim sql As String
    sql = "SELECT " & _
          "Sum(IIf([ta_data].[pwht]," & _
          "Nz([butt_wld_qty])+" & _
          "Nz([sw_qty])," & _
          "0)/" & _
          "Switch([tb_qtys].[size_id]<=8,6," & _
          "[tb_qtys].[size_id]<=13,4," & _
          "[tb_qtys].[size_id]<=15,3," & _
          "[tb_qtys].[size_id]<=18,2," & _
          "True,1))" & _
          "AS pwht_trips " & _
          "FROM ta_data " & _
          "INNER JOIN tb_qtys " & _
          "ON ta_data.est_id = tb_qtys.est_id;"
    
    ' Execute the query
    Debug.Print sql
    
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset(sql)
    
    ' Check if records are returned
    If Not rs.EOF Then
        Dim trips As Double
        trips = rs("pwht_trips")
    End If
    
    'Round the number
    trips = Int(-1 * trips) / -1
    
    Dim trips_str As String
    trips_str = CStr(trips)
    
    GetPwhtTrips = trips_str
    
    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    Exit Function
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function
