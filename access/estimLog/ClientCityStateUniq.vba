Option Compare Database
Option Explicit


Public Sub confirmClientCityStateUnique()
    Dim sql As String
    Dim rs As DAO.Recordset
    Dim CompanyLocID As Long
    Dim CompanyID As Long
    Dim CityID As Long
    Dim StateID As Long
    Dim Cancel As Boolean
    
    Forms("frm07ClientCityState").Controls("cboxCompany").SetFocus
    
    'Get Values from Combo Boxes
    CompanyID = Nz(Forms("frm07ClientCityState").Controls("cboxCompany").Value, "")
    CityID = Nz(Forms("frm07ClientCityState").Controls("cboxCityID").Value, "")
    StateID = Nz(Forms("frm07ClientCityState").Controls("cboxStateID").Value, "")

    ' Check if the combination already exists in the database
    sql = "SELECT CompanyID, CityID, StateID FROM tlkpClientCityState " & _
             "WHERE CompanyID = " & CompanyID & _
             "  AND CityID = " & CityID & _
             "  AND StateID = " & StateID

    Set rs = CurrentDb.OpenRecordset(sql)

    If Not rs.EOF Then
        MsgBox "Duplicate combination of Client, City, State already exist.", vbExclamation
        Forms("frm07ClientCityState").SetFocus
        Cancel = True ' Cancel the update
        ' Cancel any changes made to the form and close it
        DoCmd.RunCommand acCmdUndo
    Else
        Call UpdateComboClientCityState
    End If

    rs.Close
    Set rs = Nothing
End Sub


Private Sub UpdateComboClientCityState()
    Dim db As DAO.Database
    Dim CompanyLocID As Long
    Dim CompanyID As Long
    Dim CompanyNameID As Long
    Dim CityID As Long
    Dim StateID As Long
    Dim strCompany As String
    Dim strCity As String
    Dim strState As String
    Dim strClientCityState As String
    Dim sql As String
    Dim rs As Recordset
    Dim recordExists As Boolean
    
    'Temporarily Unlock the Textbox
    Forms("frm07ClientCityState").Controls("tboxClientCityState").Locked = False

    CompanyLocID = Nz(Forms("frm07ClientCityState").Controls("cboxID").Value, "")
    CompanyID = Nz(Forms("frm07ClientCityState").Controls("cboxCompany").Value, "")
    CityID = Nz(Forms("frm07ClientCityState").Controls("cboxCityID").Value, "")
    StateID = Nz(Forms("frm07ClientCityState").Controls("cboxStateID").Value, "")
    
    '1. Get Company Name Text
    sql = "SELECT Company FROM tlkpCompany " & _
         "WHERE CompanyID = " & CompanyID & ";"

    Set rs = CurrentDb.OpenRecordset(sql)
    If Not rs.EOF Then
        strCompany = rs.Fields("Company")
    End If
    rs.Close
    
    '2. Get City Text
    sql = "SELECT City FROM tlkpCity " & _
         "WHERE CityID = " & CityID & ";"

    Set rs = CurrentDb.OpenRecordset(sql)
    If Not rs.EOF Then
        strCity = rs.Fields("City")
    End If
    rs.Close
    
    '3. Get State Text
    sql = "SELECT StateAbbr FROM tlkpState " & _
         "WHERE StateID = " & StateID & ";"

    Set rs = CurrentDb.OpenRecordset(sql)
    If Not rs.EOF Then
        strState = rs.Fields("StateAbbr")
    End If
    rs.Close
    
    strClientCityState = strCompany & ", " & strCity & ", " & strState
    
    '4. Insert Into the Combined Company, City, State
    ' Assuming db is your Database object
    Set db = CurrentDb

    ' Check if the record with the given ID exists
    Forms("frm07ClientCityState").Recordset.Requery
    
    recordExists = DCount("[ID]", "tlkpClientCityState", "[ID] = " & CompanyLocID) > 0
    
    If recordExists Then
        sql = "UPDATE tlkpClientCityState SET ClientCityState = '" & strClientCityState & "' WHERE ID = " & CompanyLocID
    Else
        sql = "INSERT INTO tlkpClientCityState (clientCityState) VALUES ('" & strClientCityState & "');"
    End If
 
    ' Execute the query
    db.Execute sql

    ' Close the database
    Set db = Nothing
    
    Forms("frm07ClientCityState").Controls("tboxClientCityState").Locked = True
    
End Sub