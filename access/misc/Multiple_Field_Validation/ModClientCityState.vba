Option Compare Database
Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | ModClientCityState.vba                                      |
'| EntryPoint   | cmdEnterClientCityState, or AfterUpdate Events              |
'| Purpose      | Confirm that a combination of three fields doesnt exist,    |
'|              | confirm no fields are blank, and process cancel button press|
'| Inputs       | all read from Combo Boxes/tables                            |
'| Outputs      | Insert new value into 'ClientCityState' Field               |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 12/13/2024                                        |


Public Sub confirmClientCityStateUnique()

    ' exit the Sub if there are any null values with a message
    If IsNull(Forms("FrmNewClientAdd").Controls("cbxCompany").Value) Or _
       IsNull(Forms("FrmNewClientAdd").Controls("cbxCity").Value) Or _
       IsNull(Forms("FrmNewClientAdd").Controls("cbxState").Value) Then
       MsgBox ("You must complete all three fields 'Client, City, and State " & _
               "before entering a new 'ClientCityState'record")
       Exit Sub
    End If

    ' get Values from Combo Boxes
    Forms("FrmNewClientAdd").Controls("cbxCompany").SetFocus
    Dim client_name As String, city As String, state As String
    client_name = Forms("FrmNewClientAdd").Controls("cbxCompany").Value
    city = Forms("FrmNewClientAdd").Controls("cbxCity").Value
    state = Forms("FrmNewClientAdd").Controls("cbxState").Value
    Dim client_city_state As String
    client_city_state = client_name & ", " & city & ", " & state
    
    ' check if the combination already exists in the 'tblClientCityState' table
    Dim sql As String
    sql = "SELECT ClientCityState " & _
          "FROM tlkpClientCityState " & _
          "WHERE ClientCityState = " & _
          """" & client_city_state & """;"
          
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset(sql)
    If Not rs.EOF Then
        MsgBox "The combined entries for 'client, city & state' " & _
               "already exist therefore it does not need to be added. " & vbCr & _
               "Return to the estimate data entry form and select the combined " & _
               "'ClientCityState' from the combo box " & _
               "entry form", vbExclamation
        Forms("FrmNewClientAdd").SetFocus
    Else
        
        ' Prompt the user to confirm adding the new company name
        Dim msg As String
        msg = "The combined 'Client, City, State' was confirmed to be unique." & vbCr & _
              "Please confirm correct spelling and that you want to add this " & vbCr & _
              "combined data to the database " & vbCr & _
              client_city_state & vbCr
        
        Dim title As String
        title = "User Confirmation"
        
        Dim response As VbMsgBoxResult
        response = MsgBox(msg, vbOKCancel, title)
        
        If response = vbOK Then
            ' Enter combined 'client, city, state' into the table 'tblClientCityState'
            sql = "INSERT INTO tlkpClientCityState " & _
                  "(clientCityState) VALUES  " & _
                  "('" & client_city_state & "');"
            
            CurrentDb.Execute (sql)
        Else
            MsgBox ("user canceled the entry")
        End If
    End If
    
    ' close the previous recordset
    rs.Close
    Set rs = Nothing
    ' close the form then reopen
    DoCmd.Close acForm, "FrmNewClientAdd", acSaveYes
    DoCmd.OpenForm "FrmNewClientAdd"
End Sub


Public Sub AddNewTableData(ByVal TableName As String, _
                           ByVal FieldName As String, _
                           ByVal NewData As String)
                           
    ' exit this sub and just leave the value in the combo box if it already exist
    ' as a record in the table
    Dim sql As String
    sql = "SELECT " & FieldName & _
          " FROM " & TableName & _
          " WHERE " & FieldName & "= """ & NewData & """;"
          
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset(sql)

    If Not rs.EOF Then
        Exit Sub
    Else
        ' prompt the user to confirm adding the new company name
        Dim msg As String
        msg = "Are you sure that you want to add this new data - " & _
        "confirm exact spelling?"
        
        Dim title As String
        title = "User Confirmation"
        
        Dim response As VbMsgBoxResult
        response = MsgBox(msg, vbOKCancel, title)
        
        If response = vbOK Then
            ' develop a new sql statement to enter a new value into the table
            sql = "INSERT INTO " & _
                   TableName & _
                   "(" & FieldName & _
                   ") VALUES ('" _
                   & NewData & "');"
            ' execute the sql statement
            CurrentDb.Execute sql
        Else
            ' convert field name to combo box name
            Dim cbx_name As String
            
            cbx_name = "cbx" & _
                        UCase(Left(FieldName, 1)) & _
                        Right(FieldName, Len(FieldName) - 1)
            
            Forms("FrmNewClientAdd").Controls(cbx_name).Value = ""
        End If
    End If
End Sub