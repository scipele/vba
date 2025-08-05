Sub ListTableFields()
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    
    Dim TableName As String
    TableName = "new"
    
    ' Open the current database
    Set db = CurrentDb
    
    ' Get the table definition
    Set tdf = db.TableDefs(TableName)
    
    ' Loop through all fields in the table
    For Each fld In tdf.Fields
        Debug.Print fld.Name
    Next fld
    
    ' Clean up
    Set fld = Nothing
    Set tdf = Nothing
    Set db = Nothing
End Sub