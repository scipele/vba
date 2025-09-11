Sub ListTables()
	Dim db As DAO.Database
	Dim tdf As DAO.TableDef
	Set db = CurrentDb
	For Each tdf In db.TableDefs
	    ' ignore system and temporary tables
	    If Not (tdf.name Like "MSys*" Or tdf.name Like "~*") Then
	        Debug.Print tdf.name
	    End If
	Next
	Set tdf = Nothing
	Set db = Nothing
End Sub


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