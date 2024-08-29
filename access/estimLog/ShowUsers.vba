Option Compare Database
Option Explicit

Sub ShowUserRosterMultipleUsers()
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset

    On Error GoTo ErrorHandler

    Set cn = CurrentProject.Connection

    ' The user roster is exposed as a provider-specific schema rowset
    ' in the Jet 4.0 OLE DB provider. You have to use a GUID to
    ' reference the schema, as provider-specific schemas are not
    ' listed in ADO's type library for schema rowsets

    Set rs = cn.OpenSchema(adSchemaProviderSpecific, _
        , "{947bb102-5d43-11d1-bdbf-00c04fb92675}")


    'Output the list of all users in the current database.

    Debug.Print rs.Fields(0).Name, "", rs.Fields(1).Name, _
    "", rs.Fields(2).Name, rs.Fields(3).Name

    While Not rs.EOF
        Debug.Print rs.Fields(0), rs.Fields(1), _
        rs.Fields(2), rs.Fields(3)
        rs.MoveNext
    Wend

    Exit Sub 'Exit if there are no errors
  
ErrorHandler:
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Err.Clear

End Sub