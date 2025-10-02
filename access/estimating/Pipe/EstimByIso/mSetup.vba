'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | mGetMhs.vba                                                 |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple / BJR, 10/1/2025                                   |

Option Compare Database
Option Explicit


Public Sub DeleteAllRecords()
    On Error GoTo ErrorHandler
    
    ' Array of table names
    Dim tables As Variant
    Dim tableName As Variant
    Dim db As DAO.Database
    
    ' Define the tables to clear
    tables = Array("tb_qtys", "ta_data", "td_areas", "te_specs")
    
    ' Get reference to current database
    Set db = CurrentDb
    
    ' Confirm with user before deletion
    If MsgBox("Are you sure you want to delete all records from ta_data, tb_qtys, td_areas, and te_specs? This cannot be undone.", _
              vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
        
        ' Loop through each table and delete records
        For Each tableName In tables
            DoCmd.SetWarnings False
            DoCmd.RunSQL "DELETE * FROM " & tableName & ";"
            DoCmd.SetWarnings True
            Debug.Print "Records deleted from " & tableName
        Next tableName
        
        MsgBox "All records have been successfully deleted from the specified tables.", vbInformation, "Deletion Complete"
    Else
        MsgBox "Operation cancelled.", vbInformation, "Cancelled"
    End If
    GoTo Cleanup
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    
Cleanup:
    Set db = Nothing
End Sub