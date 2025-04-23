Option Compare Database
Option Explicit
'| Item	        | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | UpdatePricing.vba                                           |
'| EntryPoint   | UpdateRdsPricing                                            |
'| Purpose      | Create temporary table, then run update query w/ temp data  |                   
'| Inputs       | none                                                        |                                                                         
'| Outputs      | field 'rds_pricing.price_unit' update w/ new prices         |                                                           
'| Dependencies | Microsoft Office 16.0 Access database engine Object Library |                                     
'| By Name,Date | T.Sciple, 11/30/2024                                        |                                                           

Sub UpdateRdsPricing()
    ' Declare the DAO Database object
    Dim db As DAO.Database

    ' Set the current database
    Set db = CurrentDb

    ' Step 1: Create temporary table (if it doesn't already exist)
    On Error Resume Next ' Ignore error if table exists
    db.Execute "CREATE TABLE temp_data (data_id Long, price_unit DOUBLE)"
    On Error GoTo 0 ' Reset error handling
    
    ' Step 2: Insert data from the linked Product_data table into the temporary table
    db.Execute "INSERT INTO temp_data (data_id, price_unit) SELECT data_id, price_unit FROM Product_data"
    
    ' Capture the number of records affected by the UPDATE query
    Dim recordsUpdated As Long
    recordsUpdated = db.RecordsAffected
    
    ' Step 3: Update rds_pricing table from the temporary table
    db.Execute "UPDATE temp_data INNER JOIN rds_pricing ON temp_data.data_id = rds_pricing.data_id SET rds_pricing.price_unit = [temp_data].[price_unit]"

    ' Step 4: Drop the temporary table
    db.Execute "DROP TABLE temp_data"

    ' Show the number of records updated in a message box
    MsgBox recordsUpdated & " records were updated.", vbInformation, "Update Complete"
End Sub