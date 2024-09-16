Option Explicit
' filename:         sql_query.vba
'
' purpose:          queries are calculated from a named range 'myRange' that is read
'                   and output to the specified (top-left) cell
'
' usage:            run OutputQueryResults() Sub
'
' dependencies:     Microsoft Active-X Data Objects 6.1 Library
'
' By:               T.Sciple, 09/16/2024

Sub OutputQueryResults()
    Dim rs As Object ' ADO Recordset object
    Dim conn As Object ' ADO Connection object
    Dim rngOutput As Range ' Output range
    
    ' Set the output range where the results will be copied
    Set rngOutput = ThisWorkbook.Worksheets("Sheet1").Range("g7") 
    
    ' Create a new ADO Connection object
    Set conn = CreateObject("ADODB.Connection")
    
    ' Open a connection to the workbook
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & ";Extended Properties=""Excel 12.0;HDR=YES"";"
    
    ' Create a new ADO Recordset object
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Set the recordset as a result of the named range query; 'myRange' is a named range
    rs.Open "SELECT Area, Priority, SUM(Value) AS TotalValue FROM myRange GROUP BY Area, Priority;", conn
    
    ' Check if the recordset is not empty
    If Not rs.EOF Then
        ' Copy the query results to the output range
        rngOutput.CopyFromRecordset rs
    Else
        MsgBox "No records found."
    End If
    
    ' Close the recordset and connection
    rs.Close
    conn.Close
    
    ' Clean up
    Set rs = Nothing
    Set conn = Nothing
End Sub


Sub OutputQueryResults2()
    Dim rs As Object ' ADO Recordset object
    Dim conn As Object ' ADO Connection object
    Dim rngOutput As Range ' Output range
    
    ' Set the output range where the results will be copied
    Set rngOutput = ThisWorkbook.Worksheets("Sheet1").Range("o7") ' Replace "Sheet2" and the range with your desired worksheet and range
    
    ' Create a new ADO Connection object
    Set conn = CreateObject("ADODB.Connection")
    
    ' Open a connection to the workbook
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & ";Extended Properties=""Excel 12.0;HDR=YES"";"
    
    ' Create a new ADO Recordset object
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Set the recordset as a result of the named range query; 'myRange' is a named range
    rs.Open "SELECT Area, SUM(Value) AS TotalValue FROM myRange GROUP BY Area;", conn
    
    ' Check if the recordset is not empty
    If Not rs.EOF Then
        ' Copy the query results to the output range
        rngOutput.CopyFromRecordset rs
    Else
        MsgBox "No records found."
    End If
    
    ' Close the recordset and connection
    rs.Close
    conn.Close
    
    ' Clean up
    Set rs = Nothing
    Set conn = Nothing
End Sub