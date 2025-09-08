'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | mHashChunks.vba                                             |
'| EntryPoint   | ProcessStringsInChunks                                      |
'| Purpose      | Compute Sha1 Hash in chunks                                 |
'| Inputs       | cell values in column b                                     |
'| Outputs      | SHA-1 Hash lowercase hexidecimal represnetation of input str|
'| Dependencies | other module - mSha1Hash                                    |
'| By Name,Date | T.Sciple, 9/8/2025                                          |

Option Explicit

Sub ProcessStringsInChunks()
    Dim startRow As Long
    Dim endRow As Long
    Dim cell As Range
    
    Dim start_time As Double
    start_time = Timer
    
    ' Set worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("sht1")
    
    ' Find last row in column B
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Define chunk size
    Dim chunkSize As Long
    chunkSize = 5000
    
    ' Start processing from row 3
    startRow = 3
    
    ' Loop through chunks
    Dim i As Long
    While startRow <= lastRow
        ' Calculate end row for current chunk
        endRow = WorksheetFunction.Min(startRow + chunkSize - 1, lastRow)
        
        ' Set range for current chunk
        Dim inputRange As Range
        Set inputRange = ws.Range("B" & startRow & ":B" & endRow)
        
        ' Process each cell in the chunk
        For Each cell In inputRange
            Dim outputString As String
            outputString = mSha1Hash.GetSha1Hash(cell.Value)
            
            ' Output to column C in the same row
            ws.Cells(cell.Row, "C").Value = outputString
        Next cell
        
        ' Move to next chunk
        startRow = endRow + 1
    Wend
    Dim elapsed_time As Double
    elapsed_time = Round(Timer - start_time, 2)
    
    MsgBox "Processing complete in " & elapsed_time & " seconds"
End Sub