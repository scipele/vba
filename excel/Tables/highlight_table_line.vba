'| Item	        | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | highlight_table_line.vba                                    |
'| EntryPoint   | event worksheet selection change                            |
'| Purpose      | highlight the current line in spreadsheet                   |
'| Inputs       | there is a boolean toggle to turn on/off                    |
'| Outputs      | color change on the row within the table                    |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 11/07/2025                                        |


Option Explicit
'Declare a toggle true/false to confirm whether highlighting is done
Dim highlight_on_bool As Boolean


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    
    'get the toggle value
    highlight_on_bool = [Highlighter_On]
    
    'Stop if highlight is turned off
    If Not highlight_on_bool Then Exit Sub

    Dim tbl As ListObject
    Set tbl = Me.ListObjects("Table1")
    
    'Note that Target is the built in library default selected cell(s) Range object
    'End sub if user selected a cell outside of the table
    If Intersect(Target, tbl.Range) Is Nothing Then Exit Sub

   ' Clear all highlighting
     tbl.Range.Interior.ColorIndex = xlNone

    ' Highlight selected row within table 6=Yellow, 8=Cyan, 7=Magenta, ...
    Intersect(Target.EntireRow, tbl.Range).Interior.ColorIndex = 6
End Sub