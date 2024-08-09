' Sub Name - strip_formulas_for_client
'
' Purpose:
'   1. Triple Confirmation to make sure the user has backed up the sheet
'   2. Unfilter all sheets
'   3. copy and paste values for all sheets ( in a particular order !!!!!
'       - Sum2, Sum1, SPaint, FPaint, Supt, Pile, Conc, ....
'   4. Delete Rows that are "Un-X'd" and non-zero total in certain sheets  - Orange Tabs
'      Also delete columns to the right
'   5. Delete Columns to the right of I on Prod Tab
'   6. Delete all hidden sheets for (Rate, Sched, Metrics, B&G, Torq)
'
' Dependencies:  None
'
' By:  T.Sciple, 8/8/2024
'
Sub strip_formulas_for_client()
    'Step 1
    Dim response As String
    response = make_sure_before_proceeding()
    If response = "No" Then
        MsgBox ("Canceled by User")
        Exit Sub
    End If
    'Step 2
    Call unhide_sheets_and_clear_filters
    'Step 3
    Call copy_and_paste_values_in_order
    'Step 4
    Call delete_un_xd_and_zero_rows
    'Step 5
    Call delete_colms_other
    'Step 6
    Call delete_listed_sheets
End Sub


'Step 1.1
Function make_sure_before_proceeding()
    ' Ask three different times before proceeding
    Dim msgs As Variant
    Dim tmpResponse As String
    msgs = Array("Do you want to wipe out all formulas in this sheet ?", _
                "Are you sure you want to wipe out all formulas in this sheet ?", _
                "Have You Saved a Backup Copy of this File ?")
    
    Dim msg As Variant
    For Each msg In msgs
        tmpResponse = get_btn_response(msg)
        If tmpResponse = "Yes" Then
            tmpResponse = "Yes"
        Else
            make_sure_before_proceeding = "No"
            Exit Function
        End If
    Next msg
    If tmpResponse = "Yes" Then make_sure_before_proceeding = "Yes"
End Function


'Step 1.2
Function get_btn_response(ByVal msg As String)
    Dim Style As Variant
    Dim Title As String
    Dim response As Variant

    Style = vbYesNo 'Define buttons.
    Title = "MsgBox Demonstration"    ' Define title.
    response = MsgBox(msg, Style, Title)
    If response = vbYes Then    ' User chose Yes.
        get_btn_response = "Yes"
    Else
        get_btn_response = "No"
    End If
End Function


'Step 2
Sub unhide_sheets_and_clear_filters()
    ' Loop thru each sheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVisible     'set the visibility
        
        ' Clear filters but keep auto filters
        If ws.AutoFilterMode Then
            ws.AutoFilter.ShowAllData
        End If
    Next ws
End Sub


'Step 3
Sub copy_and_paste_values_in_order()
    Dim sht_order As Variant
    sht_order = Array("Sum2", "Sum1", "SPaint", "FPaint", "Supt", "Prod", "Pile", "Conc", "PipUG", "Steel", "Equip", "PipShp", _
                      "PipFld", "Insul", "Trace", "FirePrf", "EI", "Bldg", "Demo", "SpSub", "Indir", "Conting", "Owner", "KeyQty")
    
    'Dim variables used in the loop
    Dim ws As Worksheet
    Dim usedRange As Range
    Dim sht As Variant
    
    For Each sht In sht_order
        ' Set the worksheet object
        Set ws = ThisWorkbook.Sheets(sht)
           
        ' Define the used range
        Set usedRange = ws.usedRange
        
        ' Copy the used range
        usedRange.Copy
        
        ' Define the destination range (e.g., starting at cell A1 in a different location)
        ws.Range("A1").PasteSpecial Paste:=xlPasteValues
        
        ' Clear Clipboard (Optional)
        Application.CutCopyMode = False
    Next sht
    
    'cleanup delete ws object
    Set ws = Nothing
End Sub


'Step 4
Sub delete_un_xd_and_zero_rows()
    Dim sht_order As Variant
    sht_order = Array("Pile", "Conc", "PipUG", "Steel", "Equip", "PipShp", "PipFld", "Insul", "Trace", "FirePrf", "SPaint", "FPaint", "EI", "Bldg", "Demo", "SpSub", "Supt", "Indir")
    
    'Dim variables used in the loop
    Dim ws As Worksheet
    Dim usedRange As Range
    Dim sht As Variant
    
    Dim key_colms(2) As Integer
    Dim i As Long
    
    For Each sht In sht_order
        ' Set the worksheet object
        Set ws = ThisWorkbook.Sheets(sht)
            
        ' Define the used range
        Set usedRange = ws.usedRange
        
        'Locate the "X" Colm and "TOTAL"
        key_colms(0) = FindColumnByLabel("X", 1, sht)
        key_colms(1) = FindColumnByLabel("TOTAL", 1, sht)
        
        'delete rows backward from the end
        For i = usedRange.Rows.Count To 2 Step -1
            If Not LCase(usedRange(i, key_colms(0))) = "x" And Abs(usedRange(i, key_colms(1))) < 0.001 Then
                ws.Rows(i).Delete
            End If
        Next i
    
        'delete columns past total to end of used range
        For i = usedRange.Columns.Count To (key_colms(1) + 1) Step -1
            ws.Columns(i).Delete
        Next i
    Next sht
    
    'cleanup delete ws object
    Set ws = Nothing
End Sub


'Step 3
Sub delete_listed_sheets()
    Dim shts As Variant
    shts = Array("Rate", "Sched", "Metrics", "B&G", "Torq")
    
    Application.DisplayAlerts = False
    'Dim variables used in the loop
    Dim ws As Worksheet
    Dim sht As Variant
    For Each sht In shts
        ' Set the worksheet object
        Set ws = ThisWorkbook.Sheets(sht)
        ws.Delete
    Next sht
    
    Application.DisplayAlerts = True
    'cleanup delete ws object
    Set ws = Nothing
End Sub


'Step 5
Sub delete_colms_other()
    
    'Dim variables used in the loop
    Dim ws As Worksheet
    Dim usedRange As Range
    
    Dim sht As String
    sht = "Prod"
    
    ' Set the worksheet object
    Set ws = ThisWorkbook.Sheets(sht)
        
    ' Define the used range
    Set usedRange = ws.usedRange

    'Locate the Next to Last Column
    Dim key_colm As Integer
    key_colm = FindColumnByLabel("CALC RATE $/HR", 1, sht)
    
    'delete columns past total to end of used range
    Dim i As Long
    For i = usedRange.Columns.Count To (key_colm + 2) Step -1
        ws.Columns(i).Delete
    Next i
    
    'cleanup delete ws object
    Set ws = Nothing
End Sub


Function FindColumnByLabel(ByVal label As String, _
                            ByVal searchRow As Long, _
                            ByVal shtName As String) _
                            As Long
    Dim ws As Worksheet
    Dim foundCell As Range

    ' Set the worksheet (adjust "Sheet1" to your worksheet name)
    Set ws = ThisWorkbook.Sheets(shtName)

    ' Initialize the function result
    FindColumnByLabel = -1
    
    ' Search for the label in the specified row
    Set foundCell = ws.Rows(searchRow).Find(What:=label, LookIn:=xlValues, LookAt:=xlWhole)

    ' Check if the cell was found and return the column number
    If Not foundCell Is Nothing Then
        FindColumnByLabel = foundCell.Column
    End If
End Function