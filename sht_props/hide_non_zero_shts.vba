' Subs:     hide_nonzero_sheets
'           unhide_sheets_and_clear_filters
'           filter_sheets_by_x
'
'
' Purpose:  This code looks at the values in a particular estimate sheet and sets whether the visibility of the sheets depending on the values
'           found within the workbook
'
'
' Dependencies:  None - Late Binding is used with the Dictionary Object
'
' By:  T. Sciple, 8/8/2024
'


Sub hide_nonzero_sheets()

    Dim sums As Variant
    sums = ThisWorkbook.Sheets("Sum2").Range("I3:i24")
    
    Dim sht_names As Variant
    sht_names = Array("Pile", "Conc", "PipUG", "Steel", "Equip", "PipShp", "PipFld", "Insul", "Trace", "FirePrf", "SPaint", "FPaint", "EI", "Bldg", "Demo", "ignore", "SpSub", "Supt", "ignore", "ignore", "ignore", "Indir")
    
    'Create Dictionary Object to store the Sheet Name and desired visible state
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Integer
    For i = LBound(sht_names) To UBound(sht_names)
        If Not sht_names(i) = "ignore" Then
            If Abs(sums(i + 1, 1)) > 0.01 Then  'Checks to make sure that the summary value is not zero
                dict.Add sht_names(i), True
            Else
                dict.Add sht_names(i), False
            End If
            'Debug.Print "key = ", sht_names(i), "Item = ", dict(sht_names(i))
        End If
    Next i
    
    'Manually add other sheets that we want to be visible
    dict.Add "Sum1", True
    dict.Add "Sum2", True
    dict.Add "Note", True
    dict.Add "Prod", True
    dict.Add "Owner", True
    dict.Add "Rate", False
    dict.Add "Sched", False
    dict.Add "Metrics", False
    dict.Add "B&G", False
    dict.Add "Torq", False
    
    ' Change Visible state to match the dictionary value from the ws.name
    Dim ws As Variant
    For Each ws In ThisWorkbook.Worksheets
        ' check the desired visible state from the dictionary object
        If dict(ws.Name) Then
            ws.Visible = xlSheetVisible
        Else
            ws.Visible = xlSheetHidden
        End If
    Next ws
End Sub


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


Sub filter_sheets_by_x()
    Dim ws As Worksheet
    Dim filterRange As Range
    Dim firstColumn As Integer
    Dim criteria As String
    
    'Define the names of the sheets that we want to filter in an array
    Dim shts_to_filter As Variant
    shts_to_filter = Array("Pile", "Conc", "PipUG", "Steel", "Equip", "PipShp", "PipFld", "Insul", "Trace", _
                           "Fireprf", "SPaint", "FPaint", "EI", "Bldg", "Demo", "SpSub", "Supt", "Indir")
                           
    'Read the sheet names that we want filtered into a dictionary object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim sht As Variant
    For Each sht In shts_to_filter
        dict.Add sht, Nothing
    Next sht
                    
    'cleanup memory
    Erase shts_to_filter
                    
    ' Define the filter criteria
    criteria = "x"
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
                
        If dict.exists(ws.Name) Then
            On Error Resume Next
            ' Check if sheet has AutoFilter applied
            If ws.AutoFilterMode Then
                ' Remove existing filters
                ws.AutoFilterMode = False
            End If
            
            ' Define the range to filter - Assuming data starts at A1 and goes to the last used cell
            Set filterRange = ws.UsedRange
            
            ' Check if the range has at least one row of data (to prevent errors)
            If filterRange.Rows.Count > 1 Then
                ' Apply filter to the first column
                filterRange.AutoFilter Field:=1, Criteria1:=criteria
            End If
        End If
    Next ws
End Sub