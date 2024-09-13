Option Explicit

Sub UnhideAllSheets()
    Dim ws As Worksheet
    Dim activeWb As Workbook
    
    ' Reference the active workbook
    Set activeWb = ActiveWorkbook
    
    ' Loop through each worksheet in the active workbook
    For Each ws In activeWb.Worksheets
        ' Unhide the sheet if it's hidden
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
        End If
    Next ws
    ' Notify the user
    MsgBox "All sheets have been unhidden!", vbInformation
End Sub


Sub UnhideAndClearFilters()
    Dim ws As Worksheet
    Dim activeWb As Workbook
    Dim filterCol As Integer
    
    ' Reference the active workbook
    Set activeWb = ActiveWorkbook
    
    ' Loop through each worksheet in the active workbook
    For Each ws In activeWb.Worksheets
        ' Unhide the sheet if it's hidden
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
        End If
        ' Check if the sheet has a filter applied
        If ws.AutoFilterMode Then
            ' Clear the filter for each column if AutoFilter is applied
            With ws.AutoFilter
                For filterCol = 1 To .Filters.Count
                    If .Filters(filterCol).On Then
                        ws.Cells.AutoFilter Field:=filterCol
                    End If
                Next filterCol
            End With
        End If
    Next ws
    ' Notify the user
    MsgBox "All sheets have been unhidden and filters cleared!", vbInformation
End Sub


Sub Remove_Hidden_Names()
    Dim xName As Variant
    If MsgBox("Do you really want to delete all hidden names in this workbook?", vbQuestion + vbYesNoCancel, "Delete hidden names?") = vbYes Then
        For Each xName In ActiveWorkbook.Names
            If xName.Visible = False Then
                xName.Delete
            End If
        Next xName
        MsgBox "All hidden names in this workbook have been deleted.", vbInformation + vbOKOnly, "Hidden names deleted"
    End If
End Sub