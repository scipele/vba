Attribute VB_Name = "ImportFinalEstimateModule"
Option Explicit

' Entry point macro
Public Sub RunImportFinalEstimate()
    Dim filePath As String
    filePath = "C:\dev\cpp\estim\mag\FINAL_ESTIMATE.md"
    ImportFinalEstimateMarkdown filePath
End Sub

' Main import routine: parses markdown tables and writes to sheets
Public Sub ImportFinalEstimateMarkdown(ByVal filePath As String)
    Dim fnum As Integer
    Dim lineText As String
    Dim inTable As Boolean
    Dim topSection As String
    Dim subSection As String
    Dim ws As Worksheet
    Dim startRow As Long
    Dim currentRow As Long
    Dim rows As Collection

    If Dir(filePath) = "" Then
        MsgBox "File not found: " & filePath, vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Prepare target sheets
    Set ws = GetOrCreateSheet("Equipment Detail")
    ws.Cells.Clear
    Set ws = GetOrCreateSheet("Labor Detail")
    ws.Cells.Clear
    Set ws = GetOrCreateSheet("Bulk Materials")
    ws.Cells.Clear
    Set ws = GetOrCreateSheet("Subcontracts")
    ws.Cells.Clear
    Set ws = GetOrCreateSheet("Markdown Tables")
    ws.Cells.Clear

    fnum = FreeFile
    Open filePath For Input As #fnum

    Set rows = New Collection
    inTable = False
    topSection = ""
    subSection = ""

    Do While Not EOF(fnum)
        Line Input #fnum, lineText
        lineText = Trim(lineText)

        ' Track headers
        If Left$(lineText, 3) = "###" Then
            subSection = Trim$(Mid$(lineText, 4))
        ElseIf Left$(lineText, 2) = "##" Then
            topSection = Trim$(Mid$(lineText, 3))
            subSection = ""
        End If

        ' Skip code fences and non-table separators
        If lineText = "```" Or lineText = "---" Or lineText = "" Then
            If inTable Then
                ' Flush current table to a sheet
                WriteCollectedTable rows, topSection, subSection
                Set rows = New Collection
                inTable = False
            End If
            GoTo ContinueLoop
        End If

        ' Detect markdown table rows
        If Left$(lineText, 1) = "|" Then
            If Not inTable Then
                inTable = True
                Set rows = New Collection
            End If

            ' Skip separator rows like |------|-----|
            If IsSeparatorRow(lineText) Then GoTo ContinueLoop

            ' Collect row
            rows.Add lineText
        Else
            ' If we were in a table and the line stops being a table row, flush it
            If inTable Then
                WriteCollectedTable rows, topSection, subSection
                Set rows = New Collection
                inTable = False
            End If
        End If
ContinueLoop:
    Loop

    ' Flush any remaining table at EOF
    If inTable And rows.Count > 0 Then
        WriteCollectedTable rows, topSection, subSection
    End If

    Close #fnum

    ' Basic autofit formatting on known sheets
    Dim s As Variant
    For Each s In Array("Equipment Detail", "Labor Detail", "Bulk Materials", "Subcontracts", "Markdown Tables")
        Set ws = GetOrCreateSheet(CStr(s))
        On Error Resume Next
        ws.Columns.AutoFit
        On Error GoTo 0
    Next s

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Import complete.", vbInformation
End Sub

' Decide target sheet based on top-level section
Private Function DetermineTargetSheet(ByVal topSection As String) As String
    Dim t As String
    t = UCase$(topSection)
    If InStr(1, t, "DETAILED LINE ITEM ESTIMATE", vbTextCompare) > 0 Then
        DetermineTargetSheet = "Equipment Detail"
    ElseIf InStr(1, t, "DETAILED LABOR ALLOCATION", vbTextCompare) > 0 Then
        DetermineTargetSheet = "Labor Detail"
    ElseIf InStr(1, t, "BULK MATERIALS", vbTextCompare) > 0 Then
        DetermineTargetSheet = "Bulk Materials"
    ElseIf InStr(1, t, "SUBCONTRACT SUMMARY", vbTextCompare) > 0 Then
        DetermineTargetSheet = "Subcontracts"
    Else
        DetermineTargetSheet = "Markdown Tables"
    End If
End Function

' Write the collected markdown table rows to a sheet
Private Sub WriteCollectedTable(ByVal rows As Collection, ByVal topSection As String, ByVal subSection As String)
    Dim ws As Worksheet
    Dim targetName As String
    Dim nextRow As Long
    Dim headerParts() As String
    Dim i As Long

    targetName = DetermineTargetSheet(topSection)
    Set ws = GetOrCreateSheet(targetName)

    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If nextRow > 1 Then nextRow = nextRow + 2 Else nextRow = 1

    ' Title for context
    If Len(topSection) > 0 Then
        ws.Cells(nextRow, 1).Value = topSection
        ws.Cells(nextRow, 1).Font.Bold = True
        nextRow = nextRow + 1
    End If
    If Len(subSection) > 0 Then
        ws.Cells(nextRow, 1).Value = subSection
        ws.Cells(nextRow, 1).Font.Italic = True
        nextRow = nextRow + 1
    End If

    If rows.Count = 0 Then Exit Sub

    ' Write header row first and capture column names
    WriteMarkdownRow ws, nextRow, CStr(rows(1))
    headerParts = ParseParts(CStr(rows(1)))

    ' Write remaining rows and apply formatting rules
    For i = 2 To rows.Count
        WriteMarkdownRow ws, nextRow + (i - 1), CStr(rows(i))
        Call ApplyRowFormatting(ws, nextRow + (i - 1), headerParts)
    Next i
End Sub

' Write a single markdown table row to the given sheet/row
Private Sub WriteMarkdownRow(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal lineText As String)
    Dim parts() As String
    Dim clean As String

    ' Strip leading/trailing |
    clean = lineText
    If Left$(clean, 1) = "|" Then clean = Mid$(clean, 2)
    If Right$(clean, 1) = "|" Then clean = Left$(clean, Len(clean) - 1)

    parts = Split(clean, "|")

    Dim c As Long, v As String
    For c = 0 To UBound(parts)
        v = Trim$(parts(c))
        ' If it looks like a formula (starts with =), set as formula
        If Len(v) > 0 And Left$(v, 1) = "=" Then
            ws.Cells(rowIndex, c + 1).Formula = v
        Else
            ws.Cells(rowIndex, c + 1).Value = RemoveBoldMarkers(v)
            If IsBoldValue(v) Then ws.Cells(rowIndex, c + 1).Font.Bold = True
        End If
    Next c
End Sub

' Parse a markdown table row into parts
Private Function ParseParts(ByVal lineText As String) As String()
    Dim clean As String
    If Left$(lineText, 1) = "|" Then clean = Mid$(lineText, 2) Else clean = lineText
    If Right$(clean, 1) = "|" Then clean = Left$(clean, Len(clean) - 1)
    ParseParts = Split(clean, "|")
End Function

' Detect if a value is wrapped in bold markers **value**
Private Function IsBoldValue(ByVal v As String) As Boolean
    Dim t As String
    t = Trim$(v)
    IsBoldValue = (Left$(t, 2) = "**" And Right$(t, 2) = "**")
End Function

' Remove bold markers if present
Private Function RemoveBoldMarkers(ByVal v As String) As String
    Dim t As String
    t = Trim$(v)
    If Left$(t, 2) = "**" Then t = Mid$(t, 3)
    If Right$(t, 2) = "**" Then t = Left$(t, Len(t) - 2)
    RemoveBoldMarkers = t
End Function

' Apply formatting to a data row based on header parts (currency, totals, bold rows)
Public Sub ApplyRowFormatting(ByVal ws As Worksheet, ByVal rowIndex As Long, ByRef headerParts() As String)
    Dim colCount As Long
    Dim c As Long
    Dim txt As String
    Dim addrStart As String
    Dim addrEnd As String
    Dim totalCol As Long
    Dim titleTxt As String

    colCount = UBound(headerParts) + 1

    ' Currency conversion: $xxxK -> numeric dollars
    For c = 1 To colCount
        txt = CStr(ws.Cells(rowIndex, c).Value)
        If Len(txt) > 0 Then
            If Left$(txt, 1) = "$" And UCase$(Right$(txt, 1)) = "K" Then
                Dim s As String, valK As Double
                s = Replace$(txt, "$", "")
                s = Replace$(s, ",", "")
                s = Left$(s, Len(s) - 1)
                On Error Resume Next
                valK = CDbl(s)
                On Error GoTo 0
                If valK <> 0 Or s = "0" Then
                    ws.Cells(rowIndex, c).Value = valK * 1000
                    ws.Cells(rowIndex, c).NumberFormat = "$#,##0"
                End If
            End If
        End If
    Next c

    ' Compute Row Total if header includes it (Bulk Materials table)
    totalCol = FindHeaderIndex(headerParts, "Row Total")
    If totalCol > 0 And colCount >= totalCol And totalCol > 2 Then
        addrStart = ws.Cells(rowIndex, 2).Address(False, False)
        addrEnd = ws.Cells(rowIndex, totalCol - 1).Address(False, False)
        ws.Cells(rowIndex, totalCol).Formula = "=SUM(" & addrStart & ":" & addrEnd & ")"
        ws.Cells(rowIndex, totalCol).NumberFormat = "$#,##0"
    End If

    ' Bold entire row if first column contains TOTAL (after removing bold markers)
    titleTxt = UCase$(Replace$(CStr(ws.Cells(rowIndex, 1).Value), "*", ""))
    If InStr(1, titleTxt, "TOTAL", vbTextCompare) > 0 Then
        ws.Rows(rowIndex).Font.Bold = True
    End If
End Sub

' Find a header name (exact match after trim) and return its 1-based index, else 0
Private Function FindHeaderIndex(ByRef headers() As String, ByVal headerName As String) As Long
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If Trim$(headers(i)) = headerName Then
            FindHeaderIndex = i + 1
            Exit Function
        End If
    Next i
    FindHeaderIndex = 0
End Function

' Detect a markdown separator row (---, :---:, etc)
Private Function IsSeparatorRow(ByVal lineText As String) As Boolean
    Dim i As Long, ch As String
    For i = 1 To Len(lineText)
        ch = Mid$(lineText, i, 1)
        If InStr(1, "-|: ", ch, vbBinaryCompare) = 0 Then
            IsSeparatorRow = False
            Exit Function
        End If
    Next i
    IsSeparatorRow = True
End Function

' Get existing sheet or create it
Private Function GetOrCreateSheet(ByVal name As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = name
    End If
    Set GetOrCreateSheet = ws
End Function
