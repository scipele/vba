'==============================================================================
' Filename:    modSetup.vba
' EntryPoint:  SetupInputTables
' Purpose:     Creates all worksheets and ListObject tables for user input
' Dependencies: None
' By:          T. Sciple, 03/17/2026
'==============================================================================

'##############################################################################
'  MODULE: modSetup  (Standard Module)
'  Creates all worksheets and ListObject tables for user input.
'##############################################################################
Option Explicit

Private Const HDR_ROW As Long = 1
Private Const DATA_START_ROW As Long = 2


Public Sub SetupInputTables()
'------------------------------------------------------------------------------
' Entry point: creates all input sheets and tables.
' Safe to re-run — skips sheets/tables that already exist.
'------------------------------------------------------------------------------
    Application.ScreenUpdating = False

    Call CreateConfigSheet
    Call CreateRawMaterialsSheet
    Call CreateBlendingSheet
    Call CreateProcessingSheet
    Call CreateProductsSheet
    Call CreateResultsSheet

    Application.ScreenUpdating = True
    MsgBox "Input tables created. Fill in the tables then run RunSimulation.", _
           vbInformation, "Setup Complete"
End Sub


Private Sub CreateConfigSheet()
    Dim ws As Worksheet
    Set ws = EnsureSheet("Config")

    Dim headers As Variant
    headers = Array("ParamName", "ParamValue")

    Dim tbl As ListObject
    Set tbl = EnsureTable(ws, "tblRunConfig", headers, "A1")

    ' Pre-populate default rows if table is empty
    If tbl.ListRows.Count = 0 Then
        Dim defaults As Variant
        defaults = Array( _
            Array("RunDuration_Days", 30), _
            Array("TimeStep_Hours", 1), _
            Array("UnloadOnWeekends", False), _
            Array("LoadOnWeekends", False), _
            Array("StartDate", #4/1/2026#) _
        )
        Dim i As Long
        For i = LBound(defaults) To UBound(defaults)
            Dim new_row As ListRow
            Set new_row = tbl.ListRows.Add
            new_row.Range(1, 1).Value = defaults(i)(0)
            new_row.Range(1, 2).Value = defaults(i)(1)
        Next i
    End If

    ws.Columns.AutoFit
End Sub


Private Sub CreateRawMaterialsSheet()
    Dim ws As Worksheet
    Set ws = EnsureSheet("RawMaterials")

    ' Unload Schedule
    Dim h1 As Variant
    h1 = Array("ArrivalDay", "Mode", "Quantity_BBL", "MaterialName")
    Call EnsureTable(ws, "tblUnloadSchedule", h1, "A1")

    ' Unload Spots
    Dim h2 As Variant
    h2 = Array("Mode", "NumSpots", "AvgUnloadTime_Hrs", "BBLperLoad")
    Call EnsureTable(ws, "tblUnloadSpots", h2, "F1")

    ' Raw Tanks
    Dim h3 As Variant
    h3 = Array("TankName", "MaterialName", "Capacity_BBL", _
               "StartInventory_BBL", "MinInventory_BBL")
    Call EnsureTable(ws, "tblRawTanks", h3, "K1")

    ws.Columns.AutoFit
End Sub


Private Sub CreateBlendingSheet()
    Dim ws As Worksheet
    Set ws = EnsureSheet("Blending")

    Dim h1 As Variant
    h1 = Array("BlendTankName", "Capacity_BBL", "StartInventory_BBL")
    Call EnsureTable(ws, "tblBlendTanks", h1, "A1")

    Dim h2 As Variant
    h2 = Array("BlendTankName", "MaterialName", "FractionOfBlend")
    Call EnsureTable(ws, "tblBlendRecipes", h2, "E1")

    ws.Columns.AutoFit
End Sub


Private Sub CreateProcessingSheet()
    Dim ws As Worksheet
    Set ws = EnsureSheet("Processing")

    Dim h1 As Variant
    h1 = Array("UnitName", "DesignCapacity_BBL_Day", "FeedSource", "ProductName")
    Call EnsureTable(ws, "tblUnits", h1, "A1")

    ws.Columns.AutoFit
End Sub


Private Sub CreateProductsSheet()
    Dim ws As Worksheet
    Set ws = EnsureSheet("Products")

    Dim h1 As Variant
    h1 = Array("TankName", "ProductName", "Capacity_BBL", _
               "StartInventory_BBL", "MinInventory_BBL")
    Call EnsureTable(ws, "tblProductTanks", h1, "A1")

    Dim h2 As Variant
    h2 = Array("ShipDay", "ProductName", "Quantity_BBL", "Mode")
    Call EnsureTable(ws, "tblLoadSchedule", h2, "G1")

    Dim h3 As Variant
    h3 = Array("Mode", "NumSpots", "AvgLoadTime_Hrs", "BBLperLoad")
    Call EnsureTable(ws, "tblLoadSpots", h3, "L1")

    ws.Columns.AutoFit
End Sub


Private Sub CreateResultsSheet()
    Dim ws As Worksheet
    Set ws = EnsureSheet("Results")
    ' tblSimResults is created dynamically at runtime since columns
    ' depend on how many tanks/units the user configures.
    ' Just ensure the sheet exists for now.
    If ws.ListObjects.Count = 0 Then
        ws.Range("A1").Value = "(Results will be written here after simulation runs)"
    End If
End Sub


'──────────────────────────────────────────────────────────────────────────────
' Helper: ensure a worksheet exists, create if not
'──────────────────────────────────────────────────────────────────────────────
Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets( _
                    ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    Set EnsureSheet = ws
End Function


'──────────────────────────────────────────────────────────────────────────────
' Helper: ensure a ListObject table exists at a given anchor cell
'──────────────────────────────────────────────────────────────────────────────
Private Function EnsureTable(ByVal ws As Worksheet, _
                             ByVal tableName As String, _
                             ByVal headers As Variant, _
                             ByVal anchorAddress As String) As ListObject
    Dim tbl As ListObject
    Dim lo As ListObject

    ' Check if table already exists on sheet
    For Each lo In ws.ListObjects
        If lo.Name = tableName Then
            Set EnsureTable = lo
            Exit Function
        End If
    Next lo

    ' Write headers
    Dim anchor As Range
    Set anchor = ws.Range(anchorAddress)
    Dim col_count As Long
    col_count = UBound(headers) - LBound(headers) + 1

    Dim c As Long
    For c = 0 To col_count - 1
        anchor.Offset(0, c).Value = headers(c)
    Next c

    ' Create table over header row + one blank data row
    Dim tbl_range As Range
    Set tbl_range = anchor.Resize(2, col_count)
    Set tbl = ws.ListObjects.Add(xlSrcRange, tbl_range, , xlYes)
    tbl.Name = tableName
    tbl.TableStyle = "TableStyleMedium2"

    ' Delete the auto-created blank row so user starts fresh
    If tbl.ListRows.Count = 1 Then
        If Application.WorksheetFunction.CountA(tbl.ListRows(1).Range) = 0 Then
            tbl.ListRows(1).Delete
        End If
    End If

    Set EnsureTable = tbl
End Function
