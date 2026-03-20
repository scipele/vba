'==============================================================================
' Filename:    modResults.vba
' Purpose:     Writes simulation snapshots to the Results sheet as a
'              ListObject table with summary statistics
' Dependencies: modTypes
' By:          T. Sciple, 03/17/2026
'==============================================================================

'##############################################################################
'  MODULE: modResults  (Standard Module)
'  Writes simulation snapshots to the Results sheet as a ListObject table.
'##############################################################################
Option Explicit


Public Sub WriteResults(ByRef state As SimState)
'------------------------------------------------------------------------------
' Builds dynamic column headers based on configured tanks/units, then writes
' all snapshot data to tblSimResults on the Results sheet.
'------------------------------------------------------------------------------
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Results")

    ' Clear existing content
    ws.Cells.Clear

    ' ── Build headers ──
    Dim headers() As String
    Dim col As Long
    col = 0

    ' Fixed columns
    ReDim headers(0 To 0)
    Call AppendHeader(headers, col, "SimStep")
    Call AppendHeader(headers, col, "DateTime")
    Call AppendHeader(headers, col, "UnloadingActive")
    Call AppendHeader(headers, col, "LoadingActive")
    Call AppendHeader(headers, col, "Flags")

    ' Raw tank columns
    Dim i As Long
    For i = 0 To state.num_raw_tanks - 1
        Call AppendHeader(headers, col, "Raw_" & state.raw_tanks(i).tank_name & "_BBL")
    Next i

    ' Blend tank columns
    For i = 0 To state.num_blend_tanks - 1
        Call AppendHeader(headers, col, "Blend_" & state.blend_tanks(i).tank_name & "_BBL")
    Next i

    ' Product tank columns
    For i = 0 To state.num_product_tanks - 1
        Call AppendHeader(headers, col, "Prod_" & state.product_tanks(i).tank_name & "_BBL")
    Next i

    ' Unit throughput columns
    For i = 0 To state.num_units - 1
        Call AppendHeader(headers, col, "Unit_" & state.units(i).unit_name & "_BBL")
    Next i

    Dim total_cols As Long
    total_cols = col

    ' ── Write headers ──
    Dim c As Long
    For c = 0 To total_cols - 1
        ws.Cells(1, c + 1).Value = headers(c)
    Next c

    ' ── Write data rows in bulk (build 2D array) ──
    Dim total_rows As Long
    total_rows = state.total_steps

    Dim data_arr() As Variant
    ReDim data_arr(1 To total_rows, 1 To total_cols)

    Dim s As Long
    For s = 0 To total_rows - 1
        col = 0
        col = col + 1: data_arr(s + 1, col) = state.snapshots(s).sim_step
        col = col + 1: data_arr(s + 1, col) = state.snapshots(s).date_time
        col = col + 1: data_arr(s + 1, col) = state.snapshots(s).unloading_active
        col = col + 1: data_arr(s + 1, col) = state.snapshots(s).loading_active
        col = col + 1: data_arr(s + 1, col) = state.snapshots(s).flag_text

        For i = 0 To state.num_raw_tanks - 1
            col = col + 1
            If Not Not state.snapshots(s).raw_inventories Then
                data_arr(s + 1, col) = state.snapshots(s).raw_inventories(i)
            End If
        Next i

        For i = 0 To state.num_blend_tanks - 1
            col = col + 1
            If Not Not state.snapshots(s).blend_inventories Then
                data_arr(s + 1, col) = state.snapshots(s).blend_inventories(i)
            End If
        Next i

        For i = 0 To state.num_product_tanks - 1
            col = col + 1
            If Not Not state.snapshots(s).product_inventories Then
                data_arr(s + 1, col) = state.snapshots(s).product_inventories(i)
            End If
        Next i

        For i = 0 To state.num_units - 1
            col = col + 1
            If Not Not state.snapshots(s).unit_throughputs Then
                data_arr(s + 1, col) = state.snapshots(s).unit_throughputs(i)
            End If
        Next i
    Next s

    ' Write array to sheet
    ws.Range(ws.Cells(2, 1), ws.Cells(total_rows + 1, total_cols)).Value = data_arr

    ' Create ListObject
    Dim tbl_range As Range
    Set tbl_range = ws.Range(ws.Cells(1, 1), ws.Cells(total_rows + 1, total_cols))

    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, tbl_range, , xlYes)
    tbl.Name = "tblSimResults"
    tbl.TableStyle = "TableStyleMedium9"

    ' Format DateTime column
    tbl.ListColumns("DateTime").DataBodyRange.NumberFormat = "yyyy-mm-dd hh:mm"

    ws.Columns.AutoFit
End Sub


Private Sub AppendHeader(ByRef headers() As String, _
                           ByRef idx As Long, _
                           ByVal headerName As String)
    If idx > UBound(headers) Then
        ReDim Preserve headers(0 To idx)
    End If
    headers(idx) = headerName
    idx = idx + 1
End Sub


Public Sub WriteSummaryStats(ByRef state As SimState)
'------------------------------------------------------------------------------
' Writes min/max/avg inventory stats below or beside the results table.
'------------------------------------------------------------------------------
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Results")

    Dim summary_row As Long
    summary_row = state.total_steps + 4  ' leave a gap

    ws.Cells(summary_row, 1).Value = "=== SUMMARY STATISTICS ==="
    ws.Cells(summary_row, 1).Font.Bold = True
    summary_row = summary_row + 1

    ws.Cells(summary_row, 1).Value = "Tank"
    ws.Cells(summary_row, 2).Value = "Type"
    ws.Cells(summary_row, 3).Value = "Min BBL"
    ws.Cells(summary_row, 4).Value = "Max BBL"
    ws.Cells(summary_row, 5).Value = "Avg BBL"
    ws.Cells(summary_row, 6).Value = "Capacity"
    ws.Cells(summary_row, 7).Value = "Min % Full"
    ws.Cells(summary_row, 8).Value = "Max % Full"
    Dim hdr_row As Long
    hdr_row = summary_row
    summary_row = summary_row + 1

    ' Raw tanks
    Dim i As Long
    Dim s As Long
    For i = 0 To state.num_raw_tanks - 1
        Dim r_min As Double, r_max As Double, r_sum As Double
        r_min = 1E+30: r_max = -1E+30: r_sum = 0
        For s = 0 To state.total_steps - 1
            Dim r_val As Double
            r_val = state.snapshots(s).raw_inventories(i)
            If r_val < r_min Then r_min = r_val
            If r_val > r_max Then r_max = r_val
            r_sum = r_sum + r_val
        Next s
        ws.Cells(summary_row, 1).Value = state.raw_tanks(i).tank_name
        ws.Cells(summary_row, 2).Value = "Raw"
        ws.Cells(summary_row, 3).Value = Round(r_min, 1)
        ws.Cells(summary_row, 4).Value = Round(r_max, 1)
        ws.Cells(summary_row, 5).Value = Round(r_sum / state.total_steps, 1)
        ws.Cells(summary_row, 6).Value = state.raw_tanks(i).capacity_bbl
        ws.Cells(summary_row, 7).Value = Round(r_min / state.raw_tanks(i).capacity_bbl * 100, 1)
        ws.Cells(summary_row, 8).Value = Round(r_max / state.raw_tanks(i).capacity_bbl * 100, 1)
        summary_row = summary_row + 1
    Next i

    ' Product tanks
    For i = 0 To state.num_product_tanks - 1
        Dim p_min As Double, p_max As Double, p_sum As Double
        p_min = 1E+30: p_max = -1E+30: p_sum = 0
        For s = 0 To state.total_steps - 1
            Dim p_val As Double
            p_val = state.snapshots(s).product_inventories(i)
            If p_val < p_min Then p_min = p_val
            If p_val > p_max Then p_max = p_val
            p_sum = p_sum + p_val
        Next s
        ws.Cells(summary_row, 1).Value = state.product_tanks(i).tank_name
        ws.Cells(summary_row, 2).Value = "Product"
        ws.Cells(summary_row, 3).Value = Round(p_min, 1)
        ws.Cells(summary_row, 4).Value = Round(p_max, 1)
        ws.Cells(summary_row, 5).Value = Round(p_sum / state.total_steps, 1)
        ws.Cells(summary_row, 6).Value = state.product_tanks(i).capacity_bbl
        ws.Cells(summary_row, 7).Value = Round(p_min / state.product_tanks(i).capacity_bbl * 100, 1)
        ws.Cells(summary_row, 8).Value = Round(p_max / state.product_tanks(i).capacity_bbl * 100, 1)
        summary_row = summary_row + 1
    Next i

    ' Bold the header row
    ws.Range(ws.Cells(hdr_row, 1), ws.Cells(hdr_row, 8)).Font.Bold = True
    ws.Columns.AutoFit
End Sub
