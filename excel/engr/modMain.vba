'==============================================================================
' Filename:    modMain.vba
' EntryPoint:  RunSimulation, RunAndShowDashboard
' Purpose:     Entry point and orchestration
' Dependencies: modTypes, modHelpers, modSimEngine, modResults, frmDashboard
' By:          T. Sciple, 03/17/2026
'==============================================================================

'##############################################################################
'  MODULE: modMain  (Standard Module)
'  Entry point and orchestration
'##############################################################################
Option Explicit


Public Sub RunSimulation()
'------------------------------------------------------------------------------
' Main entry point.  Loads data, runs simulation, writes results.
'------------------------------------------------------------------------------
    Dim sim As SimState
    Dim start_time As Double
    start_time = Timer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Validate that input sheets exist
    If Not SheetExists("Config") Then
        MsgBox "Run SetupInputTables first to create input sheets.", _
               vbExclamation, "Missing Setup"
        Exit Sub
    End If

    ' Load all input data
    Call LoadSimData(sim)

    ' Validate minimum configuration
    If sim.num_raw_tanks = 0 And sim.num_product_tanks = 0 Then
        MsgBox "No tanks configured. Please fill in the input tables.", _
               vbExclamation, "No Data"
        GoTo CleanUp
    End If

    ' Run the simulation loop
    Call RunSimLoop(sim)

    ' Write results
    Call WriteResults(sim)
    Call WriteSummaryStats(sim)

    Dim elapsed As Double
    elapsed = Timer - start_time

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    If elapsed > 0 Then
        MsgBox "Simulation complete." & vbCrLf & _
               "Steps: " & sim.total_steps & vbCrLf & _
               "Elapsed: " & Format$(elapsed, "0.00") & " seconds" & vbCrLf & _
               "See the Results sheet for output.", _
               vbInformation, "Done"
    End If
End Sub


Public Sub RunAndShowDashboard()
'------------------------------------------------------------------------------
' Runs simulation then opens the graphical dashboard UserForm.
'------------------------------------------------------------------------------
    Call RunSimulation
    frmDashboard.Show vbModeless
End Sub


Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
