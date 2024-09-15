Sub speedup_restore(ByVal at_end As Boolean)
    'Use the boolean 'at_end' to restore settings if true or make them false at the start
    Application.ScreenUpdating = at_end
    Application.DisplayStatusBar = at_end
    Application.EnableEvents = at_end
    ActiveSheet.DisplayPageBreaks = at_end
    Application.Calculation = IIf(at_end, xlAutomatic, xlManual)
End Sub
