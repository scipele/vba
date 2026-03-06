Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | startTeamsMtgNow.vba                                        |
'| EntryPoint   | StartTeamsMeetNow                                           |
'| Purpose      | Launch a Teams instant meeting, automatically copy the       |
'|              | join link, and open a new Outlook email with the link so     |
'|              | the user can address and send manually.                      |
'| Inputs       | None (fully automated; fallback prompt if auto-copy fails)   |
'| Outputs      | New Outlook email with meeting link in body                  |
'| Dependencies | Microsoft Teams desktop app (new), Windows Script Host       |
'| By Name,Date | T.Sciple, 03/06/2026                                        |
'
' NOTES:
'   - SendKeys-based UI automation is fragile. If Teams updates change
'     keyboard shortcuts or layout, the navigation steps may need adjustment.
'   - Timing constants (WAIT_*) may need tuning for slower machines.
'   - The macro targets the NEW Teams desktop app (ms-teams.exe).
'     If you still use classic Teams (Teams.exe), update TEAMS_PROCESS_NAME.
'
' AUTOMATED FLOW:
'   1. Click Meet Now            (keyboard shortcut)
'   2. Click Start Meeting       (Enter on pre-join screen)
'   3. Click Join Now            (Enter on join confirmation)
'   4. Copy Meeting Link         (navigate meeting toolbar via SendKeys)
'   5. Create email with link    (Outlook MailItem)

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" ( _
        ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" ( _
        ByVal dwMilliseconds As Long)
#End If

' ── Timing constants (milliseconds) ── adjust for your system speed
Private Const WAIT_SHORT  As Long = 1500
Private Const WAIT_MEDIUM As Long = 3000
Private Const WAIT_LONG   As Long = 5000
Private Const WAIT_LAUNCH As Long = 12000

' ── Process name for the NEW Teams desktop app
Private Const TEAMS_PROCESS_NAME As String = "ms-teams.exe"


'=============================================================================
' EntryPoint – call this from a ribbon button or the Macros dialog
'=============================================================================
Public Sub StartTeamsMeetNow()
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")

    ' ── Step 0: Open Teams if it is not already running ──
    If Not IsTeamsRunning() Then
        LaunchTeams wsh
    End If

    ' ── Bring Teams to the foreground ──
    If Not ActivateTeamsWindow(wsh) Then
        MsgBox "Could not find the Teams window." & vbCrLf & _
               "Please open Teams manually and try again.", _
               vbExclamation, "Teams Not Found"
        Exit Sub
    End If
    Sleep WAIT_SHORT

    ' ── Step 1: Click Meet Now ──
    ClickMeetNow
    Sleep WAIT_LONG

    ' ── Step 2: Click Start Meeting ──
    ClickStartMeeting
    Sleep WAIT_LONG

    ' ── Step 3: Click Join Now ──
    ClickJoinNow
    Sleep WAIT_LONG               ' let the meeting fully load

    ' ── Step 4: Copy Meeting Link ──
    ClearClipboard
    Sleep 500
    CopyMeetingLinkAuto
    Sleep WAIT_MEDIUM

    ' ── Read link from clipboard ──
    Dim meeting_link As String
    meeting_link = Trim(ReadClipboard())

    ' If auto-copy did not land a valid link, ask the user once
    If Not IsValidTeamsLink(meeting_link) Then
        meeting_link = PromptForMeetingLink()
        If meeting_link = "" Then Exit Sub
    End If

    ' ── Step 5: Create a new email with the link ──
    CreateMeetNowEmail meeting_link
End Sub


'=============================================================================
' Check whether Teams is running via WMI process query
'=============================================================================
Private Function IsTeamsRunning() As Boolean
    Dim wmi As Object
    Dim procs As Object

    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set procs = wmi.ExecQuery( _
        "SELECT Name FROM Win32_Process WHERE Name = '" & TEAMS_PROCESS_NAME & "'")

    IsTeamsRunning = (procs.Count > 0)
End Function


'=============================================================================
' Launch the Teams desktop app
'=============================================================================
Private Sub LaunchTeams(ByVal wsh As Object)
    ' New Teams is a packaged (MSIX) app – launch via its App User Model ID.
    ' For classic Teams, replace with the full path to Teams.exe.
    wsh.Run "explorer shell:AppsFolder\MSTeams_8wekyb3d8bbwe!MSTeams"
    Sleep WAIT_LAUNCH
End Sub


'=============================================================================
' Bring the Teams window to the foreground
'
' AppActivate throws an error (rather than returning False) when the window
' is not found, so we must wrap each attempt with error handling.
' We also retry several times because Teams may still be loading.
' The title is matched as a substring, so "Microsoft Teams" will match
' window titles like "Microsoft Teams (work or school)" or
' "John Doe | Microsoft Teams".
'=============================================================================
Private Function ActivateTeamsWindow(ByVal wsh As Object) As Boolean
    Dim titles As Variant
    titles = Array("Microsoft Teams", "Teams")

    Dim MAX_RETRIES As Long
    MAX_RETRIES = 6                ' 6 retries x 2 s = 12 s total wait

    Dim attempt As Long
    Dim i As Long

    For attempt = 1 To MAX_RETRIES
        For i = LBound(titles) To UBound(titles)
            On Error Resume Next
            wsh.AppActivate titles(i)
            If Err.Number = 0 Then
                On Error GoTo 0
                ActivateTeamsWindow = True
                Exit Function
            End If
            Err.Clear
            On Error GoTo 0
        Next i
        Sleep 2000               ' wait 2 s before retrying
    Next attempt

    ActivateTeamsWindow = False
End Function


'=============================================================================
' Step 1 – Click Meet Now
'
' Sends Ctrl+Shift+R which is the "Meet now" shortcut in new Teams.
' This opens the pre-join / meeting-setup screen.
'
' ALTERNATIVE (if shortcut does not work on your build):
'   SendKeys "^4", True       ' Ctrl+4 = open Calendar
'   Sleep WAIT_MEDIUM
'   ' Then Tab to the "Meet now" button and press Enter
'=============================================================================
Private Sub ClickMeetNow()
    SendKeys "^+r", True
End Sub


'=============================================================================
' Step 2 – Click Start Meeting on the pre-join screen
'=============================================================================
Private Sub ClickStartMeeting()
    SendKeys "{ENTER}", True
End Sub


'=============================================================================
' Step 3 – Click Join Now (second confirmation)
'
' Some Teams builds show a separate "Join now" button after "Start meeting".
' If your build goes straight into the meeting after Step 2, this extra
' Enter is harmless.
'=============================================================================
Private Sub ClickJoinNow()
    SendKeys "{ENTER}", True
End Sub


'=============================================================================
' Step 4 – Copy the meeting link from the in-meeting toolbar
'
' After joining, the meeting toolbar is visible at the top.  This routine
' attempts to navigate it via keyboard to reach "Copy meeting link".
'
' APPROACH (new Teams - 2024/2025 builds):
'   1. Dismiss any overlay popup (ESC)
'   2. Ctrl+Shift+I  – opens the Meeting Info / Details side-pane
'      (If this shortcut does not exist in your build, fall back to the
'       Tab-through-toolbar approach below.)
'   3. Tab through the pane to the "Copy meeting link" button
'   4. Press Enter to copy the link to the clipboard
'
' Because the number of Tab presses can vary between builds, the main sub
' falls back to a one-click InputBox if the clipboard does not contain a
' valid Teams link after this routine runs.
'=============================================================================
Private Sub CopyMeetingLinkAuto()
    ' Dismiss any welcome / tips overlay
    SendKeys "{ESC}", True
    Sleep WAIT_SHORT

    ' ── Primary attempt: Ctrl+Shift+I for Meeting Info pane ──
    SendKeys "^+i", True
    Sleep WAIT_MEDIUM

    ' Tab through the info pane looking for "Copy meeting link"
    Dim i As Long
    For i = 1 To 4
        SendKeys "{TAB}", True
        Sleep 400
    Next i
    SendKeys "{ENTER}", True      ' press the Copy button
    Sleep WAIT_SHORT

    ' Check if we got a valid link; if not, try the (...) menu approach
    If IsValidTeamsLink(Trim(ReadClipboard())) Then Exit Sub

    ' ── Fallback attempt: (...) More actions menu ──
    ' Close whatever pane we opened
    SendKeys "{ESC}", True
    Sleep 500

    ' Alt+F6 cycles focus between meeting regions (toolbar, content, panel)
    SendKeys "%{F6}", True
    Sleep WAIT_SHORT

    ' Tab through toolbar buttons to reach (...) More actions
    For i = 1 To 8
        SendKeys "{TAB}", True
        Sleep 300
    Next i
    SendKeys "{ENTER}", True      ' open the (...) menu
    Sleep WAIT_SHORT

    ' Arrow-down through the menu looking for "Meeting info" / "Copy join info"
    For i = 1 To 6
        SendKeys "{DOWN}", True
        Sleep 300
    Next i
    SendKeys "{ENTER}", True      ' select the menu item
    Sleep WAIT_MEDIUM

    ' Tab to "Copy meeting link" in the pane that opens
    For i = 1 To 4
        SendKeys "{TAB}", True
        Sleep 400
    Next i
    SendKeys "{ENTER}", True
    Sleep WAIT_SHORT
End Sub


'=============================================================================
' Fallback: simplified prompt when auto-copy did not capture the link.
' The user just needs to click "Copy meeting link" in Teams, then click OK.
'=============================================================================
Private Function PromptForMeetingLink() As String
    Dim meeting_link As String
    Dim attempt As Long
    Const MAX_ATTEMPTS As Long = 3

    For attempt = 1 To MAX_ATTEMPTS
        MsgBox "The meeting link was not captured automatically." & vbCrLf & vbCrLf & _
               "In the Teams meeting window:" & vbCrLf & _
               "  1. Click (...) More actions  >  Meeting info" & vbCrLf & _
               "  2. Click  'Copy meeting link'" & vbCrLf & vbCrLf & _
               "Then click OK here.", _
               vbInformation, "Copy Meeting Link  (attempt " & attempt & " of " & MAX_ATTEMPTS & ")"

        meeting_link = Trim(ReadClipboard())
        If IsValidTeamsLink(meeting_link) Then
            PromptForMeetingLink = meeting_link
            Exit Function
        End If
    Next attempt

    MsgBox "Could not get a valid Teams meeting link after " & MAX_ATTEMPTS & " attempts." & vbCrLf & _
           "The email will not be created.", _
           vbExclamation, "No Meeting Link"
    PromptForMeetingLink = ""
End Function


'=============================================================================
' Validate that a string looks like a Teams meeting URL
'=============================================================================
Private Function IsValidTeamsLink(ByVal url As String) As Boolean
    If Left$(LCase$(url), 8) <> "https://" Then
        IsValidTeamsLink = False
        Exit Function
    End If

    If InStr(1, url, "teams.microsoft.com", vbTextCompare) > 0 Or _
       InStr(1, url, "teams.live.com", vbTextCompare) > 0 Then
        IsValidTeamsLink = True
    Else
        IsValidTeamsLink = False
    End If
End Function


'=============================================================================
' Clear the Windows clipboard so stale content does not get mistaken for
' a freshly copied meeting link
'=============================================================================
Private Sub ClearClipboard()
    On Error Resume Next
    Dim data_obj As Object
    Set data_obj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    data_obj.SetText ""
    data_obj.PutInClipboard
    On Error GoTo 0
End Sub


'=============================================================================
' Read text from the Windows clipboard via MSForms.DataObject (late binding)
'=============================================================================
Private Function ReadClipboard() As String
    On Error GoTo Err_ReadClipboard

    ' Late-bind the MSForms.DataObject by its CLSID
    Dim data_obj As Object
    Set data_obj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    data_obj.GetFromClipboard
    ReadClipboard = data_obj.GetText
    Exit Function

Err_ReadClipboard:
    ReadClipboard = ""
End Function


'=============================================================================
' Create a new Outlook email with the meeting link
'=============================================================================
Private Sub CreateMeetNowEmail(ByVal meetingLink As String)
    Dim mail_item As Outlook.MailItem
    Set mail_item = Application.CreateItem(olMailItem)

    mail_item.Subject = "Meet Now"
    mail_item.HTMLBody = "<p>Join the Teams meeting now:</p>" & _
                         "<p><a href=""" & meetingLink & """>" & meetingLink & "</a></p>" & _
                         "<br><p>Click the link above or paste it into your browser.</p>"

    ' Display the email so the user can address and send manually
    mail_item.Display
End Sub
