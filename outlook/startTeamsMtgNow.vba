Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | startTeamsMtgNow.vba                                        |
'| EntryPoint   | StartTeamsMeetNow                                           |
'| Purpose      | Launch a Teams "Meet Now", then prompt the user to copy the  |
'|              | join link and create a new Outlook email with it.            |
'| Inputs       | None                                                        |
'| Outputs      | New Outlook email with meeting link in body                  |
'| Dependencies | Microsoft Teams desktop app (new), Windows Script Host       |
'| By Name,Date | T.Sciple, 03/06/2026                                        |
'
' FLOW:
'   1. Launch Teams if needed, bring it to foreground
'   2. Send Ctrl+Shift+R  (Meet Now shortcut in new Teams)
'   3. Send Enter twice   (Start Meeting -> Join Now)
'   4. Prompt user to copy the meeting link in Teams
'   5. Create Outlook email with the link

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const TEAMS_PROCESS As String = "ms-teams.exe"


'=============================================================================
' EntryPoint
'=============================================================================
Public Sub StartTeamsMeetNow()
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Debug.Print "===== StartTeamsMeetNow BEGIN " & Now & " ====="

    ' ── Launch Teams if not running ──
    If Not IsProcessRunning(TEAMS_PROCESS) Then
        Debug.Print "Step 0: Teams not running - launching..."
        wsh.Run "explorer shell:AppsFolder\MSTeams_8wekyb3d8bbwe!MSTeams"
        Sleep 10000
        Debug.Print "Step 0: launch wait complete"
    Else
        Debug.Print "Step 0: Teams already running"
    End If

    ' ── Activate Teams ──
    Debug.Print "Step 0: activating Teams window..."
    If Not TryActivate(wsh, 5) Then
        Debug.Print "Step 0: FAILED - could not activate Teams"
        MsgBox "Could not find the Teams window.", vbExclamation
        Exit Sub
    End If
    Debug.Print "Step 0: Teams activated OK"
    Sleep 1000

    ' ── Meet Now  (Ctrl+Shift+R) ──
    Debug.Print "Step 1: Meet Now - sending Ctrl+Shift+R"
    TryActivate wsh, 1
    SendKeys "^+r", True
    Sleep 4000
    Debug.Print "Step 1: done"

    ' ── Start Meeting (Enter) ──
    Debug.Print "Step 2: Start Meeting - sending ENTER"
    TryActivate wsh, 1
    SendKeys "{ENTER}", True
    Sleep 4000
    Debug.Print "Step 2: done"

    ' ── Join Now (Enter) ──
    Debug.Print "Step 3: Join Now - sending ENTER"
    TryActivate wsh, 1
    SendKeys "{ENTER}", True
    Sleep 3000
    Debug.Print "Step 3: done"

    ' ── Ask user to copy the link ──
    Debug.Print "Step 4: prompting user to copy meeting link"
    Dim link As String
    link = GetLinkFromUser()
    If link = "" Then
        Debug.Print "Step 4: no valid link - EXIT"
        Exit Sub
    End If
    Debug.Print "Step 4: got link = [" & link & "]"

    ' ── Create the email ──
    Debug.Print "Step 5: creating email"
    Dim mi As Outlook.MailItem
    Set mi = Application.CreateItem(olMailItem)
    mi.Subject = "Meet Now"
    mi.HTMLBody = "<p>Join the Teams meeting now:</p>" & _
                  "<p><a href=""" & link & """>" & link & "</a></p>"
    mi.Display
    Debug.Print "Step 5: email displayed"
    Debug.Print "===== StartTeamsMeetNow END " & Now & " ====="
End Sub


'=============================================================================
' Prompt user to copy the meeting link in Teams, read it from clipboard.
'=============================================================================
Private Function GetLinkFromUser() As String
    Dim clip As String
    Dim attempt As Long

    For attempt = 1 To 3
        MsgBox "In the Teams meeting:" & vbCrLf & vbCrLf & _
               "  1. Click the (i) info icon  -or-" & vbCrLf & _
               "     (...) More actions > Meeting info" & vbCrLf & _
               "  2. Click  ""Copy meeting link""" & vbCrLf & vbCrLf & _
               "Then click OK.", _
               vbInformation, "Copy Meeting Link"

        clip = Trim(ReadClipboard())
        Debug.Print "  GetLinkFromUser: attempt " & attempt & " clipboard = [" & Left$(clip, 80) & "]"
        If IsTeamsLink(clip) Then
            Debug.Print "  GetLinkFromUser: valid link found"
            GetLinkFromUser = clip
            Exit Function
        End If
        Debug.Print "  GetLinkFromUser: not a valid Teams link"
    Next attempt

    Debug.Print "  GetLinkFromUser: gave up after 3 attempts"
    MsgBox "No valid Teams link found on the clipboard.", vbExclamation
    GetLinkFromUser = ""
End Function


'=============================================================================
' Try to activate the Teams window. Returns True on success.
'=============================================================================
Private Function TryActivate(ByVal wsh As Object, ByVal retries As Long) As Boolean
    Dim titles As Variant
    titles = Array("Microsoft Teams", "Teams")
    Dim attempt As Long, i As Long

    For attempt = 1 To retries
        For i = LBound(titles) To UBound(titles)
            On Error Resume Next
            wsh.AppActivate titles(i)
            If Err.Number = 0 Then
                On Error GoTo 0
                Sleep 300
                Debug.Print "  TryActivate: matched '" & titles(i) & "' on attempt " & attempt
                TryActivate = True
                Exit Function
            End If
            Err.Clear
            On Error GoTo 0
        Next i
        If attempt < retries Then Sleep 2000
    Next attempt
    Debug.Print "  TryActivate: FAILED after " & retries & " retries"
    TryActivate = False
End Function


'=============================================================================
' Check if a process is running
'=============================================================================
Private Function IsProcessRunning(ByVal procName As String) As Boolean
    Dim wmi As Object, procs As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set procs = wmi.ExecQuery("SELECT Name FROM Win32_Process WHERE Name='" & procName & "'")
    IsProcessRunning = (procs.Count > 0)
End Function


'=============================================================================
' Validate a Teams meeting URL
'=============================================================================
Private Function IsTeamsLink(ByVal url As String) As Boolean
    If Left$(LCase$(url), 8) <> "https://" Then
        IsTeamsLink = False
    ElseIf InStr(1, url, "teams.microsoft.com", vbTextCompare) > 0 Or _
           InStr(1, url, "teams.live.com", vbTextCompare) > 0 Then
        IsTeamsLink = True
    Else
        IsTeamsLink = False
    End If
End Function


'=============================================================================
' Read text from the Windows clipboard
'=============================================================================
Private Function ReadClipboard() As String
    On Error GoTo ErrOut
    Dim obj As Object
    Set obj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    obj.GetFromClipboard
    ReadClipboard = obj.GetText
    Exit Function
ErrOut:
    ReadClipboard = ""
End Function
