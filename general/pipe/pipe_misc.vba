Option Explicit
' filename:         filename.vba
'
' purpose:          xxxx
'
' usage:            xxxx
'
' dependencies:     xxxx
'
' By:               T.Sciple, MM/DD/YYYY

'Define functions that retrieve/calculate data within the pipeData Class Module



Sub SetFunctionDescription()
    Application.MacroOptions Macro:="pipedata", Description:="Returns pipe data given size and schedule" & vbCrLf & _
    "Parameters:" & vbCrLf _
    & "   nps (Nominal Pipe Size)" & vbCrLf _
    & "   sch (Schedule format xs, 40, 80)  " & vbCrLf _
    & "   returnType (Type of data to return:" & vbCrLf _
    & "      'thk' for thickness" & vbCrLf _
    & "      'id' for inside diameter" & vbCrLf _
    & "      'od' for outside diameter)."
End Sub


Public Function pipedata2(ByVal nps As Double, _
                ByVal sch As String, _
                ByVal returnType As String _
                ) As Double

    'returns pipe data given size and schedule
    'returnType - 'thk' returns thickness
    'returnType - 'id' returns id in inches
    'returnType - 'od' returns od in inches
    'errors:    -1 schedule not found
    '           -2 nps not found
    '           -3 invalid return type specified
    
    Dim i As Integer
    Dim j As Integer
    Dim pip_ary() As Double
    Dim sch_ary() As String
    Dim sch_dic As Object   'dictionary object to hold the pipe schedules
    Dim nps_dic As Object   'dictionary object to hold the nominal pipe sizes
    Dim err_code As Integer
    Dim thk_override As Double
    
    sch = LCase(sch)    'convert schedule to lcase
    returnType = LCase(returnType)    'convert schedule to lcase
    thk_override = 0    'initialize to zero
    
    'read the data table for pipe sizes and schedules
    Call get_pipe_ary(pip_ary, sch_ary)
    
    'setup a dictionary for nps as keys and counting index
    Set nps_dic = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(pip_ary, 1) + 1
        nps_dic(pip_ary(i - 1, 0)) = i
    Next i
    
    'print dictionary for ref to immediate window
    'Call print_dict(nps_dic)  (only used for troubleshooting)
    
    'setup a dictionary for schedules with schedule as keys and index
    Set sch_dic = CreateObject("Scripting.Dictionary")
    For j = 1 To UBound(sch_ary) + 1
        sch_dic(sch_ary(j - 1)) = j
    Next j
    
    'print dictionary for ref to immediate window
    'Call print_dict(sch_dic)  (only used for troubleshooting)
    
    'get the index numbers of both dictionary's if the keys exist
    If nps_dic.Exists(nps) Then
        i = nps_dic(nps) - 1 'subtract one because disctionary index is 1 to ... wherase array is base 0
    Else
        err_code = -2
        GoTo err:
    End If
    
    'get the index numbers of both dictionary's if the keys exist
    If sch_dic.Exists(sch) Then
        j = sch_dic(sch) - 1    'subtract one because disctionary index is 1 to ... wherase array is base 0
    Else
        'now check to see if a specified thickness was passed where we will check to see if the number is between 0 and 4 and then assume
        'that this is a custom thickness passed
        If IsNumeric(sch) Then
            If CDbl(sch) > 0 And CDbl(sch) < 4 Then
                thk_override = CDbl(sch)
            End If
        Else
            If returnType <> "od" Then
                err_code = -1
                GoTo err:
            End If
        End If
    End If
    
    'if both keys are retieved now get the pipe data requested
    Select Case returnType
        Case "od"
            pipedata2 = pip_ary(i, 1)
        Case "thk"
            If thk_override > 0 Then
                pipedata2 = thk_override
            Else
                pipedata2 = pip_ary(i, j)
            End If
        Case "id"
            If thk_override > 0 Then
                pipedata2 = pip_ary(i, 1) - 2 * thk_override
            Else
                pipedata2 = pip_ary(i, 1) - 2 * pip_ary(i, j)
            End If
        Case Else
            err_code = -3
            GoTo err:
    End Select

    'cleanup
    Erase pip_ary
    Erase sch_ary
    
    'if errors are not encountered then exit the function bypass error handler below
    Exit Function

err:
    pipedata2 = err_code
    'see error codes at top
End Function


public Function getSize1(strg)
    Dim inchLoc1 As Integer
    inchLoc1 = InStr(1, strg, """", vbTextCompare)
    getSize1 = convFtInToDecIn(Left(strg, inchLoc1))
End Function

Public Function getSize2(strg)
    Dim inchLoc1, inchLoc2, locX, LenLoc As Integer
    Dim tmpSize2 As String
    
    inchLoc1 = InStr(1, strg, """", vbTextCompare)
    inchLoc2 = InStr(inchLoc1 + 1, strg, """", vbTextCompare)
    
    'Make Sure that Size 2 is not actually a length
        LenLoc = InStr(inchLoc1 + 1, LCase(strg), """ long", vbTextCompare)
    
        If LenLoc = 0 Then
            LenLoc = InStr(inchLoc1 + 1, LCase(strg), """ lg", vbTextCompare)
        End If
    
    If inchLoc2 = LenLoc Then
        inchLoc2 = 0
    End If
    
    
    If inchLoc2 = 0 Then
        getSize2 = ""
    Else
        tmpSize2 = Mid(strg, inchLoc1, inchLoc2 - inchLoc1 + 1)
        locX = InStr(1, LCase(tmpSize2), "x", vbTextCompare)
        tmpSize2 = Right(tmpSize2, Len(tmpSize2) - locX)
        getSize2 = convFtInToDecIn(tmpSize2)
    End If
End Function


Public Function get_sch_1(ByVal strg As String) As String
    Dim locX As Integer
    locX = InStr(1, strg, "x", vbBinaryCompare)
    If locX > 0 Then
        get_sch_1 = Left(strg, locX - 2)
    Else
        get_sch_1 = strg
    End If
End Function


Public Function get_sch_2(ByVal strg As String) As String
    Dim locX As Integer
    locX = InStr(1, strg, "x", vbBinaryCompare)
    
    If locX > 0 Then
        get_sch_2 = Right(strg, Len(strg) - locX - 1)
    Else
        get_sch_2 = ""
    End If
End Function