Option Explicit

'  filename:    find_latest_rev.xlsm
'  Purpose:     program to find the latest revision of a drawing given job specific circumstanes such as
'               revision information containing "FR" were field revisions that were considered to be the most recent
'
'           List1: base drwing list with revisions
'                            col A          col B
'                           -----------------------
'           header row 1    | Orig Doc No |  Rev  |
'                           -----------------------
'
'           Results: Unique, Latest Revisions w/o "-IR" revisions
'                             col D      col E
'                            ----------------------
'                            | Doc No     |  Rev  |
'                            ----------------------
'
'  Dependencies: None
'
'  update T.Sciple, 8/2/2024, scipele@yahoo.com

Public Sub FindLatestRev()
    'speedup settings
    Call Speedup
    
    'Confirm the Last Row to Clear out from any previous Run
    Dim last_row As Integer
    last_row = LastRowInColm("sht1", "D")
    If last_row > 2 Then
        ThisWorkbook.Sheets("sht1").Range("D2:E" & last_row).ClearContents
    End If
    
    'Read the data from the range
    Dim orig_list As Variant
    Call ReadNamedRangeToAry2D(orig_list, "input_list")
    
    'Transfer Array Name so that the orig List will remain unchanged
    Dim dwg_list As Variant
    dwg_list = orig_list
    
    'Sort the Two Dimensional Array 
    Call SortArrayAtoZ(dwg_list)
    
    'Remove Duplicate Items in the Array
    Call MakeAryUnique(dwg_list)
    
    'Remove any designations in the Revision that Represent "-IR"
    dwg_list = RemoveIRs(dwg_list)
    
    Call RemoveOldRevs(dwg_list)
    
    'Remove Blank Items from the Array that were deleted above
    Call RemTwoDimAryElem(dwg_list, "")
    Call OutputAryToSheet("sht1", "D", "2", dwg_list)
    
    'highlight any old revisions that were deleted
    Call highlightOldRevs(orig_list, dwg_list)
    
    'cleanup
    Call Restore
    Erase orig_list
    Erase dwg_list
End Sub


Private Sub ReadNamedRangeToAry2D(ByRef myAry As Variant, _
                              ByVal namedRangeStr As String)
    'This Sub receives an array from caller passed by reference and re-dimensions the array
    ' to match the number of elements in the named range given

    ' Set the named range
    Dim namedRange As Range
    Set namedRange = ThisWorkbook.Names(namedRangeStr).RefersToRange
    
    myAry = namedRange
End Sub


Private Sub highlightOldRevs(orList, dwg_list)
    Dim highlightRow As Variant
    Dim i, j As Integer
    Dim dontHighlightFlag As Boolean
    
    'setup Array to indicate which elements should be highlighted
    ReDim highlightRow(LBound(orList) To UBound(orList))
    
    'Remove -IR designations since the -IR is basically removed in the final revision listing
    orList = RemoveIRs(orList)
    
    'Do all the Calculation Work before interacting with sheet
    For i = LBound(orList) To UBound(orList)
        dontHighlightFlag = False
        For j = LBound(dwg_list) To UBound(dwg_list)
            If orList(i, 1) & orList(i, 2) = dwg_list(j, 1) & dwg_list(j, 2) Then dontHighlightFlag = True
        Next j
        If dontHighlightFlag = True Then
            highlightRow(i) = 0
        Else:
            highlightRow(i) = 1
        End If
    Next i
    
    'Apply highlight Color on Old Revs
    For i = LBound(highlightRow) To UBound(highlightRow)
        If highlightRow(i) = 1 Then
            Range("A" & i + 1 & ":b" & i + 1).Interior.Color = RGB(146, 208, 80)
        End If
    Next i
End Sub


Private Function RemoveIRs(tmpList)
    Dim i As Integer
    
    'Create a Loop to go thru each element of the array
    For i = LBound(tmpList, 1) To UBound(tmpList, 1) - 1
        If InStr(1, tmpList(i, 2), "-IR", vbTextCompare) > 0 Then
            tmpList(i, 2) = Replace(tmpList(i, 2), "-IR", "", 1, , vbTextCompare)
        End If
    Next i
    
    RemoveIRs = tmpList

End Function


Private Function LastRowInColm(shtName, ColmLtr)
    'Finds the last non-blank cell in a single row or column uses cells function to count all rows in the column,
    'then .end(xlUp) moves up to find the last cell that contains data
    'Usage -> last_row = LastRowInColm("sht1", "F")
    
    Dim colmNo As Long
    
    ThisWorkbook.Sheets(shtName).Activate
    'convert the Column Letter to its numeric value
    colmNo = Range(ColmLtr & 1).Column
    
    LastRowInColm = Cells(Rows.count, colmNo).End(xlUp).Row
End Function


Private Sub RemoveOldRevs(tmpList As Variant)
    Dim i As Integer
    'Create a Loop to go thru each element of the array
    For i = LBound(tmpList, 1) + 1 To UBound(tmpList, 1)
        If tmpList(i, 1) = tmpList(i + 1, 1) Then
             If IsFirstItemLatestRev(tmpList(i, 2), tmpList(i + 1, 2)) = True Then
                tmpList(i + 1, 1) = ""
                tmpList(i + 1, 2) = ""
                Call RemTwoDimAryElem(tmpList, "")
                'move back to recheck item when element is deleted
                i = i - 1
            Else
                tmpList(i, 1) = ""
                tmpList(i, 2) = ""
                Call RemTwoDimAryElem(tmpList, "")
                'move back to recheck item when element is deleted
                i = i - 1
            End If
        End If
        If i = UBound(tmpList, 1) - 1 Then Exit For
    Next i
End Sub


Private Function IsFirstItemLatestRev(first, second)
Dim revTypeA, revTypeB As Integer
    revTypeA = revisionPrecendence(first)
    revTypeB = revisionPrecendence(second)

    If revTypeA < revTypeB Then
        IsFirstItemLatestRev = True
    Else
        IsFirstItemLatestRev = False
    End If

    If revTypeA = revTypeB Then
        Select Case revTypeA
        
        Case 1
            first = UCase(first)
            second = UCase(second)
            first = Val(Replace(first, "FR", "", 1, , vbTextCompare))
            second = Val(Replace(second, "FR", "", 1, , vbTextCompare))
            IsFirstItemLatestRev = chkGreater(first, second)
        Case 2
            If chkGreater(getNumericPart(first), getNumericPart(second)) = True Then
                IsFirstItemLatestRev = True
            Else
                If getNumericPart(first) = getNumericPart(second) Then
                    If chkGreater(UCase(getAlphaPart(first)), UCase(getAlphaPart(second))) = True Then
                        IsFirstItemLatestRev = True
                    End If
                Else
                    IsFirstItemLatestRev = False
                End If
            End If
        
        Case 3
            IsFirstItemLatestRev = chkGreater(first, second)
        Case 4
            first = Replace(first, ".", "/", 1, , vbTextCompare)
            second = Replace(second, ".", "/", 1, , vbTextCompare)
            
            first = Replace(first, "-", "/", 1, , vbTextCompare)
            second = Replace(second, "-", "/", 1, , vbTextCompare)
            
            first = CDate(first)
            second = CDate(second)
            
            IsFirstItemLatestRev = chkGreater(first, second)
    
        End Select
    End If
End Function


Private Function getNumericPart(NumAlphaStr As Variant)
    Dim i As Integer
    Dim tmp As String
    
    For i = 1 To Len(NumAlphaStr)
        If IsNumeric(Mid(NumAlphaStr, i, 1)) = True Then
            tmp = tmp & Mid(NumAlphaStr, i, 1)
        Else:
            Exit For
        End If
    Next i
    getNumericPart = Val(tmp)
End Function


Private Function getAlphaPart(NumAlphaStr As Variant)
    'Purpose - Gets Alpha Character Portion
    Dim i As Integer
    Dim tmp As String
    For i = 1 To Len(NumAlphaStr)
        If Not (IsNumeric(Mid(NumAlphaStr, i, 1))) = True Then
            tmp = tmp & Mid(NumAlphaStr, i, 1)
        End If
    Next i
    
    getAlphaPart = tmp
End Function


Private Function chkGreater(a As Variant, b As Variant)
    If IsNumeric(a) And IsNumeric(b) Then
        If Val(a) > Val(b) Then
            chkGreater = True
        Else
            chkGreater = False
        End If
    Else
        If a > b Then
            chkGreater = True
        Else
            chkGreater = False
        End If
    End If
End Function


Private Function revisionPrecendence(tmpRev As Variant)
    'Purpose - Determine the Hierarchy of a Revions defined from Highest to Lowest as follows
    '1      FR1 - Assume FR is the highest precedence since it indicates a "Field Revision"
    '2      5C - Starts with Number followed by Alpha
    '3      B - Letter Only
    '4      1.10.21 or 1/10/21 Dates are assumed to be relatively low precedence since this is not in the form of a proper revision number
    '5      "-" dash only
    '6      ""  Empty

    'section 1
    If InStr(1, UCase(tmpRev), "FR", vbTextCompare) > 0 Then
        revisionPrecendence = 1
    End If
    'section 2
    If IsNumeric(tmpRev) = True Then
        revisionPrecendence = 2
    End If

    If IsNumeric(tmpRev) = False Then
        If IsNumeric(Left(tmpRev, 1)) = True Then
            revisionPrecendence = 2
        End If
    End If

    'section 3
    If revisionPrecendence = "" Then
        If IsLetters(tmpRev) = True Then
            revisionPrecendence = 3
        End If
    End If

    'section 4
    If InStr(1, tmpRev, ".", vbTextCompare) > 0 Then
        If InStr(InStr(1, tmpRev, ".", vbTextCompare) + 1, tmpRev, ".", vbTextCompare) > 0 Then
            revisionPrecendence = 4
        End If
    End If

    If InStr(1, tmpRev, "-", vbTextCompare) > 0 Then
        If InStr(InStr(1, tmpRev, ".", vbTextCompare) + 1, tmpRev, "-", vbTextCompare) > 0 Then
            revisionPrecendence = 4
        End If
    End If
    
    If InStr(1, tmpRev, "/", vbTextCompare) > 0 Then
        If InStr(InStr(1, tmpRev, ".", vbTextCompare) + 1, tmpRev, "/", vbTextCompare) > 0 Then
            revisionPrecendence = 4
        End If
    End If

    'section 5
    If InStr(1, tmpRev, "-", vbTextCompare) > 0 Then revisionPrecendence = 5

    'section 6
    If tmpRev = "" Then revisionPrecendence = 6
    If tmpRev = " " Then revisionPrecendence = 6
End Function


Private Function IsLetters(Str As Variant) As Boolean
    Dim i As Integer
    For i = 1 To Len(Str)
        Select Case Asc(Mid(Str, i, 1))
            Case 65 To 90, 97 To 122
                IsLetters = True
            Case Else
                IsLetters = False
                Exit For
        End Select
    Next i
End Function


Private Sub SortArrayAtoZ(ByRef Ary As Variant)
    Dim i As Long
    Dim j As Long
    Dim temp1, temp2 As Variant
    
    'Sort the Array A-Z by first dimension of the array
    For i = LBound(Ary) To UBound(Ary) - 1
        For j = i + 1 To UBound(Ary)
            'case in which ist parameter is greater than
            If UCase(Ary(i, 1)) > UCase(Ary(j, 1)) Then
                temp1 = Ary(j, 1)
                temp2 = Ary(j, 2)
                'swap places with 1st column
                Ary(j, 1) = Ary(i, 1)
                Ary(i, 1) = temp1
                'swap places with 2nd column
                Ary(j, 2) = Ary(i, 2)
                Ary(i, 2) = temp2
            End If
        
            'case in which ist parameter is equal to, but second dimension is greater
            If UCase(Ary(i, 1)) = UCase(Ary(j, 1)) Then
                If UCase(Ary(i, 2)) > UCase(Ary(j, 2)) Then
                    temp2 = Ary(j, 2)
                    'swap places with 2nd column
                    Ary(j, 2) = Ary(i, 2)
                    Ary(i, 2) = temp2
                End If
            End If
        Next j
    Next i
End Sub


Private Sub OutputAryToSheet(ByVal shtName As String, _
                             ByVal ColmLoc As String, _
                             ByVal RowTopLoc As String, _
                             ByRef my_ary As Variant)
    'Prints a two dimensional array to a worksheet
    'Usage-> Call OutputAryToSheet("sht1", "D", "2", dwg_list)
    
    'Dimension Variables
    Dim ary_btm_row As Long
    Dim ary_colms As Long
    Dim start_colm_num As Long
    Dim end_colm_num As Long
    Dim EndColmLtr As String
    
    'Initialize Variable Values
    ary_btm_row = RowTopLoc + UBound(my_ary, 1) - 1
    ary_colms = UBound(my_ary, 2)
    start_colm_num = Range(ColmLoc & 1).Column
    
    end_colm_num = start_colm_num + ary_colms - 1
    EndColmLtr = Split(Cells(1, end_colm_num).Address, "$")(1)
    
    'Set Range based on values determined above from array dimensions
    Dim rngTarget As Range
    Set rngTarget = ActiveWorkbook.Worksheets(shtName).Range(ColmLoc & RowTopLoc & ":" & EndColmLtr & ary_btm_row)
    rngTarget = my_ary
End Sub


Private Sub MakeAryUnique(ByRef myAry As Variant)
    ' This function will remove duplicates by using a dictionary object to be able to key if the key exist
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
 
    'populate the dictionary object with the unique values from 'myAry'
    Dim i As Long
    For i = LBound(myAry, 1) To UBound(myAry, 1)
        Dim dwg_rev As String
        dwg_rev = myAry(i, 1) & myAry(i, 2)
        If Not dict.exists(dwg_rev) Then
            dict.Add dwg_rev, i  'i is used as the item so that we can use this index to create the new unique array later
        End If
    Next i
    
    'Create a temporary array to store the unique values of the array
    Dim tmpAry As Variant
    ReDim tmpAry(0 To (dict.count - 1), 1 To 2)
    Dim keyVal As Variant
    
    i = 0
    Dim index As Long
    For Each keyVal In dict
        index = dict(keyVal)
        tmpAry(i, 1) = myAry(index, 1)
        tmpAry(i, 2) = myAry(index, 2)
        i = i + 1
    Next keyVal
        
    'Now reset the original array passed by reference to the temporary array
    myAry = tmpAry
End Sub


Private Sub RemTwoDimAryElem(ByRef Ary As Variant, elemToRemove As String)
    'Next Run a Function to Remove Specified Items from Array
    'Remove "-" and "" from Array
    'categ = RemAryElem(categ, "-")
    'categ = RemAryElem(categ, "")
    
    Dim CleanedAry As Variant
    Dim tempAry As Variant
    Dim i, count As Long
    
    ReDim tempAry(1 To UBound(Ary), 1 To 2)
    
    count = 0
    For i = LBound(Ary, 1) To UBound(Ary, 1)
        If Ary(i, 1) <> elemToRemove Then
            count = count + 1
            tempAry(count, 1) = Ary(i, 1)
            tempAry(count, 2) = Ary(i, 2)
        End If
    Next i
    
    'Now that you know the count then read values into the clean array
    ReDim CleanedAry(1 To count, 1 To 2)
    For i = 1 To count
        CleanedAry(i, 1) = tempAry(i, 1)
        CleanedAry(i, 2) = tempAry(i, 2)
    Next
    
    Ary = CleanedAry
    Erase CleanedAry
    Erase tempAry

End Sub


Private Sub Speedup()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
End Sub


Private Sub Restore()
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
End Sub