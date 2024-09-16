Option Explicit

' This program uses two methods of search and compares how fast the search is completed
' Sheet 'data'
'      unique_sorted_data_to_be_searched
'      A2 thru A586
'
'     items_to_search_for with several duplicates located in range
'     C2 thru C5000
'
'     Binary Search Result          Using Excel Match Formula
'     D2 thru D5000                 D2 thru F5000

Sub Search_using_Excel_Match_Formula()
    Dim start_time, seconds_elapsed As Double
    'This macro clears and then copies and paste the formula throughout the entire range
    'clear range
    Range("F3:F5000").ClearContents
    
    'copy formula located in cell  F2 ->   =MATCH(C2,dataRng,0)
    Range("F2").Select
    Selection.Copy
    
    'paste over entire range
    Range("F3:F5000").Select
    ActiveSheet.Paste
    
    'start the timer
    start_time = Timer
    
    'Recalc
    Calculate
    
    'compute the time taken to recalculate
    seconds_elapsed = Timer - start_time
    
    'place the time in milli-seconds (i.e. x 1000) onto sheet "run" cell d7
    ActiveWorkbook.Worksheets("run").Range("d7") = seconds_elapsed * 1000
End Sub


Sub TestBinarySearchSubstring()
    Dim start_time, seconds_elapsed As Double
    Dim tmpAry() As Variant
    Dim testAry() As String
    Dim targetSubstringAry As Variant
    Dim targetSubstringItem As String
    Dim result, resultAry As Variant
    Dim i As Long
    Dim shtName As String
    
    ' Assuming tmpAry is a sorted array of strings
    tmpAry = RngToOneDimAry("dataRng")
    testAry = convertAryToString(tmpAry)
    Erase tmpAry

    ' Specify the substring you want to search for (with spaces)
    targetSubstringAry = RngToOneDimAry("targetSubstringRng")
    targetSubstringAry = convertAryToString(targetSubstringAry)
    
    ReDim resultAry(LBound(targetSubstringAry) To UBound(targetSubstringAry), 1 To 2)
    
    'Remember time when macro starts
    start_time = Timer
    
    For i = LBound(targetSubstringAry) To UBound(targetSubstringAry)
        targetSubstringItem = targetSubstringAry(i)
        ' Call the BinarySearchSubstring function
        result = BinarySearchSubstring(testAry, targetSubstringItem)
        resultAry(i, 1) = result
    Next

    seconds_elapsed = Timer - start_time
    shtName = "data"
    'Place Results on the Sheet
   
    Call OutputAryToSheet(shtName, "D", "2", resultAry)
    ActiveWorkbook.Worksheets("run").Range("b7") = seconds_elapsed * 1000
End Sub


Private Function convertAryToString(oldAry As Variant)
    Dim tmpAry() As String
    Dim i As Long
    
    ReDim tmpAry(LBound(oldAry) To UBound(oldAry))
    For i = LBound(oldAry) To UBound(oldAry)
        tmpAry(i) = CStr(oldAry(i))
    Next
    convertAryToString = tmpAry
End Function


Private Function RngToOneDimAry(rngStr As String)
    Dim TwoDimArray As Variant
    Dim OneDimAry As Variant
    Dim i
    'Dump the range into a 2D array
    TwoDimArray = Sheets("data").Range(rngStr).Value

    'Resize the 1D array
    ReDim OneDimAry(1 To UBound(TwoDimArray, 1))

    'Convert 2D to 1D
    For i = 1 To UBound(OneDimAry, 1)
        OneDimAry(i) = TwoDimArray(i, 1)
    Next
    Erase TwoDimArray
    RngToOneDimAry = OneDimAry
End Function


Function BinarySearchSubstring(arr() As String, targetSubstringItem As String) As Long
    Dim low As Long
    Dim high As Long
    Dim mid As Long

    ' Convert to lowercase for a case-insensitive comparison
    targetSubstringItem = LCase(targetSubstringItem)

    low = LBound(arr)
    high = UBound(arr)

    Do While low <= high
        mid = (low + high) \ 2

        ' Convert to lowercase for a case-insensitive comparison, and remove special characters
        Dim currentElement As String
        currentElement = LCase(Replace(Replace(arr(mid), "#", ""), """", ""))

'        Debug.Print "Searching for: " & targetSubstringItem
'        Debug.Print "Current element: " & currentElement

        If currentElement = targetSubstringItem Then
            BinarySearchSubstring = mid
            Exit Function
        ElseIf StrComp(targetSubstringItem, currentElement, vbTextCompare) > 0 Then
            low = mid + 1
        Else
            high = mid - 1
        End If
    Loop

    BinarySearchSubstring = -1 ' Substring not found
End Function


Sub OutputAryToSheet(shtName As String, ColmLoc As String, RowTopLoc As String, resultAry As Variant)
    'Prints a two dimensional array to a worksheet
    'Usage-> Call OutputAryToSheet("sht1", "D", "2", dwgList)
    Dim rngTarget As Range
    Dim AryColms, AryBtmRow, StartColmNum, EndColmNum As Long
    Dim EndColmLtr As String
    
    AryBtmRow = RowTopLoc + UBound(resultAry, 1) - 1
    AryColms = UBound(resultAry, 2)
    
    StartColmNum = ColmLetterToNumber(ColmLoc)
    EndColmNum = StartColmNum + AryColms - 1
    
    EndColmLtr = colmNoToLetter(EndColmNum)
    
    Set rngTarget = ActiveWorkbook.Worksheets(shtName).Range(ColmLoc & RowTopLoc & ":" & EndColmLtr & AryBtmRow)
    rngTarget.Value = resultAry
End Sub


Function ColmLetterToNumber(ColLtr As String)
    'Convert a given letter into it's corresponding Numeric Reference
    'Example Usage  colNum = ColmLetterToNumber("H")

    Dim ColumnNumber As Long
    'Convert To Column Number
   ColmLetterToNumber = Range(ColLtr & 1).Column
End Function


Private Function colmNoToLetter(EndColmNum As Long)
    'Convert To Column Letter
    'Example Usage -> tempColmLtr = colmNoLetter(tempColmNo)
    colmNoToLetter = Split(Cells(1, EndColmNum).Address, "$")(1)
End Function