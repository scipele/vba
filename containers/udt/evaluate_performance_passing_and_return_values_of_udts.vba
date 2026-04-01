Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | evaluate_performance_passing_and_return_values_of_udts.vba  |
'| EntryPoint   | Main                                                        |
'| Purpose      | Evaluate performance of passing and returning UDTs in VBA   |
'| Inputs       | none                                                        |
'| Outputs      | elapsed time in milliseconds, shown in immediate window     |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 04/01/2026                                        |

Public Type MatchData
    isMatch As Boolean
    NextSearchPosition As Integer
    MatchingId As Long
End Type


Public Type ResultsUdt
    methodNames As String
    elapsedTime As Double
    MatchingId As Long
    NextSearchPosition As Integer
    pctFaster As Double
End Type
    
Const NUM_METHODS As Long = 3
    
'******************************************************************************
'*********************************** MAIN *************************************
'******************************************************************************
    
Sub Main()
    Const ITERATIONS As Long = 1000000
    Dim r(1 To NUM_METHODS) As ResultsUdt
    r(1).methodNames = "A - Func ret w/ temp "
    r(2).methodNames = "B - Func ret w/o temp"
    r(3).methodNames = "C - ByRef Sub        "
    
    Dim startTime As Double
    Dim md As MatchData
    Dim i As Long
    Dim m As Long
    
    ' Run each method and record elapsed time
    For m = 1 To NUM_METHODS
        startTime = Timer
        For i = 0 To ITERATIONS
            'Reset Values to Defaults
            md.isMatch = False
            md.MatchingId = 0
            md.NextSearchPosition = 0
            
            Select Case m
                Case 1: md = GetDataFromMiddleFuncA()
                Case 2: md = GetDataFromMiddleFuncB()
                Case 3: Call GetDataFromMiddleSubC(md)
            End Select
        Next i
        r(m).elapsedTime = Round((Timer - startTime) * 1000, 1)
        r(m).MatchingId = md.MatchingId
        r(m).NextSearchPosition = md.NextSearchPosition
        If m > 1 Then r(m).pctFaster = Round((r(1).elapsedTime / r(m).elapsedTime - 1) * 100, 1)
    Next m
    
    Call OutPutResultsToExcel(r())

End Sub


Sub OutPutResultsToExcel(ByRef r() As ResultsUdt)

    Dim i As Integer
    For i = 1 To NUM_METHODS
        ThisWorkbook.Sheets("ShtResults").Range("B" & i + 3) = r(i).methodNames
        ThisWorkbook.Sheets("ShtResults").Range("C" & i + 3) = r(i).elapsedTime
        ThisWorkbook.Sheets("ShtResults").Range("D" & i + 3) = r(i).MatchingId
        ThisWorkbook.Sheets("ShtResults").Range("E" & i + 3) = r(i).NextSearchPosition
        
        If i > 1 Then
            ThisWorkbook.Sheets("ShtResults").Range("F" & i + 3) = r(i).pctFaster & "% faster than Method " & r(i).methodNames
        End If
    Next i
End Sub



'******************************************************************************
'************ METHOD A - Returning the UDT from the function ******************
'******************************************************************************
' The Middle Function (The Caller)
Function GetDataFromMiddleFuncA() As MatchData
    ' Simply assign the result of the inner to the outer
    GetDataFromMiddleFuncA = GetDataFromDeepestFuncA()
End Function


' The Deepest Function
Function GetDataFromDeepestFuncA() As MatchData
    Dim tmp As MatchData
    tmp.isMatch = True
    tmp.MatchingId = 21
    tmp.NextSearchPosition = 31
    
    GetDataFromDeepestFuncA = tmp ' Return the UDT
End Function


'******************************************************************************
'************ METHOD B - Returning the UDT from the function ******************
'******************************************************************************
' The Middle Function (The Caller)
Function GetDataFromMiddleFuncB() As MatchData
    ' Simply assign the result of the inner to the outer
    GetDataFromMiddleFuncB = GetDataFromDeepestFuncB()
End Function


' The Deepest Function
Function GetDataFromDeepestFuncB() As MatchData
    GetDataFromDeepestFuncB.isMatch = True
    GetDataFromDeepestFuncB.MatchingId = 22
    GetDataFromDeepestFuncB.NextSearchPosition = 32
End Function


'******************************************************************************
'************ METHOD C - Returning the UDT from the sub **********************
'******************************************************************************
' The Middle Function (The Caller)
Sub GetDataFromMiddleSubC(ByRef md As MatchData)
    ' Simply assign the result of the inner to the outer
    Call GetDataFromDeepestSubC(md)
End Sub


' The Deepest Sub
Sub GetDataFromDeepestSubC(ByRef md As MatchData)
    md.isMatch = True
    md.MatchingId = 23
    md.NextSearchPosition = 33
End Sub