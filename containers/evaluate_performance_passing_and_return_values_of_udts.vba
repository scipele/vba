Option Explicit
'| Item	        | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | evaluate_performance_passing_and_return_values_of_udts.vba  |
'| EntryPoint   | Main                                                        |
'| Purpose      | Evaluate performance of passing and returning UDTs in VBA   |
'| Inputs       | none                                                        |
'| Outputs      | elapsed time in milliseconds                                |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 03/31/2026                                        |

Public Type MatchData
    MatchingId As Long
    NextSearchPosition As Integer
End Type


Sub Main()
    
    Const NUM_METHODS As Long = 3
    Const ITERATIONS As Long = 1000000
    
    Dim methodNames(1 To NUM_METHODS) As String
    methodNames(1) = "A - Func ret w/ temp "
    methodNames(2) = "B - Func ret w/o temp"
    methodNames(3) = "C - ByRef Sub        "
    
    Dim elapsed(1 To NUM_METHODS) As Double
    Dim startTime As Double
    Dim md As MatchData
    Dim i As Long
    Dim m As Long
    
    ' Run each method and record elapsed time
    For m = 1 To NUM_METHODS
        startTime = Timer
        For i = 0 To ITERATIONS
            Select Case m
                Case 1: md = GetDataFromMiddleFuncA()
                Case 2: md = GetDataFromMiddleFuncB()
                Case 3: Call GetDataFromMiddleSubC(md)
            End Select
        Next i
        elapsed(m) = (Timer - startTime) * 1000
        Debug.Print "Method " & methodNames(m) & " - Elapsed: " & elapsed(m) & " ms", "md.MatchingId " & md.MatchingId, "md.NextSearchPosition " & md.NextSearchPosition
    Next m
    
    ' Print comparison to Method A
    Debug.Print vbCrLf
    For m = 2 To NUM_METHODS
        Dim pctFaster As Double
        pctFaster = Round((elapsed(1) / elapsed(m) - 1) * 100, 2)
        Debug.Print "Method " & methodNames(m) & " is " & pctFaster & "% faster than Method " & methodNames(1)
    Next m
    Debug.Print vbCrLf

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
    md.MatchingId = 23
    md.NextSearchPosition = 33
End Sub