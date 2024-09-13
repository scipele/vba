Option Explicit

'The example Sub Below is just to illustrate a few ways in which arguments (or values) can be passed to subs
' and how the parameters (place holder for the values) can be defined in subs/functions

'From Microsoft:
'Choice of Passing Mechanism
'You should choose the passing mechanism carefully for each argument.

'Protection. In choosing between the two passing mechanisms, the most important criterion is the exposure of calling variables to change.
'The advantage of passing an argument ByRef is that the procedure can return a value to the calling code through that argument.
'The advantage of passing an argument ByVal is that it protects a variable from being changed by the procedure.

'Performance. Although the passing mechanism can affect the performance of your code,
'the difference is usually insignificant. One exception to this is a value type passed ByVal.
'In this case, Visual Basic copies the entire data contents of the argument.
'Therefore, for a large value type such as a structure, it can be more efficient to pass it ByRef.

'When to Pass an Argument by Value
'If the calling code element underlying the argument is a nonmodifiable element, declare the corresponding parameter ByVal. No code can change the value of a nonmodifiable element.

'If the underlying element is modifiable, but you do not want the procedure to be able to change its value, declare the parameter ByVal.
'Only the calling code can change the value of a modifiable element passed by value.

'When to Pass an Argument by Reference
'If the procedure has a genuine need to change the underlying element in the calling code, declare the corresponding parameter ByRef.

'If the correct execution of the code depends on the procedure changing the underlying element in the calling code, declare the parameter ByRef.
'If you pass it by value, or if the calling code overrides the ByRef passing mechanism by enclosing the argument in parentheses, the procedure call might produce unexpected results.

Sub main()
    Dim FruitType As String
    Dim color As String
    Dim size As String
    
    FruitType = "Banana"
    color = "Yellow"
    size = "Med-Large"
    Debug.Print _
        "Before Sub Call" & Chr(13) _
        ; ("FruitType = " & FruitType & Chr(13) _
        & "color = " & color & Chr(13) _
        & "size = " & size & Chr(13))
    
    ' the values of FruitType and color below are called arguments in the calling code below in this sub
    ' these values are passed to the 'parameters' in the sub definition below
    Call checkFruitDetails( _
        FruitType, _
        color, _
        size)
    
    'print after sub call to illustrate how the original 'color' variab byRef variable can be changed
    Debug.Print _
        "After Sub Call" & Chr(13) _
        ; ("FruitType = " & FruitType & Chr(13) _
        & "color = " & color & "     *** Note how the change is reflected back in main sub b/c param is defined byRef *** " & Chr(13) _
        & "size = " & size & "         *** Note how the change is not reflected back in main sub b/c param is defined byVal *** " & Chr(13))
    ' note that the color variable gets changed in the sub below because it was passed byRef
    ' note that the change in the size variable made in the sub below is not reflected in this main sub because
    ' the parameter is defined in the sub below as a ByVal
    End

End Sub

Sub checkFruitDetails( _
    ByVal FType As String, _
    ByRef color As String, _
    Optional ByVal size) ' not this parameter is defined as optional which means it cn be omitted
               
    ' passing ByVal allows you to change the name. passing byVal allows the variable to change
    ' Ftype and color are considered parameters that are defined in the sub parenthesis
    ' when passing byRef names need to match
               '
    If FType = "Banana" Then
        'Change the color in this sub to show how this is possible to overide
        color = "green-yellow"
        size = "med"
    End If
    
    Debug.Print _
    "Values within the Sub" & Chr(13) _
    ; ("FType = " & FType & Chr(13) _
    & "color = " & color & "     *** Note how the value gets changed in the sub ***" & Chr(13) _
    & "size = " & size & "               *** Note how the value gets changed in the sub ***" & Chr(13))
   
End Sub