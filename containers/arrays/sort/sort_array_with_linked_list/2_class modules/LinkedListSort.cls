' LinkedListSort class that manages the nodes
Option Explicit

Private head As node ' The head (first node) of the list
Private count As Long ' Count of nodes in the list


' Insert a value in order
Public Sub InsertInOrder(ByVal value As Variant)
    
    value = LCase(value) '<- convert all values to lowercase so that case is ignored
    Dim newNode As node
    Set newNode = New node
    newNode.value = value

    If head Is Nothing Then
        ' If the list is empty, newNode becomes the head
        Set head = newNode
    ElseIf head.value > value Then
        ' If newNode should be inserted before the head
        Set newNode.NextNode = head
        Set head = newNode
    Else
        ' Traverse the list to find the correct position
        Dim current As node
        Set current = head
        Do While Not current.NextNode Is Nothing
            If current.NextNode.value >= value Then Exit Do
            Set current = current.NextNode
        Loop

        ' Insert newNode in its sorted position
        Set newNode.NextNode = current.NextNode
        Set current.NextNode = newNode
    End If
    count = count + 1
End Sub


' class public Function or method to return the sorted array
Public Function return_sorted_ary() As Variant
    Dim current As node
    Set current = head

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Do While Not current Is Nothing
        dict.Add i, current.value
        i = i + 1
        Set current = current.NextNode
    Loop
    
    return_sorted_ary = dict.items
    Set dict = Nothing
End Function


' Get the count of nodes in the list
Public Property Get NodeCount() As Long
    NodeCount = count
End Property