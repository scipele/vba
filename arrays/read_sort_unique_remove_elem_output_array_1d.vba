Option Explicit
    
' read_sort_unique_remove_elem_output_array_1d
'
' This program illustrates how you can work with arrays to remove elements, sort, make them unique
' - place a list within the workbook where this code is placed that contains a named range for the list named 'myListRng'
' - create a named range desired output location named 'outputStartLoc'
'
' last update cleanup by Tony Sciple, 8/2/2024

Public Sub main()

    'Read the data from the range
    Dim myAry As Variant
    Call ReadNamedRangeToAry1D(myAry, "myListRng")
            
    'Call Sub to Sort the Array
    Call SortAryAtoZ(myAry)

    'Call Sub to Remove Duplicates
    Call MakeAryUnique(myAry)

    'Sub to Remove Specified Items from Array
    Dim items_to_remove As Variant
    items_to_remove = Array("-", "")
    Call Rem1dAryElements(myAry, items_to_remove)

    'Sub to place the Array starting at the nmed range 'outputStartLoc'
    Dim startLoc As String
    startLoc = "outputStartLoc"
    Call outputAryToNamedRangeTopLeft(myAry, startLoc)
    
    'clear array data from the heap memory allocation
    Erase myAry
End Sub


Sub outputAryToNamedRangeTopLeft(ByRef myAry As Variant, _
                                ByVal startLoc As String)
    
    Dim start_rng As Range
    Set start_rng = ThisWorkbook.Names(startLoc).RefersToRange
   
    Dim i As Long
    For i = LBound(myAry) To UBound(myAry)
        start_rng.Offset(i, 0).Value = myAry(i)
    Next i

End Sub


Sub ReadNamedRangeToAry1D(ByRef myAry As Variant, _
                              ByVal namedRangeStr As String)
                                
    'This Sub receives an array from caller passed by reference and re-dimensions the array
    ' to match the number of elements in the named range given

    ' Set the named range
    Dim namedRange As Range
    Set namedRange = ThisWorkbook.Names(namedRangeStr).RefersToRange
    
    'Determine the count of items for array sizing
    Dim count As Long
    count = namedRange.count
    ReDim myAry(0 To count - 1)
        
    'Loop thru each cell in the range and read the values into the re-dimensioned array
    Dim i As Long
    Dim cell As Variant
   
    For Each cell In namedRange
        myAry(i) = cell.Value
        i = i + 1
    Next cell
End Sub


Private Sub SortAryAtoZ(ByRef myAry As Variant)
    ' This is a basic bubble sort which is straight forward this is good enough relatively for small arrays,
    ' but consider more efficient sort algorithm for larger arrays
    Dim i As Long
    Dim j As Long
    Dim Temp

    'Sort the Array A-Z
    For i = LBound(myAry) To UBound(myAry) - 1
        For j = i + 1 To UBound(myAry)
            If UCase(myAry(i)) > UCase(myAry(j)) Then
                Temp = myAry(j)
                myAry(j) = myAry(i)
                myAry(i) = Temp
            End If
        Next j
    Next i
End Sub


Private Sub MakeAryUnique(ByRef myAry As Variant)
    ' This function will remove duplicates by using a dictionary object to be able to key if the key exist
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
 
    'populate the dictionary object with the unique values from 'myAry'
    Dim elem As Variant
    For Each elem In myAry
        If Not dict.exists(elem) Then
            dict.Add elem, Nothing  'Nothing refers to the fact that we are not storing any item but only the key
        End If
    Next elem
    
    'Create a temporary array to store the unique values of the array
    Dim tmpAry As Variant
    ReDim tmpAry(0 To dict.count - 1)
    Dim keyVal As Variant
    
    Dim i As Long
    i = 0
    For Each keyVal In dict
        tmpAry(i) = keyVal
        i = i + 1
    Next keyVal
        
    'Now reset the original array passed by reference to the temporary array
    myAry = tmpAry
End Sub


Private Sub Rem1dAryElements(ByRef myAry As Variant, _
                         ByRef items_to_remove As Variant)
    'Sub receives an array passed by reference and a string array list of element values to remove
   
    'dim and initialize a dictionary object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim elem As Variant
    
    'Setup other variables that will be used in the loop below
    Dim index As Long
    index = 1
    Dim item As Variant
    Dim add_flag As Boolean
    
    For Each elem In myAry
        'starting assumption is true unless otherwise set to false
        add_flag = True

        'nested loop to remove any items in the list
        For Each item In items_to_remove
            If elem = item Then add_flag = False
        Next item
        
        'if the flag was not set to false after checking thru all items them it shall be done
        If add_flag Then
            dict.Add index, elem
            index = index + 1
        End If
    Next elem
    
    'Now set the 'tmpAry' equal to the dictionary items
    Dim tmpAry As Variant
    tmpAry = dict.items
        
    'Next reset the array passed by reference to the temporary array
    myAry = tmpAry
    
    Erase tmpAry
End Sub