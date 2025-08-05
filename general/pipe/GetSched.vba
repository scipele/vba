Public Function GetSched(orig_sch As String, _
                         nps As Double) _
                         As String
    
    
    If IsNull(orig_sch) Or orig_sch = "" Then
        GetSched = ""
        Exit Function
    End If
    
    ' Static collection to store NPS-to-OD mappings
    Static SchData As Collection
    Dim orig_sch_ary As Variant
    Dim new_sch_ary As Variant
    Dim i As Integer
    
    ' Initialize collection if not already done
    If SchData Is Nothing Then
        Set SchData = New Collection
        orig_sch_ary = Array("0.562""", "0.688""", "3000", "6000", "NULL", "S-120", "S-140", "S-160", "S-40", "S-40S", "S-80", "S-80S", "S-STD", "S-XS", "S-XXS")
        new_sch_ary = Array("0.562", ".688", "", "", "", "120", "140", "160", "40_Std_size", "40_Std_size", "80_xs_size", "80_xs_size", "std", "xs", "xxs")
        
        ' Populate collection with NPS as key (string for precise matching) and OD as value
        For i = LBound(orig_sch_ary) To UBound(orig_sch_ary)
            SchData.Add new_sch_ary(i), CStr(orig_sch_ary(i))
        Next i
    End If
    
    If SchData(CStr(orig_sch)) = "80_xs_size" Then
        If nps < 10 Then
            GetSched = "xs"
        Else
            GetSched = "80"
        End If
    End If

    If SchData(CStr(orig_sch)) = "40_Std_size" Then
        If nps < 12 Then
            GetSched = "std"
        Else
            GetSched = "40"
        End If
    End If
    
    ' Retrieve new_sch from collection
    On Error Resume Next
    If GetSched = "" Then
        GetSched = SchData(CStr(orig_sch))
    End If
    If Err.Number <> 0 Then
        ' Debug print parameters to Immediate Window
        Debug.Print "Error in GetSched: Err.Number = " & Err.Number & ", Err.Description = " & Err.Description
        Debug.Print "Parameters: orig_sch = '" & orig_sch & "', nps = " & nps
        GetSched = 0 ' Return 0 if NPS not found
    End If
    On Error GoTo 0
End Function
