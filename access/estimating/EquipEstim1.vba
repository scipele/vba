' filename:     EquipEstim1.vba
'
' Purpose:      Read Estimate Data, Perform Calculations, and Output Results back to Table
'               1. Reads equipment estimate data 'table 'eq_data' into user defined type
'               2. Reads 'z_calc_data' table data from recordset into user defined type
'               3. Performs calculations and saves results into user defined type variables
'               4. Outputs the calculated results back to the table ('eq_data')
'
' Dependencies: Library - Microsoft Scripting Runtime required when early binding method is used
'               which is more efficient.
'
' By:  T.Sciple, 09/11/2024

Option Compare Database

Type est_data
    eq_id As Integer
    item_no As String
    eq_type As Long
    eq_desc As String
    hp As Double
    wt_lbs As Double
    driver_qty As Double
    mh_handle As Double
    mh_install As Double
    sub_mh_mill As Double
    misc_matl_cost As Double
End Type


Type calc_data
    cd_id As Long
    eq_type As String
    grade_yn As Boolean
    grout_yn As Boolean
    grout_matl As String
    param_desc As String
    base_val As Double
    base_mh As Double
    exponent_a As Double
    mill_base_mh As Double
    mill_base_val As Double
    mill_exponent As Double
    base_matl_cost As Double
End Type


Sub main()
                
    Dim ed() As est_data
    Dim cd() As calc_data
    
    '1.  Read equipment data into custom structure and assign the structure values to an array
    Call read_eq_data("eq_data", ed())
    
    '2.  read calculation data into custom structure and assign the structure values to an array
    Call read_calculation_data("z_calc_data", cd())
    
    '3.  perform estimate calculations
    Call perform_estim_calculations(ed(), cd())
    
    '4.  output the calculations to table
    Call output_calculations("eq_data", ed())
    
    MsgBox ("Run Completed")
    
    'cleanup
    Erase cd
    Erase ed
End Sub


'1.  Read Equipment Data
Private Sub read_eq_data(ByVal tblName As String, _
                        ByRef ed() As est_data)
                        
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim rs As Object  'rs is defined as a record set object that can get the table data
    Set rs = CurrentDb.OpenRecordset(tblName)
    
    ' Check if the recordset is empty
    If rs.EOF Then
        MsgBox "Recordset is empty."
        Exit Sub
    End If
    
    rs.MoveFirst
    ' Initialize the array to hold the data
    ReDim ed(0 To rs.RecordCount - 1)
    
    ' Loop through the records and store data in the array
    For i = 0 To rs.RecordCount - 1
        ed(i).eq_id = rs!eq_id
        ed(i).item_no = get_val_if_not_null(rs!item_no, s)
        ed(i).eq_type = get_val_if_not_null(rs!eq_type, n)
        ed(i).eq_desc = get_val_if_not_null(rs!eq_desc, s)
        ed(i).hp = get_val_if_not_null(rs!hp, n)
        ed(i).wt_lbs = get_val_if_not_null(rs!wt_lbs, n)
        ed(i).driver_qty = get_val_if_not_null(rs!driver_qty, n)
        If i < (rs.RecordCount - 1) Then rs.MoveNext
    Next i

    ' Close the recordset
    rs.Close
    Set rs = Nothing

    'Exit the sub if no error is encountered
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Descriptionee
    On Error Resume Next
    rs.Close
End Sub


'2.  Read calculation data into custom structure and assign the structure values to an array
Private Sub read_calculation_data(ByVal tblName As String, _
                        ByRef cd() As calc_data)
                        
    On Error GoTo ErrorHandler
    
    Dim rs As Object  'rs is defined as a record set object that can get the table data
    Set rs = CurrentDb.OpenRecordset(tblName)

    ' Ensure the recordset is not empty
    If Not rs.EOF And Not rs.BOF Then
        rs.MoveLast
        rs.MoveFirst
    Else
        MsgBox "Recordset is empty."
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    ' Initialize the array to hold the data
    ReDim cd(0 To rs.RecordCount - 1)
    
    ' Loop through the records and store data in the array
    Dim i As Long
    For i = 0 To (rs.RecordCount - 1)
        cd(i).cd_id = rs!cd_id
        cd(i).eq_type = rs!eq_type
        cd(i).grade_yn = rs!grade
        cd(i).grade_yn = rs!grout_yn
        cd(i).grout_matl = get_val_if_not_null(rs!grout_matl, "d")
        cd(i).param_desc = get_val_if_not_null(rs!param_desc, "s")
        cd(i).base_val = get_val_if_not_null(rs!base_val, "d")
        cd(i).base_mh = get_val_if_not_null(rs!base_mh, "d")
        cd(i).exponent_a = get_val_if_not_null(rs!exponent_a, "d")
        cd(i).mill_base_mh = get_val_if_not_null(rs!mill_base_mh, "d")
        cd(i).mill_base_val = get_val_if_not_null(rs!mill_base_val, "d")
        cd(i).mill_exponent = get_val_if_not_null(rs!mill_exponent, "d")
        cd(i).base_matl_cost = get_val_if_not_null(rs!base_matl_cost, "d")
        If i < (rs.RecordCount - 1) Then rs.MoveNext
    Next i
    
    ' Close the recordset
    rs.Close
    Set rs = Nothing

    'Exit the sub if no error is encountered
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description
    On Error Resume Next
    rs.Close
End Sub


'3.  perform estimate calculations
Sub perform_estim_calculations(ByRef ed() As est_data, _
                    ByRef cd() As calc_data)
                    
    Dim i As Long   'Loop index for the 'est' data
    
    'create a dictionary object to map the
    '   index of the z_calc_data [0-(recordCount-1)]
    '   with the actual 'cd_id' field
    Dim dict As Scripting.dictionary
    Set dict = New Scripting.dictionary     'Early Binding Method used
    
    Dim j As Long   'Loop index for the 'z_calc_data'
    For j = LBound(cd) To UBound(cd)
        dict.Add Key:=cd(j).cd_id, Item:=j
        'key = table 'z_calc_data' field 'cd_id'
        ', item = Index [0 to (recordCount-1)]
    Next j
   
    For i = LBound(ed) To UBound(ed)
        'set j to index of the equipment type using the cooresponding type id
        Dim cd_indx As Long
        'return the index given the eq_type id number which is saved in the dictionary object as the key
        cd_indx = dict(ed(i).eq_type)
        
        If cd_indx = 0 Then
            'zero the calculation if no type is identified
            ed(i).mh_install = 0
        Else
            'Calculate the install manhours
            If cd(cd_indx).param_desc = "hp" Then
                ed(i).mh_install = RoundUpToIncr(0.9 * cd(cd_indx).base_mh * (ed(i).hp / _
                                   cd(cd_indx).base_val) ^ cd(cd_indx).exponent_a, 5)
                ed(i).misc_matl_cost = RoundUpToIncr(0.9 * cd(cd_indx).base_matl_cost * _
                                       (ed(i).hp / cd(cd_indx).base_val) ^ cd(cd_indx).exponent_a, 5)
            Else
                'otherwise assume weight based calculation of manhours
                If cd(cd_indx).base_mh = 0 Then
                    ed(i).mh_install = 0
                Else
                    ed(i).mh_install = RoundUpToIncr(0.9 * cd(cd_indx).base_mh * _
                                       (ed(i).wt_lbs / cd(cd_indx).base_val) ^ cd(cd_indx).exponent_a, 5)
                    ed(i).misc_matl_cost = RoundUpToIncr(cd(cd_indx).base_matl_cost * _
                                           (ed(i).wt_lbs / cd(cd_indx).base_val) ^ cd(cd_indx).exponent_a, 5)
                End If
            End If
        
            'Calculate the handling hours assume 10% of installation hours
            ed(i).mh_handle = RoundUpToIncr(ed(i).mh_install / 0.9 - ed(i).mh_install, 5)
        
            'Calculate the millwright hours
            If (ed(i).driver_qty) > 0 And cd(cd_indx).mill_base_val > 0 Then
                ed(i).sub_mh_mill = RoundUpToIncr(ed(i).driver_qty * cd(cd_indx).mill_base_mh * _
                                    (ed(i).hp / cd(cd_indx).mill_base_val) ^ cd(cd_indx).mill_exponent, 5)
            Else
                ed(i).sub_mh_mill = 0
            End If
        End If
    Next i
End Sub


'4.  output the calculations to table
Sub output_calculations(ByVal tblName As String, _
                    ByRef ed() As est_data)
                    'Array Parameter is setup as estimate data
                    
    On Error GoTo ErrorHandler
    Dim last_record As Long
    Dim rs As Object  'rs is defined as a record set object that can get the table data
    Set rs = CurrentDb.OpenRecordset(tblName)
    
    rs.MoveFirst
    ' Loop through the records and write the array to the matching records
    Dim i As Long
    For i = 0 To rs.RecordCount - 1
        rs.Edit
        rs("mh_handle") = ed(i).mh_handle
        rs("mh_install") = ed(i).mh_install
        rs("sub_mh_mill") = ed(i).sub_mh_mill
        rs("misc_matl_cost") = ed(i).misc_matl_cost
        rs.Update
        If i < (rs.RecordCount - 1) Then rs.MoveNext
    Next i

    ' Close the recordset and release objects from memory
    rs.Close
    Set rs = Nothing

    'Exit the sub if no error is encountered
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description
    On Error Resume Next
    rs.Close
End Sub


Function get_val_if_not_null(ByVal rs_value As Variant, ByVal udt_type As String) As Variant
    'udt_type = "s" for string, "n" = numeric (integer, long, double)
    If IsNull(rs_value) Then
        If udt_type = "s" Then get_val_if_not_null = ""
        If udt_type = "n" Then get_val_if_not_null = 0
    Else
        get_val_if_not_null = rs_value
    End If
End Function


Function RoundUpToIncr(num As Double, increment As Double) As Double
    Dim remainder As Double
    remainder = num - Int(num / increment) * increment
    
    If Abs(remainder) < 0.00001 Then  ' epsilon check for floating-point precision
        RoundUpToIncr = num
    ElseIf num < 0 Then
        RoundUpToIncr = Int(num / increment) * increment
    Else
        RoundUpToIncr = (Int(num / increment) + 1) * increment
    End If
End Function