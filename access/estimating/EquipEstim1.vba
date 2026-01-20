Option Explicit
'| Item	        | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | EquipEstim1.vba                                             |
'| EntryPoint   | module - calcs, main sub                                    |
'| Purpose      | compute estimate work hours for various equipment types     |
'| Inputs       | read from the database tables                               |
'| Outputs      | number of work hours, matl costs                            |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 1/20/2026                                         |

Option Compare Database

'===============================================================================================
' User-defined data types
'===============================================================================================
Type est_data
    eq_id As Integer
    item_no As String
    eq_type As Long
    eq_desc As String
    hp As Double
    wt_lbs As Double
    d1_val As Double
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


'===============================================================================================
' Main entry point for equipment estimating calculations
'===============================================================================================
Sub main()
    
    '0.  Count the records
    Dim eq_data_record_cnt As Long
    eq_data_record_cnt = GetRecordCount("eq_data")
       
    '1.  Read equipment data into custom structure and assign the structure values to an array
    Dim est() As est_data
    Call read_eq_data("eq_data", est(), eq_data_record_cnt)
    
    '2.  read calculation data into custom structure and assign the structure values to an array
     Dim calc_data_record_cnt As Long
    calc_data_record_cnt = GetRecordCount("z_calc_data")
    
    Dim cd() As calc_data
    Call read_calculation_data("z_calc_data", cd(), calc_data_record_cnt)
    
    '3.  perform estimate calculations
    Call perform_estim_calculations(est(), cd(), eq_data_record_cnt)
    
    '4.  output the calculations to table
    Call output_calculations("eq_data", est(), eq_data_record_cnt)
    
    MsgBox ("Calculations Completed")
    
    'cleanup
    Erase cd
    Erase est
End Sub


'===============================================================================================
' Count the records in a table
'===============================================================================================
Function GetRecordCount(ByVal tbl_name As String)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim indx As Long
    
    Set db = CurrentDb
    ' Open a recordset using a query or table name
    
    Set rs = db.OpenRecordset("SELECT * FROM " & tbl_name)
    
    If Not rs.EOF Then
        rs.MoveLast  ' Move to the last record to populate the count property
        indx = rs.recordCount
        Debug.Print "Total records (DAO): " & indx
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    GetRecordCount = indx
End Function


'===============================================================================================
'1.  Read Equipment Data
'===============================================================================================
Private Sub read_eq_data(ByVal tblName As String, _
                        ByRef est() As est_data, _
                        ByVal eq_data_record_cnt As Long)
                        
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim maxId As Long
    Dim rs As Object  'rs is defined as a record set object that can get the table data
    
    ' First, get the maximum eq_id to properly size the array (handles gaps in autonumber)
    maxId = Nz(DMax("eq_id", tblName), 0)
    If maxId = 0 Then
        MsgBox "No records found or invalid eq_id values."
        Exit Sub
    End If
    
    Set rs = CurrentDb.OpenRecordset(tblName)
    
    ' Check if the recordset is empty
    If rs.EOF Then
        MsgBox "Recordset is empty."
        Exit Sub
    End If
    
    rs.MoveFirst
    
    ' Initialize the array to hold the data based on max ID (not record count)
    ReDim est(1 To maxId)
    
    ' Loop through the recordset directly (handles gaps in autonumber)
    Do While Not rs.EOF
        i = rs!eq_id
        est(i).eq_id = rs!eq_id
        If IsNull(rs!item_no) Then est(i).item_no = "" Else est(i).item_no = rs!item_no
        If IsNull(rs!eq_type) Then est(i).eq_type = 0 Else est(i).eq_type = rs!eq_type
        If IsNull(rs!eq_desc) Then est(i).eq_desc = "" Else est(i).eq_desc = rs!eq_desc4
        If IsNull(rs!hp) Then est(i).hp = 0 Else est(i).hp = rs!hp
        If IsNull(rs!wt_lbs) Then est(i).wt_lbs = 0 Else est(i).wt_lbs = rs!wt_lbs
        If IsNull(rs!d1_val) Then est(i).d1_val = 0 Else est(i).d1_val = rs!d1_val
        If IsNull(rs!driver_qty) Then est(i).driver_qty = 0 Else est(i).driver_qty = rs!driver_qty
        
        rs.MoveNext
    Loop

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


'===============================================================================================
'2.  Read calculation data into custom structure and assign the structure values to an array
'===============================================================================================
Private Sub read_calculation_data(ByVal tblName As String, _
                        ByRef cd() As calc_data, _
                        ByVal eq_data_record_cnt As Long)
                        
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim maxId As Long
    Dim rs As Object  'rs is defined as a record set object that can get the table data
    
    ' First, get the maximum cd_id to properly size the array (handles gaps in autonumber)
    maxId = Nz(DMax("cd_id", tblName), 0)
    If maxId = 0 Then
        MsgBox "No records found or invalid cd_id values."
        Exit Sub
    End If
    
    Set rs = CurrentDb.OpenRecordset(tblName)
    
    ' Check if the recordset is empty
    If rs.EOF Then
        MsgBox "Recordset is empty."
        Exit Sub
    End If
    
    rs.MoveFirst
    
    ' Initialize the array to hold the data based on max ID (not record count)
    ReDim cd(1 To maxId)
    
    ' Loop through the recordset directly (handles gaps in autonumber)
    Do While Not rs.EOF
        i = rs!cd_id
        cd(i).cd_id = rs!cd_id
        cd(i).eq_type = rs!eq_type
        cd(i).grade_yn = rs!grade
        cd(i).grout_yn = rs!grout_yn
        If IsNull(rs!grout_matl) Then cd(i).grout_matl = "" Else cd(i).grout_matl = rs!grout_matl
        If IsNull(rs!param_desc) Then cd(i).param_desc = "" Else cd(i).param_desc = rs!param_desc
        If IsNull(rs!base_val) Then cd(i).base_val = 0 Else cd(i).base_val = rs!base_val
        If IsNull(rs!base_mh) Then cd(i).base_mh = 0 Else cd(i).base_mh = rs!base_mh
        If IsNull(rs!exponent_a) Then cd(i).exponent_a = 0 Else cd(i).exponent_a = rs!exponent_a
        If IsNull(rs!mill_base_mh) Then cd(i).mill_base_mh = 0 Else cd(i).mill_base_mh = rs!mill_base_mh
        If IsNull(rs!mill_base_val) Then cd(i).mill_base_val = 0 Else cd(i).mill_base_val = rs!mill_base_val
        If IsNull(rs!mill_exponent) Then cd(i).mill_exponent = 0 Else cd(i).mill_exponent = rs!mill_exponent
        If IsNull(rs!base_matl_cost) Then cd(i).base_matl_cost = 0 Else cd(i).base_matl_cost = rs!base_matl_cost
        
        rs.MoveNext
    Loop

    ' Close the recordset
    rs.Close
    Set rs = Nothing

    'Exit the sub if no error is encountered
    Exit Sub

ErrorHandler:
    MsgBox "Sub 2 Error: " & Err.Description & "Loop Index = " & i
    On Error Resume Next
    rs.Close
    Set rs = Nothing
End Sub


'===============================================================================================
'3.  perform estimate calculations
'===============================================================================================
Sub perform_estim_calculations(ByRef est() As est_data, _
                    ByRef cd() As calc_data, _
                        ByVal eq_data_record_cnt As Long)
                    
    Dim i As Long
    Dim j As Long
   
    For i = LBound(est) To UBound(est)
        'set j to index of the equipment type using the cooresponding
        
        j = est(i).eq_type
       
        'used for troubleshooting
        'If i = 20 Then
        '    i = 20
        'End If
        
        If j = 0 Then
            'zero the calculation if not type is identified
            est(i).mh_install = 0
        Else
            'Calculate the install manhours
            Dim mh_type As String
            
            
            Select Case cd(j).param_desc
            
            Case "hp"
                est(i).mh_install = misc.RoundUpToIncrement(0.9 * cd(j).base_mh * (est(i).hp / cd(j).base_val) ^ cd(j).exponent_a, 5)
                est(i).misc_matl_cost = misc.RoundUpToIncrement(cd(j).base_matl_cost * (est(i).hp / cd(j).base_val) ^ cd(j).exponent_a, 5)
            Case "cf"
                If cd(j).base_mh = 0 Then
                    est(i).mh_install = 0
                Else
                    est(i).mh_install = misc.RoundUpToIncrement(0.9 * cd(j).base_mh * (est(i).wt_lbs / cd(j).base_val) ^ cd(j).exponent_a, 5)
                    est(i).misc_matl_cost = misc.RoundUpToIncrement(cd(j).base_matl_cost * (est(i).wt_lbs / cd(j).base_val) ^ cd(j).exponent_a, 5)
                End If
            
            Case "wt"
                If cd(j).base_mh = 0 Then
                    est(i).mh_install = 0
                Else
                    est(i).mh_install = misc.RoundUpToIncrement(0.9 * cd(j).base_mh * (est(i).wt_lbs / cd(j).base_val) ^ cd(j).exponent_a, 5)
                    est(i).misc_matl_cost = misc.RoundUpToIncrement(cd(j).base_matl_cost * (est(i).wt_lbs / cd(j).base_val) ^ cd(j).exponent_a, 5)
                End If
            Case Else
                'MsgBox ("invalid estim param, wt, cf, hp provided")
            
            End Select
            
            'Calculate the handling hours assume 10% of installation hours
            est(i).mh_handle = misc.RoundUpToIncrement(est(i).mh_install / 0.9 - est(i).mh_install, 5)
        
            'Calculate the millwright hours assume
            If (est(i).driver_qty) > 0 And cd(j).mill_base_val > 0 Then
                est(i).sub_mh_mill = misc.RoundUpToIncrement(est(i).driver_qty * cd(j).mill_base_mh * (est(i).hp / cd(j).mill_base_val) ^ cd(j).mill_exponent, 5)
            Else
                est(i).sub_mh_mill = 0
            End If
        End If
        
    Next i
    
    'Exit the sub if no error is encountered
    Exit Sub

ErrorHandler:
    MsgBox "Sub 3 Error: " & Err.Description
    On Error Resume Next
    rs.Close
    Set rs = Nothing
End Sub


'===============================================================================================
'4.  output the calculations to table
'===============================================================================================
Sub output_calculations(ByVal tblName As String, _
                    ByRef est() As est_data, _
                        ByVal eq_data_record_cnt As Long)
                    'Array Parameter is setup as estimate data
                    
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
    
    ' Loop through the recordset directly (handles gaps in autonumber)
    Do While Not rs.EOF
        i = rs!eq_id
        
        ' Skip if the array element is empty (eq_id = 0 means it was a gap)
        If est(i).eq_id > 0 Then
            rs.Edit
            rs("mh_handle") = est(i).mh_handle
            rs("mh_install") = est(i).mh_install
            rs("sub_mh_mill") = est(i).sub_mh_mill
            rs("misc_matl_cost") = est(i).misc_matl_cost
            rs.Update
        End If
        
        rs.MoveNext
    Loop

    ' Close the recordset and release objects from memory
    rs.Close
    Set rs = Nothing

    'Exit the sub if no error is encountered
    Exit Sub

ErrorHandler:
    MsgBox "Sub 4 Error: " & Err.Description
    On Error Resume Next
    rs.Close
    Set rs = Nothing
End Sub