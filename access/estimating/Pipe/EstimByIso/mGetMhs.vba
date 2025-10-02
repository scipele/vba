Attribute VB_Name = "mGetMhs"
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | mGetMhs.vba                                                 |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple / BJR, 10/1/2025                                   |


Option Compare Database
Option Explicit
    ' Gather variables to help build the lookup code
    ' fErSp__std_10.0
    ' 12345     |1234
    '      12345|  |
    '   |    |  |  |--- Nominal Pipe Size truncated to 4 places w dec point
    '   |    |  |------ Space '_'
    '   |    |--------- Schedule padded with 5 total spaces
    '   |-------------- This is the string for the labor type

Public Enum RatingSchType
    schBased = 0
    sbRtgBased = 1
    lbRtgBased = 2
    noRtgOrSch = 3
End Enum

Public Type labData
    size_id As Long
    size As Double
    sz_str As String
    sch_str As String
    rtg As String
    rt_param As String
    code_str As String
    un_mh As Double
    qty As Double
End Type


Sub TotalMhForIso()
    
    Dim total_mh As Double
    
    total_mh = Nz(Forms!fe_data!instr_mh.Value) + _
               Nz(Forms!fe_data!sp_mh.Value) + _
               Nz(Forms!fe_data!tie_mh.Value) + _
               Nz(Forms!fe_data!supt_mh.Value) + _
               Nz(Forms!fe_data!grout_mh.Value)
               
    ' Declare variables
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim estID As Long ' Assuming est_id is a Long; adjust if needed
    
    On Error GoTo ErrorHandler
    
    ' Get est_id from the active form record
    ' Replace "YourFormName" with the actual name of your form
    ' Replace "est_id" with the actual name of the control/field on the form
    estID = Forms!fe_data!est_id
    If IsEmpty(estID) Or IsNull(estID) Then
        MsgBox "Error: No valid est_id selected in the form."
        Exit Sub
    End If
    
    ' Use the current database session
    Set db = CurrentDb
    
    ' Define the SQL query with dynamic est_id
    sql = "SELECT Sum(Nz([spool_mhs])+" & _
          "Nz([str_run_mhs])+" & _
          "Nz([butt_wld_mhs])+" & _
          "Nz([sw_mhs])+" & _
          "Nz([bu_mhs])+" & _
          "Nz([vlv_hnd_mhs])+" & _
          "Nz([make_on_mhs])+" & _
          "Nz([mo_bckwld_mhs])+" & _
          "Nz([cut_bev_mhs])) " & _
          "AS tot_mh " & _
          "FROM tb_qtys " & _
          "GROUP BY tb_qtys.est_id " & _
          "HAVING tb_qtys.est_id = " & estID & ";"
    
    ' Execute the query
    Set rs = db.OpenRecordset(sql)
    
    ' Check if records are returned
    If Not rs.EOF Then
        ' Convert the tot_mh (Double) to text using CStr
        total_mh = total_mh + CStr(rs("tot_mh"))
    End If
    
    'Round the number
    total_mh = Int(-10 * total_mh) / -10
    
    Dim total_mh_str As String
    total_mh_str = CStr(total_mh)
    
    Forms!fe_data!tbxTotMh.Value = total_mh_str
    
    ' Clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub


Function GetSpoolMhs(ByVal sz_id As Long)
    Dim ld As labData
    ld.size_id = sz_id
    ld.qty = Form_ff_qtys.spool_qty.Value
    GetSpoolMhs = GetSzSchRtgAndMhs("fErSp", ld, schBased)
End Function


Function GetStrMhs(ByVal sz_id As Long)
    Dim ld As labData
    ld.qty = Form_ff_qtys.str_run_qty.Value
    ld.size_id = sz_id
    GetStrMhs = GetSzSchRtgAndMhs("fErSr", ld, schBased)
End Function


Function GetBwMhs(ByVal sz_id As Long)
    Dim ld As labData
    ld.qty = Form_ff_qtys.butt_wld_qty.Value
    ld.size_id = IIf(sz_id < 5, 5, sz_id)
    GetBwMhs = GetSzSchRtgAndMhs("fBwld", ld, schBased)
End Function


Function GetSwMhs(ByVal sz_id As Long)
    Dim ld As labData
    ld.qty = Form_ff_qtys.sw_qty.Value
    ld.size_id = sz_id
    GetSwMhs = GetSzSchRtgAndMhs("fSwld", ld, sbRtgBased)
End Function


Function GetBuMhs(ByVal sz_id As Long)
    Dim ld As labData
    ld.qty = Form_ff_qtys.bu_qty.Value
    ld.size_id = sz_id
    GetBuMhs = GetSzSchRtgAndMhs("fBuFl", ld, lbRtgBased)
End Function


Function GetVhMhs(ByVal sz_id As Long)
    Dim ld As labData
    ld.qty = Form_ff_qtys.vlv_hnd_qty.Value
    ld.size_id = sz_id
    GetVhMhs = GetSzSchRtgAndMhs("fVlvh", ld, lbRtgBased)
End Function


Function GetMoMhs(ByVal sz_id As Long)
    Dim ld As labData
    ld.qty = Form_ff_qtys.make_on_qty.Value
    ld.size_id = sz_id
    GetMoMhs = GetSzSchRtgAndMhs("fMoTr", ld, noRtgOrSch)
End Function


Function GetMbMhs(ByVal sz_id As Long)
    Dim ld As labData
    ld.qty = Form_ff_qtys.mo_bckwld_qty.Value
    ld.size_id = sz_id
    GetMbMhs = GetSzSchRtgAndMhs("fMoBw", ld, noRtgOrSch)
End Function


Function GetCbMhs(ByVal sz_id As Long)
    Dim ld As labData
    ld.qty = Form_ff_qtys.cut_bev_qty.Value
    ld.size_id = IIf(sz_id < 8, 8, sz_id)

    Dim mh1 As Double, mh2 As Double
    
    mh1 = GetSzSchRtgAndMhs("fCutP", ld, schBased)
    mh2 = GetSzSchRtgAndMhs("fBevP", ld, schBased)
    GetCbMhs = mh1 + mh2
End Function


Private Function GetSzSchRtgAndMhs(labType As String, _
                                 ByRef ld As labData, _
                                 rt As RatingSchType) _
                                 As Double
    '1. Get the size and format
    ld.size = Nz(DLookup("size", "tx_sizes", "size_id = " & ld.size_id), "Not Found")
    ld.sz_str = PadStr(Format(ld.size, "0.00"), 4)
    
    Select Case rt
        Case schBased
            '2a. Get the schedule and format
            If IsNull(Form_ff_qtys.Parent!sch_id) Then
                MsgBox ("Please enter the missing schedule and re-enter quantity")
                Exit Function
            End If
            ld.sch_str = Nz(DLookup("sch", "tx_scheds", "sch_id = " & Form_ff_qtys.Parent!sch_id), "Not Found")
            If ld.sch_str = "40" And ld.size <= 10 Then ld.sch_str = "std"
            If ld.sch_str = "80" And ld.size <= 8 Then ld.sch_str = "xs"
            ld.rt_param = PadStr(ld.sch_str, 5)
        Case sbRtgBased
            '2b. Get the sb rating and format
            If IsNull(Form_ff_qtys.Parent!sb_rtg_id) Then
                MsgBox ("Please enter the missing small bore rating and re-enter quantity")
                Exit Function
            End If
            ld.rtg = Nz(DLookup("sb_rtg", "tx_sb_rtgs", "sb_rtg_id = " & Form_ff_qtys.Parent!sb_rtg_id), "Not Found")
            ld.rt_param = PadStr(ld.rtg, 5)
        Case lbRtgBased
            '2c. Get the flange rating and format
            If IsNull(Form_ff_qtys.Parent!flg_rtg_id) Then
                MsgBox ("Please enter the missing Large Bore Rating and re-enter")
                Exit Function
            End If
            ld.rtg = Nz(DLookup("flg_rtg", "tx_flg_rtgs", "flg_rtg_id = " & Form_ff_qtys.Parent!flg_rtg_id), "Not Found")
            ld.rtg = ld.rtg & "."
            ld.rt_param = PadStr(ld.rtg, 5)
        Case noRtgOrSch
            ld.rt_param = "_____"
    End Select
  
    '3. Combine other strings to create the library lookup code
    ld.code_str = labType & ld.rt_param & "_" & ld.sz_str
        
    '4. Look up the value from tx_mhs where code matches
    ld.un_mh = Nz(DLookup("un_mh", "tx_mhs", "code = '" & ld.code_str & "'"), 0)
    'Call debugPrintValues(ld)
    If IsNotInLibrary(ld.un_mh) Then Exit Function
     
    '5. Calculated total mhs
    GetSzSchRtgAndMhs = ld.qty * ld.un_mh
End Function


Private Function IsNotInLibrary(ByVal val As Double) As Boolean
    If val = 0 Then
        MsgBox ("unit man hour not found in table 'tx_mhs'")
       IsNotInLibrary = True
    Else
        IsNotInLibrary = False
    End If
End Function


Private Sub debugPrintValues(ByRef ld As labData)
     ' used for debugging purposes
     Debug.Print "ld.rt_param ", ld.rt_param
     Debug.Print "sz_str", ld.sz_str
     Debug.Print "code_str", ld.code_str
     Debug.Print "un_mh", ld.un_mh
End Sub
