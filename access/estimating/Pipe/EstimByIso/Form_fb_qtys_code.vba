Option Explicit
'| Item         | Documentation Notes                                         |
'|--------------|-------------------------------------------------------------|
'| Filename     | Form_fb_qtys_code.vba                                       |
'| EntryPoint   | varies                                                      |
'| Purpose      | calculate mhs                                               |
'| Inputs       | event driven                                                |
'| Outputs      | number of manhours                                          |
'| Dependencies | none                                                        |
'| By Name,Date | T.Sciple, 9/10/2025                                         |

Option Compare Database

Private Sub spool_qty_AfterUpdate()
    ' Gather variables to help build the lookup code
    ' fErSp__std_10.0
    ' 123451234512345
    '   |    |    |---- Nominal Pipe Size truncated to 4 places w dec point
    '   |    |--------- Schedule padded with 5 total spaces
    '   |-------------- This is the string for the labor type
    Dim lab_type As String
    lab_type = "fErSp"
    
    Dim spool_qty As Double
    spool_qty = Me.spool_qty.Value
    
    Dim sch_str As String
    sch_str = Nz(DLookup("sch", "tx_scheds", "sch_id = " & Me.Parent!sch_id), "Not Found")
    sch_str = PadStr(sch_str, 5)
    
    Dim size As Double
    size = Nz(DLookup("size", "tx_sizes", "size_id = " & Me.size_id.Value), "Not Found")
    
    Dim sz_str As String
    sz_str = PadStr(Format(size, "0.0"), 4)
    
    Dim code_str As String
    code_str = lab_type & sch_str & "_" & sz_str
    
    ' Look up the value from tx_mhs where code matches
    Dim un_mh As Double
    un_mh = Nz(DLookup("un_mh", "tx_mhs", "code = '" & code_str & "'"), 0)
    Me.spool_mhs.Value = spool_qty * un_mh
    Me.Requery
    
    ' used for debugging purposes
    ' Debug.Print Me.Parent!sch_id
    ' Debug.Print "sch ", sch_str
    ' Debug.Print "sz_str", sz_str
    ' Debug.Print "code_str", code_str
    ' Debug.Print "un_mh", un_mh
End Sub
