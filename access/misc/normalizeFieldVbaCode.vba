' filename:  normalizeFieldVbaCode.vba
'
' Purpose:
'   1. This Macro helps to normalize a field in an access database
'
' Dependencies:    Header Files            Class Name
'
' By:  T.Sciple, 8/29/2024

Option Compare Database
Option Explicit


Sub mainNormalize(tblName As String, existFld As String, newFldID As String, lkpTblName As String)

    'Remember time when macro starts
    Dim StartTime As Double
    StartTime = Timer
    
    'Create New Lookup Table
    Call createNewTable(lkpTblName, existFld)
    
    'Add New Field to Existing Table
    Call createNewFieldInExistTable(tblName, newFldID)
    
    'Get Column data that needs to be normalized
    Dim ColmData As Variant
    ColmData = getTableData(tblName, existFld)
    
    'Next Lookup the ID number from the dictionary
    Call createUniqWithDict(ColmData, tblName, lkpTblName, newFldID, existFld)
    
    'Determine how many seconds code took to run
    Dim SecondsElapsed As Double
    SecondsElapsed = Round(Timer - StartTime, 2)
    
    MsgBox ("Completed Normalization in " & SecondsElapsed & " seconds")
End Sub


Sub createNewTable(lkpTblName As String, fldName As String)
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fldID, fldNew As DAO.Field
     
    Set db = CurrentDb
    
    'Create Table Definition
    Set tdf = db.CreateTableDef(lkpTblName)
   
    'Create Field Definition
    Set fldID = tdf.CreateField("ID", dbLong)
    fldID.Attributes = dbAutoIncrField
    fldID.Required = True
    
    Set fldNew = tdf.CreateField(fldName, dbText)
    fldNew.AllowZeroLength = True
    fldNew.Required = False
        
    'Add the fields to the Table
    tdf.Fields.Append fldID
    tdf.Fields.Append fldNew
    
    'Add the Table to the Database
    db.TableDefs.Append tdf
    db.TableDefs.Refresh
    Application.RefreshDatabaseWindow
    
    Set fldID = Nothing
    Set fldNew = Nothing
    Set tdf = Nothing
    Set db = Nothing
End Sub


Sub createNewFieldInExistTable(tblName As String, fldName As String)
    Dim db As DAO.Database
    Set db = CurrentDb

    Dim tdf As DAO.TableDef
    Set tdf = db.TableDefs(tblName)
   
    With tdf
        .Fields.Append .CreateField(fldName, dbLong)
    End With
End Sub


Private Function getTableData(tblName, existFld)
    Dim numRecords, i As Long
    Dim rs As Object  'rs is defined as a record set object that can get the table data
    Dim tempAry As Variant
    
    Set rs = CurrentDb.OpenRecordset(tblName)
    'determines the number of records that will be used to set the 1st dimension of the array size
    numRecords = rs.RecordCount
        
    ReDim tempAry(1 To numRecords)
    
    For i = 1 To numRecords
        tempAry(i) = rs.Fields(existFld)
        rs.MoveNext
    Next i

    getTableData = tempAry
    'cleanup
    Erase tempAry
End Function


Private Sub createUniqWithDict(tmpAry, tblName, lkpTblName, newFldID, existFld)
    'This Sub uses a scripting dictionary to feed in all unique only values into a list that are defined as keys which will have a unique reference number
    
    Dim dic As Object
    Set dic = CreateObject("scripting.dictionary")

    Dim IdAry As Variant
    ReDim IdAry(LBound(tmpAry) To UBound(tmpAry))
    
    'create a scipting dictionary to store unique list as keys and a unique counter i
    Dim i As Long
    For i = LBound(tmpAry) To UBound(tmpAry)
        If Not dic.exists(tmpAry(i)) Then
            dic.Add Key:=tmpAry(i), Item:=dic.Count
        End If
    Next i

    'next loop thru the Scripting Dictionary Keys and place the key unique items into an Array then pass to sub
    Dim dicKeyAry As Variant
    ReDim dicKeyAry(1 To dic.Count)
    For i = 0 To dic.Count - 1
         dicKeyAry(i + 1) = dic.keys()(i)
    Next i
    
    'next sort the dictionary key array
    dicKeyAry = SortArray(dicKeyAry)
    
    'Place the sorted Array into Lookup Table
    Call PlaceAryInTableWithAdd(dicKeyAry, lkpTblName, existFld)

    'next redefine the dictionary in sorted order
    dic.RemoveAll
    
    'redefine dictionary in sorted order
    For i = LBound(dicKeyAry) To UBound(dicKeyAry)
         dic.Add Key:=dicKeyAry(i), Item:=dic.Count + 1
    Next i

   'next loop thru the original array and lookup the id in the dictionary
    For i = LBound(tmpAry) To UBound(tmpAry)
        IdAry(i) = dic.Item(tmpAry(i))
    Next i

    'Place the Lookup ID Array in the original Table
    Call PlaceAryInTableWithEdit(tblName, IdAry, newFldID)
End Sub


Private Sub PlaceAryInTableWithEdit(tblName, tmpAry, newFldID)
    Dim rs As Object
    Dim adder As Integer
    Set rs = CurrentDb.OpenRecordset(tblName)
    
    If tmpAry(LBound(tmpAry)) = 0 Then adder = 1
    Dim i As Long
    For i = LBound(tmpAry) To UBound(tmpAry)
        rs.Edit
        
        rs.Fields(newFldID) = tmpAry(i) + adder
        rs.Update
        rs.MoveNext
    Next i

End Sub


Private Sub PlaceAryInTableWithAdd(tmpAry, lkpTblName, fldName)
    Dim i As Long
    Dim rs As Object
    Set rs = CurrentDb.OpenRecordset(lkpTblName)
    
    For i = LBound(tmpAry) To UBound(tmpAry)
        rs.AddNew
        rs.Fields(fldName) = tmpAry(i)
        rs.Update
    Next i
End Sub


Function getTableList()
    Dim db As Database
    Dim tdf As TableDef
    Dim strList As String
    
    Set db = CurrentDb
    
    For Each tdf In db.TableDefs
        If Left(tdf.Name, 4) <> "MSys" Then
            strList = strList & tdf.Name & ";"
        End If
    Next tdf
    
   getTableList = Left(strList, Len(strList) - 1)
   Set tdf = Nothing
   db.Close
   Set db = Nothing
End Function


Function getFieldNamesForTable(tblName)
    Dim rs As DAO.Recordset
    Dim strFldList As String
    Dim i As Integer
    Dim list As Variant
    
    Set rs = CurrentDb.OpenRecordset(tblName)

    ' Create list of field names
    list = "FIELDS - [" & tblName & "]" & vbNewLine
    With rs
        For i = 0 To .Fields.Count - 1
            strFldList = strFldList & .Fields(i).Name & ";"
        Next i
    End With

    getFieldNamesForTable = strFldList
End Function


Function SortArray(arr As Variant)
    Dim temp As Variant
    Dim i As Long, j As Long

    For i = LBound(arr) To UBound(arr) - 1
    For j = i + 1 To UBound(arr)
        If arr(i) > arr(j) Then
        temp = arr(i)
        arr(i) = arr(j)
        arr(j) = temp
        End If
    Next j
    Next i

    SortArray = arr
End Function