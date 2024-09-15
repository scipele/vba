' This File Writes Data to a Structure from an excel sheet, and then writes the structure data to Binary Format that can be read by any programming language in the
' same order as it it written using a single byte number before every string to indicate length (<255 characters in this program, otherwise would need to change
' cbyte to  int)
'
'  Written by T. Sciple  6/30/2024

Option Explicit

Type estimData
    rowNo As Long
    filterX As String
    desc As String
    brkd_ref As String
    other_mh As Double
    param1 As String
    param2 As String
    param3 As String
    type As String
    area() As Double
    qty As Double
    uom As String
    umh As Double
    mh_tot As Double
    rate As Double
    labor As Double
    matl As Double
    sub As Double
    eq As Double
    total As Double
    div As String
    discp As String
    labtype As String
End Type

Sub ExportData()
    Dim fileNum As Integer
    Dim row As Integer
    Dim col As Integer
    Dim dataSheet As Worksheet
    Dim startRow As Integer
    Dim noColumns As Long
    Dim sheetName As String
    Dim headerRow As Long
    Dim lastfieldNameStr As String
    Dim fieldAfterLastArea As String
    Dim fieldBeforeFirstArea As String
    Dim noAreas As Integer
    Dim fieldNames As Variant
    Dim endColLtr As String
    Dim columnNos() As Integer
    Dim fldDict As New Dictionary
    Dim i As Long
    Dim filePath As String
    
    '1. Get the number data columns in the sheet except for the last Pct Column
    sheetName = "Estimate"
    lastfieldNameStr = "LabType"
    headerRow = 9
    noColumns = FindColumnByLabel(lastfieldNameStr, headerRow, sheetName)
    
    '2. Get the number of columns in the sheet then resize the structure array accordingly
    fieldAfterLastArea = "Qty"
    fieldBeforeFirstArea = "Type"
    noAreas = FindColumnByLabel(fieldAfterLastArea, headerRow, sheetName) - FindColumnByLabel(fieldBeforeFirstArea, headerRow, sheetName) - 1
    
    '3. Read in all the field names into a dictionary object
    endColLtr = Split(Cells(headerRow, noColumns).Address, "$")(1)
    fieldNames = ThisWorkbook.Sheets(sheetName).Range("a" & headerRow & ":" & endColLtr & headerRow)
    Set fldDict = CreateObject("scripting.dictionary")
    For i = 1 To UBound(fieldNames, 2)
        fldDict.Add fieldNames(1, i), i    'Add unique field name pairs to dictionary object
    Next i
    
    ' Set the file path for the binary file
    filePath = "C:\Users\mscip\cpp\excel\data.bin"
    
    Call FillStructureFromSheet(sheetName, headerRow, noAreas, fldDict, noColumns, filePath)
    
End Sub


Sub FillStructureFromSheet(ByVal shtName As String, _
                            ByVal headerRow As Long, _
                            ByVal noAreas As Integer, _
                            ByRef fldDict As Dictionary, _
                            ByVal noColumns As Long, _
                            ByVal filePath As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim ed() As estimData
    Dim numFields As Integer
    Dim lastTotalsRow As Long
    
    Set ws = ThisWorkbook.Sheets(shtName)
    lastTotalsRow = FindRowByLabel("TOTALS w/ all Options", fldDict("Rate"), shtName)
    
    ' Resize the array to hold all structures
    ReDim ed(headerRow + 1 To lastTotalsRow)
    For i = headerRow + 1 To lastTotalsRow
        ReDim ed(i).area(1 To noAreas)
    Next i

    For i = (headerRow + 1) To lastTotalsRow
        ' Read values from the spreadsheet into the structure max 255 characters for strings
        ed(i).rowNo = i
        ed(i).filterX = getMaxStr(ws.Cells(i, fldDict("X")).value)
        ed(i).desc = getMaxStr(ws.Cells(i, fldDict("Desc")).value)
        ed(i).brkd_ref = getMaxStr(ws.Cells(i, fldDict("Brkd_Ref")).value)
        ed(i).other_mh = ForceZeroIfNonNumeric(ws.Cells(i, fldDict("Other_Mh")).value)
        ed(i).param1 = getMaxStr(ws.Cells(i, fldDict("Param1")).value)
        ed(i).param2 = getMaxStr(ws.Cells(i, fldDict("Param2")).value)
        ed(i).param3 = getMaxStr(ws.Cells(i, fldDict("Param3")).value)
        ed(i).type = getMaxStr(ws.Cells(i, fldDict("Type")).value)
        ed(i).qty = ForceZeroIfNonNumeric(ws.Cells(i, fldDict("Qty")).value)
        ed(i).uom = getMaxStr(ws.Cells(i, fldDict("Uom")).value)
        ed(i).umh = ForceZeroIfNonNumeric(ws.Cells(i, fldDict("Umh")).value)
        ed(i).mh_tot = ForceZeroIfNonNumeric(ws.Cells(i, fldDict("Mh_Tot")).value)
        ed(i).rate = ForceZeroIfNonNumeric(ws.Cells(i, fldDict("Rate")).value)
        ed(i).labor = ForceZeroIfNonNumeric(ws.Cells(i, fldDict("Labor")).value)
        ed(i).matl = ForceZeroIfNonNumeric(ws.Cells(i, fldDict("Matl")).value)
        ed(i).sub = ForceZeroIfNonNumeric(ws.Cells(i, fldDict("Sub")).value)
        ed(i).eq = ForceZeroIfNonNumeric(ws.Cells(i, fldDict("Eq")).value)
        ed(i).total = ForceZeroIfNonNumeric(ws.Cells(i, fldDict("Total")).value)
        ed(i).div = getMaxStr(ws.Cells(i, fldDict("Div")).value)
        ed(i).discp = getMaxStr(ws.Cells(i, fldDict("Discp")).value)
        ed(i).labtype = getMaxStr(ws.Cells(i, fldDict("LabType")).value)
        
        ' Read Area fields into the array
        For j = 1 To noAreas
            ed(i).area(j) = ForceZeroIfNonNumeric(ws.Cells(i, fldDict("Type") + j).value)    'assumes that area if 1 after type
        Next j
    Next i

    Call WriteStructureToBinaryFile(ed(), filePath, noColumns)
End Sub


Function getMaxStr(ByVal str As String)

    Dim tmp As String
    Dim maxStrLenAllowed As Byte
    
    maxStrLenAllowed = 255
    
    tmp = str
    If Len(str) > maxStrLenAllowed Then
        tmp = Left(str, maxStrLenAllowed)
        getMaxStr = tmp
    Else
        getMaxStr = str
    End If

End Function


Sub WriteStructureToBinaryFile(ByRef ed() As estimData, _
                                ByVal filePath As String, _
                                ByVal noColumns As Long)
    Dim fileNum As Integer
    Dim i As Long, j As Long
    fileNum = FreeFile

    'delete any previous file
    On Error Resume Next
    Kill filePath
    MsgBox "Deleted Previous File: " & filePath, vbInformation

    ' Open binary file for writing
    Open filePath For Binary Access Write As #fileNum

    'Write the number of area fields
    Put #fileNum, , (UBound(ed(i).area()) - LBound(ed(i).area()))

    ' Write each structure individually
    For i = LBound(ed) To UBound(ed)
        ' Write the fixed part of the structure
        Put #fileNum, , ed(i).rowNo
        Put #fileNum, , CByte(Len(ed(i).filterX))
        Put #fileNum, , ed(i).filterX
        Put #fileNum, , CByte(Len(ed(i).desc))
        Put #fileNum, , ed(i).desc
        Put #fileNum, , CByte(Len(ed(i).brkd_ref))
        Put #fileNum, , ed(i).brkd_ref
        Put #fileNum, , ed(i).other_mh
        Put #fileNum, , CByte(Len(ed(i).param1))
        Put #fileNum, , ed(i).param1
        Put #fileNum, , CByte(Len(ed(i).param2))
        Put #fileNum, , ed(i).param2
        Put #fileNum, , CByte(Len(ed(i).param3))
        Put #fileNum, , ed(i).param3
        Put #fileNum, , CByte(Len(ed(i).type))
        Put #fileNum, , ed(i).type
        Put #fileNum, , ed(i).qty
        Put #fileNum, , CByte(Len(ed(i).uom))
        Put #fileNum, , ed(i).uom
        Put #fileNum, , ed(i).umh
        Put #fileNum, , ed(i).mh_tot
        Put #fileNum, , ed(i).rate
        Put #fileNum, , ed(i).labor
        Put #fileNum, , ed(i).matl
        Put #fileNum, , ed(i).sub
        Put #fileNum, , ed(i).eq
        Put #fileNum, , ed(i).total
        Put #fileNum, , CByte(Len(ed(i).div))
        Put #fileNum, , ed(i).div
        Put #fileNum, , CByte(Len(ed(i).discp))
        Put #fileNum, , ed(i).discp
        Put #fileNum, , CByte(Len(ed(i).labtype))
        Put #fileNum, , ed(i).labtype
        ' Write the number of Area fields
        Put #fileNum, , UBound(ed(i).area())

        ' Write the Area fields
        For j = 1 To UBound(ed(i).area)
            Put #fileNum, , ed(i).area(j)
        Next j
    Next i
    ' Close the file
    Close #fileNum
    
    MsgBox ("Binary File written and closed")
    
End Sub

Function FindColumnByLabel(ByVal label As String, _
                            ByVal searchRow As Long, _
                            ByVal shtName As String) _
                            As Long
    Dim ws As Worksheet
    Dim foundCell As Range

    ' Set the worksheet (adjust "Sheet1" to your worksheet name)
    Set ws = ThisWorkbook.Sheets(shtName)

    ' Initialize the function result
    FindColumnByLabel = -1
    
    ' Search for the label in the specified row
    Set foundCell = ws.Rows(searchRow).Find(What:=label, LookIn:=xlValues, LookAt:=xlWhole)

    ' Check if the cell was found and return the column number
    If Not foundCell Is Nothing Then
        FindColumnByLabel = foundCell.Column
    End If
End Function


Function FindRowByLabel(ByVal label As String, _
                        ByVal searchColumn As Long, _
                        ByVal shtName As String) _
                        As Long
    Dim ws As Worksheet
    Dim foundCell As Range

    ' Set the worksheet (adjust "Sheet1" to your worksheet name)
    Set ws = ThisWorkbook.Sheets(shtName)

    ' Initialize the function result
    FindRowByLabel = -1
    
    ' Search for the label in the specified column
    Set foundCell = ws.Columns(searchColumn).Find(What:=label, LookIn:=xlValues, LookAt:=xlWhole)

    ' Check if the cell was found and return the row number
    If Not foundCell Is Nothing Then
        FindRowByLabel = foundCell.row
    End If
End Function


Function ForceZeroIfNonNumeric(inputValue As Variant) As Double
    If IsNumeric(inputValue) Then
        ForceZeroIfNonNumeric = CDbl(inputValue)
    Else
        ForceZeroIfNonNumeric = 0
    End If
End Function