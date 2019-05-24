Sub PivotTable()

Dim sheetNameStr As Variant
Dim cellCount As Integer
Dim uniqName As New Collection, a
Dim allNames() As Variant
Dim i As Long


'for each sheet ending in CWPO
Set sourceSheet = ActiveSheet

If (InStr(1, sourceSheet.Name, "CWPO") > 0) Then
    sheetNameStr = Split(sourceSheet.Name, "CWPO")
End If

'if it doeset not already exists, create new result page whose name is sourcePage.name Pivot CWPO
If sheetExists(sheetNameStr(0) & "Pivot") = False Then
    Sheets.Add.Name = sheetNameStr(0) & "Pivot"
    Set destSheet = ActiveSheet
Else
    'if it does exist re-define it
    Set destSheet = Sheets(sheetNameStr(0) & "Pivot")
End If

sourceSheet.Activate

'find cell with value "proposal status"
Set sourceRng = Cells.Find(What:="Proposal Status", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)
sourceRng.Select
Range(Selection, Selection.End(xlDown)).Select
RowCount = Selection.Rows.Count
sourceRng.Select
Range(Selection, Selection.End(xlToRight)).Select
colCount = Selection.Columns.Count
sourceRng.Resize(RowCount, colCount).Select
Set pivotSource = Selection

'get all the unique names
'find the cell with "Dawson Capture Lead"
Cells.Find(What:="Dawson Capture Lead", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False).Offset(1, 0).Select

Set capLeadCol = Range(Selection, Selection.End(xlDown))

ReDim allNames(1 To RowCount) As Variant

i = 1
For Each cell In capLeadCol
    allNames(i) = cell.Value
i = i + 1
Next cell

On Error Resume Next
For Each a In allNames
   uniqName.Add a, a
Next

'insert pivot tables
destSheet.Activate



End Sub
Function sheetExists(sheetToFind As String) As Boolean
    sheetExists = False
    For Each Sheet In Worksheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
End Function
