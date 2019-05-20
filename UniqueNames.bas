
Sub UniqueNames()

Dim cellCount As Integer
Dim uniqName As New Collection, a
Dim allNames() As Variant
Dim i As Long

'find the cell with "Dawson Capture Lead"
Cells.Find(What:="Dawson Capture Lead", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False).Offset(1, 0).Select


Set capLeadCol = Range(Selection, Selection.End(xlDown))
cellCount = capLeadCol.Rows.Count

ReDim allNames(1 To cellCount) As Variant

i = 1
For Each cell In capLeadCol
    allNames(i) = cell.Value
i = i + 1
Next cell

On Error Resume Next
For Each a In allNames
   uniqName.Add a, a
Next

Sheets.Add.Name = "TEST"
Set readyRespSheet = Sheets("TEST")
readyRespSheet.Activate
readyRespSheet.Range("A1").Select

For i = 1 To uniqName.Count
    Cells(i, 1) = uniqName(i)
Next

End Sub
