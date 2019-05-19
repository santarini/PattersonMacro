Sub Test1()

Dim cellCount As Integer
Dim arr As New Collection, a
Dim allNames() As Variant
Dim i As Long

Range("A1").Select
Set nameCol = Range(Selection, Selection.End(xlDown))
cellCount = nameCol.Rows.Count

ReDim allNames(1 To cellCount) As Variant

i = 1
For Each cell In nameCol
    allNames(i) = cell.Value
i = i + 1
Next cell

On Error Resume Next
For Each a In allNames
   arr.Add a, a
Next

For i = 1 To arr.Count
    Cells(i, 2) = arr(i)
Next



End Sub
