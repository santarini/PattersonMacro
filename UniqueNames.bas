Sub Test1()

Dim cellCount As Integer
Dim arr As New Collection, a
Dim aFirstArray() As Variant
Dim i As Long

Range("A1").Select
Set nameCol = Range(Selection, Selection.End(xlDown))
cellCount = nameCol.Rows.Count

ReDim aFirstArray(1 To cellCount) As Variant

i = 1
For Each cell In nameCol
    aFirstArray(i) = cell.Value
i = i + 1
Next cell

On Error Resume Next
For Each a In aFirstArray
   arr.Add a, a
Next

ReDim uniqueNames(1 To cellCount) As Variant

For i = 1 To arr.Count
   uniqueNames(i) = arr(i)
Next

batch = Join(uniqueNames, ",")

MsgBox batch


End Sub
