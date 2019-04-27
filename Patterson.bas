Sub workspace()

Dim projectCode As String
Dim wb As Workbook
Dim rng, cell, codeRng As Range
Dim cellCount As Integer
Dim projCodes As New Dictionary



'Set wb = ActiveWorkbook

'take the active sheet and copy it to a new sheet
'ActiveSheet.Name = "RawData"
'wb.Sheets("RawData").Copy Before:=Sheets("RawData")
'ActiveSheet.Name = "RawDataCopy"

'Select top cell

Cells.Range("A1").Offset(1, 0).Select

'select all cells beneath
Range(Selection, Selection.End(xlDown)).Select
cellCount = Selection.Rows.Count
Set codeRng = Selection

i = 1
For Each cell In codeRng
    projCodes(i) = cell.Value
    i = i + 1
Next cell

MsgBox projCodes(1)


End Sub
