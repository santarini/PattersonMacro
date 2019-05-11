Sub generateReport()

Dim sourceSheet, destSheet As Worksheet
Dim sheetNameArr As Variant
Dim sheetNameStr As String

'create a report page
Sheets.Add.Name = "Report"
Set destSheet = ActiveSheet
destSheet.Move Before:=Sheets(2)

i = 1
For Each Sheet In Worksheets

If (InStr(1, Sheet.Name, "Pivot") > 0) Then

'define the source sheet
Set sourceSheet = Sheet

sourceSheet.Activate

'get the source page name until the word Pivot
sheetNameArr = Split(sourceSheet.Name, "Pivot")
sheetNameStr = sheetNameArr(0)

'If (InStr(1, Sheet.Name, "Pivot") > 0) Then
If IsEmpty(sourceSheet.Range("A1")) = False Then
    destSheet.Activate
    destSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    Application.CutCopyMode = False
    ActiveChart.SetSourceData Source:=sourceSheet.Range("A1")
    ActiveChart.ChartTitle.Text = sheetNameStr & " - Planned vs. Actual"
    ActiveChart.Parent.Cut
    Range("A" & i).Select
    ActiveSheet.Paste
'generate a graph
End If
If IsEmpty(sourceSheet.Range("F1")) = False Then
    destSheet.Activate
    destSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    Application.CutCopyMode = False
    ActiveChart.SetSourceData Source:=sourceSheet.Range("F1")
    ActiveChart.ChartTitle.Text = sheetNameStr & " - In Progress vs Submitted"
    ActiveChart.Parent.Cut
    Range("J" & i).Select
    ActiveSheet.Paste
End If
i = i + 20

End If

Next Sheet

End Sub
