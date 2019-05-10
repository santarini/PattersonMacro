Sub generateReport()

Dim sourceSheet, destSheet As Worksheet
Dim sheetNameArr As Variant
Dim sheetNameStr As String


Set sourceSheet = ActiveSheet

'create a report page

Sheets.Add.Name = "Report"
Set destSheet = ActiveSheet

'sourceSheet.Activate

'for each sheet containing pivot
'For Each Sheet In Worksheets

'get the source page name until the word Pivot
sheetNameArr = Split(sourceSheet.Name, "Pivot")

sheetNameStr = sheetNameArr(0)

'make sure it's not the first page
'If (InStr(1, Sheet.Name, "Pivot") > 0) Then
If IsEmpty(sourceSheet.Range("A1")) = False Then
'''''''''''''''''''''''''
''''''''''''''''''''''''
    destSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    Application.CutCopyMode = False
    ActiveChart.SetSourceData Source:=sourceSheet.Range("A1")
    ActiveChart.ChartTitle.Text = sheetNameStr & " - Planned vs. Actual"
''''''''''''''''''''''''
''''''''''''''''''''''''''
'generate a graph
End If
If IsEmpty(sourceSheet.Range("F1")) = False Then
sourceSheet.Activate
'generate a graph
End If

'End If


End Sub
