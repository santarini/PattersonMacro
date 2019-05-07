Sub QtrComparison()
Dim workingSourcePage, workingResultPage As Worksheet
Dim source_rng, AM_qtr_rng As Range
Dim year As String


'navigate to to main page
Set workingSourcePage = Sheets("Asset Mgmt")
workingSourcePage.Activate

'define first cell in main page
Set source_rng = workingSourcePage.Range("A1")
source_rng.Select

'copy main header
Range(Selection, Selection.End(xlToRight)).Select
Selection.Copy

'create new sub work space
Sheets.Add.Name = "Asset Mgmt Qtr"
Set workingResultPage = Sheets("Asset Mgmt Qtr")

'define first cell in work space
Set result_rng = Sheets("Asset Mgmt Qtr").Range("A1")
result_rng.Select

'paste header
ActiveSheet.Paste
'offset rng
Set result_rng = Sheets("Asset Mgmt Qtr").Range("A2")
result_rng.Select

'navigate to to main page
workingSourcePage.Activate
source_rng.Select

'define working range
Range(Selection, Selection.End(xlDown)).Select
cellCount = Selection.Rows.Count
Set titleRng = Selection

'filter
i = 0
For Each cell In titleRng
'sort for Closed Won
    If InStr(1, cell.Value, "Closed Won") > 0 Then
        workingSourcePage.Activate
        cell.Select
        Selection.End(xlToLeft).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
        workingResultPage.Activate
        result_rng.Select
        ActiveSheet.Paste
        result_rng.Offset(1, 0).Select
        Set result_rng = Selection
        i = i + 1
    End If
'sort for Pipeline Opportunity
    If InStr(1, cell.Value, "Pipeline Opportunity") > 0 Then
        workingSourcePage.Activate
        cell.Select
        Selection.End(xlToLeft).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
        workingResultPage.Activate
        result_rng.Select
        ActiveSheet.Paste
        result_rng.Offset(1, 0).Select
        Set result_rng = Selection
        i = i + 1
    End If
'sort for Proposal in Progress
    If InStr(1, cell.Value, "Proposal In Progress") > 0 Then
        workingSourcePage.Activate
        cell.Select
        Selection.End(xlToLeft).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
        workingResultPage.Activate
        result_rng.Select
        ActiveSheet.Paste
        result_rng.Offset(1, 0).Select
        Set result_rng = Selection
        i = i + 1
    End If
Next cell

workingResultPage.Activate

Range("A1").Select

'delete from B to C
Columns("B:C").Select
Application.CutCopyMode = False
Selection.Delete Shift:=xlToLeft

Range("A1").Select

'delete from C to Q
Columns("C:Q").Select
Application.CutCopyMode = False
Selection.Delete Shift:=xlToLeft

Range("A1").Select

'delete from E to V
Columns("E:V").Select
Application.CutCopyMode = False
Selection.Delete Shift:=xlToLeft

Range("A1").Select

'convert to currency
Columns("K:L").Select
Selection.Style = "Currency"

Range("A1").Select

'move data
Columns("C:D").Select
Selection.Cut
Columns("B:B").Select
Selection.Insert Shift:=xlToRight

Range("A1").Select

'move data some more

Columns("K:M").Select
Selection.Cut
Range("D1").Select
Selection.Insert Shift:=xlToRight

Range("A1").Select

'delete H to P
Columns("H:P").Select
Selection.Delete Shift:=xlToLeft

'Add columns
Range("H1").Value = "Useable Year"
Range("I1").Value = "Useable Qtr"
Range("J1").Value = "Proj/Actual"

'define working range
Range(Selection, Selection.End(xlDown)).Select
cellCount = Selection.Rows.Count
Set titleRng = Selection

'sort data
Range("A1").Select
Columns("A:A").Select
ActiveWorkbook.Worksheets("Asset Mgmt Qtr").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Asset Mgmt Qtr").Sort.SortFields.Add2 Key:=Range( _
    "A2:A26"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
With ActiveWorkbook.Worksheets("Asset Mgmt Qtr").Sort
    .SetRange Range("A1:G26")
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
Range("A1").Select

'filter
For Each cell In titleRng
'sort for Closed Won
    If InStr(1, cell.Value, "Closed Won") > 0 Then
    'take the f value
    'figure out the year and qtr
    'populate H and I accordingly
    End If
'sort for Pipeline Opportunity
    If InStr(1, cell.Value, "Pipeline Opportunity") > 0 Then
    year = cell.Offset(0, 1).Value
    qtr = cell.Offset(0, 2).Value
    cell.Offset(0, 7).Value = year
    cell.Offset(0, 8).Value = qtr
    cell.Offset(0, 9).Value = "Projected"
    End If
'sort for Proposal in Progress
    If InStr(1, cell.Value, "Proposal In Progress") > 0 Then
    year = cell.Offset(0, 1).Value
    qtr = cell.Offset(0, 2).Value
    cell.Offset(0, 7).Value = year
    cell.Offset(0, 8).Value = qtr
    cell.Offset(0, 9).Value = "Projected"
    End If
Next cell



'format cells
Cells.Select
Range("Q19").Activate
Cells.EntireColumn.AutoFit

Range("A1").Select

End Sub
