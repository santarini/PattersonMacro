Sub QtrComparison()
Dim workingSourcePage, workingResultPage As Worksheet
Dim source_rng, AM_qtr_rng As Range

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
End Sub

Range("A1").Select

'delete from B to S
Columns("B:S").Select
Application.CutCopyMode = False
Selection.Delete Shift:=xlToLeft

Range("A1").Select

'delete from D to AA
Columns("D:AA").Select
Selection.Delete Shift:=xlToLeft
