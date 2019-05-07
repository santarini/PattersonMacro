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
Sheets("OpportunityDetails").Activate
Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Copy

'create new sub work space
Sheets.Add.Name = "Asset Mgmt Qtr"
Set workingResultPage = Sheets("Asset Mgmt Qtr")

'define first cell in work space
Set AM_qtr_rng = Sheets("Asset Mgmt Qtr").Range("A1")



'define working range
Range(Selection, Selection.End(xlDown)).Select
cellCount = Selection.Rows.Count
Set titleRng = Selection

'filter
i = 0
For Each cell In titleRng
'sort for PMO
    If InStr(1, cell.Value, "Closed Won") > 0 Then
        workingSourcePage.Activate
        cell.Select
        Selection.End(xlToLeft).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
        workingResultPage.Activate
        AM_qtr_rng.Select
        ActiveSheet.Paste
        AM_qtr_rng.Offset(1, 0).Select
        Set AM_qtr_rng = Selection
        i = i + 1
    End If
Next cell

'go through proposal status
'filter out closed won
'filter out pipeline opportunity
'filter out proposal submitted

'if proposed get proposed date
'if actual get actual date



End Sub
