Sub QtrComparison()

Dim AM_main_rng, AM_qtr_rng As Range

'create new sub work space
Sheets.Add.Name = "Asset Mgmt Qtr"

Set AM_qtr_rng = Sheets("Asset Mgmt").Range("A1")

Sheets("Asset Mgmt").Activate

Set AM_main_rng = Sheets("Asset Mgmt").Range("A1")

Range(Selection, Selection.End(xlDown)).Select
cellCount = Selection.Rows.Count
Set titleRng = Selection

i = 0
For Each cell In titleRng
'sort for PMO
    If InStr(1, cell.Value, "Closed Won") > 0 Then
        Sheets("Asset Mgmt Qtr").Activate
        cell.Select
        Selection.End(xlToLeft).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
        Sheets("PMO Support").Activate
        PMOrng.Select
        ActiveSheet.Paste
        PMOrng.Offset(1, 0).Select
        Set PMOrng = Selection
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
