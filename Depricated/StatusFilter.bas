Sub StatusFilter()
'
'
'
    ActiveSheet.Range("A:A").AutoFilter Field:=1, Criteria1:=Array("Closed Won", "Pipeline Opportunity", "Proposal In Progress", "Proposal Submitted"), Operator:=xlFilterValues
    Range("A1").Select
    'Select to bottom
    Range(Selection, Selection.End(xlDown)).Select
    RowCount = Selection.Rows.Count
    'resize selection
    ActiveCell.Resize(RowCount, 70).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    'NavigateToNewSheet
    Sheets("Sheet8").Activate
    ActiveSheet.Paste
End Sub
