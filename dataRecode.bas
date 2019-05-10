Sub DataRecode()

Dim sourceSheet As Worksheet
Dim sourceRng, fundedValue, awardStart, contractValue, awardYear, awardQtr, statusColumn As Range
Dim cellCount As Integer

'if page contains CWPO
If (InStr(1, ActiveSheet.Name, "CWPO") > 0) Then

'set page as sourcePage
Set sourceSheet = ActiveSheet

'find cell with value "proposal status"
Set sourceRng = Cells.Find(What:="Proposal Status", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)

sourceRng.Select

'find cell with value "Contract Funded Value"
Set fundedValue = Cells.Find(What:="Contract Funded Value", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)

fundedValue.Select

'find cell with value "Award Start Date"
Set awardStart = Cells.Find(What:="Award Start Date", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)

awardStart.Select


'find cell with value "Contract Value"
Set contractValue = Cells.Find(What:="Contract Value", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)

contractValue.Select

'find cell with value "Projected Contract Award (Year)"
Set awardYear = Cells.Find(What:="Projected Contract Award (Year)", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)

awardYear.Select

'find cell with value "Projected Contract Award (Quarter)"
Set awardQtr = Cells.Find(What:="Projected Contract Award (Quarter)", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)

awardYear.Select


'define the proposal column range
sourceRng.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Select
cellCount = Selection.Rows.Count
Set statusColumn = Selection
sourceRng.Select

'go to farthest filled cell to right plus one and define new headers
Selection.End(xlToRight).Offset(0, 1).Value = "Planned"
Selection.End(xlToRight).Offset(0, 1).Value = "Actual"
Selection.End(xlToRight).Offset(0, 1).Value = "Date"

'if page contains PPPS
For Each cell In statusColumn
'if cell.value contains Closed Won
'then


'if cell.value contains Proposal Submitted
'if cell.value contains Proposal In Progress
'go to farthest filled cell to right plus one
'create header "In Progress"
'just right of that, create header "Submitted"
'just right of that, create header "Date"
'go to sourceRng
'select entire column beneath soureRng, define as statusCol
Next cell


'for each cell in status column


'if cell.value contains Pipeline Opportunity






End If

End Sub
