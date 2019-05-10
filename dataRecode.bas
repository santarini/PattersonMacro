Sub DataRecode()

Dim sourceSheet As Worksheet
Dim sourceRng, fundedValue, awardStart, contractValue, awardYear, awardQtr, statusColumn, plannedHeader, acutalHeader, dateHeader, inProgressHeader, submittedHeader As Range
Dim i, cellCount, qtrValue, yearValue As Integer
Dim dollarValue As Currency
Dim fullDate As Date


'if page contains CWPO
If (InStr(1, ActiveSheet.Name, "CWPO") > 0) Then

'set page as sourcePage
Set sourceSheet = ActiveSheet

'find cell with value "proposal status"
Set sourceRng = Cells.Find(What:="Proposal Status", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)

'find cell with value "Contract Funded Value"
Set fundedValue = Cells.Find(What:="Contract Funded Value", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)

'find cell with value "Award Start Date"
Set awardStart = Cells.Find(What:="Award Start Date", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)
 
'find cell with value "Contract Value"
Set contractValue = Cells.Find(What:="Contract Value", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)

'find cell with value "Projected Contract Award (Year)"
Set awardYear = Cells.Find(What:="Projected Contract Award (Year)", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)

'find cell with value "Projected Contract Award (Quarter)"
Set awardQtr = Cells.Find(What:="Projected Contract Award (Quarter)", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)

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

i = 1
'if page contains PPPS
For Each cell In statusColumn
    'if cell.value contains Closed Wonn
    If InStr(1, cell.Value, "Closed Won") > 0 Then
        dollarValue = fundedValue.Offset(i, 0).Value
        
    
        
        'create header "In Progress"
        'just right of that, create header "Submitted"
        'just right of that, create header "Date"
        'go to sourceRng
        'select entire column beneath soureRng, define as statusCol
    End If
    'if cell.value contains "Pipeline Opportunity"
    If InStr(1, cell.Value, "Pipeline Opportunity") > 0 Then
    End If
    'if cell.value contains Proposal In Progress
    If InStr(1, cell.Value, "Proposal In Progress") > 0 Then
    End If
    'if cell.value contains Proposal Submitted
    If InStr(1, cell.Value, "Proposal Submitted") > 0 Then
    End If
    i = i + 1
Next cell


'for each cell in status column


'if cell.value contains Pipeline Opportunity






End If

End Sub
