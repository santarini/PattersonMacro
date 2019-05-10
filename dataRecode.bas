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

'go to farthest filled cell to right plus one and define new headers
sourceRng.Select
Selection.End(xlToRight).Offset(0, 1).Select
Set plannedHeader = Selection
plannedHeader.Value = "Planned"

sourceRng.Select
Selection.End(xlToRight).Offset(0, 1).Select
Set acutalHeader = Selection
acutalHeader.Value = "Actual"

sourceRng.Select
Selection.End(xlToRight).Offset(0, 1).Select
Set dateHeader = Selection
dateHeader.Value = "Date"

End If


'if page contains CWPO
If (InStr(1, ActiveSheet.Name, "PPPS") > 0) Then

'set page as sourcePage
Set sourceSheet = ActiveSheet

'find cell with value "proposal status"
Set sourceRng = Cells.Find(What:="Proposal Status", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)

'go to farthest filled cell to right plus one and define new headers
sourceRng.Select
Selection.End(xlToRight).Offset(0, 1).Select
Set inProgressHeader = Selection
inProgressHeader.Value = "In Progress"

sourceRng.Select
Selection.End(xlToRight).Offset(0, 1).Select
Set submittedHeader = Selection
submittedHeader.Value = "Submitted"

sourceRng.Select
Selection.End(xlToRight).Offset(0, 1).Select
Set dateHeader = Selection
dateHeader.Value = "Date"

End If

If (InStr(1, ActiveSheet.Name, "CWPO") > 0) Or (InStr(1, ActiveSheet.Name, "PPPS") > 0) Then

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

i = 1
For Each cell In statusColumn
    'if cell.value contains Closed Wonn
    If InStr(1, cell.Value, "Closed Won") > 0 Then
        dollarValue = fundedValue.Offset(i, 0).Value
        acutalHeader.Offset(i, 0).Value = dollarValue
        fullDate = awardStart.Offset(i, 0).Value
        dateHeader.Offset(i, 0).Value = fullDate
    End If
    'if cell.value contains "Pipeline Opportunity"
    If InStr(1, cell.Value, "Pipeline Opportunity") > 0 Then
        dollarValue = contractValue.Offset(i, 0).Value
        plannedHeader.Offset(i, 0).Value = dollarValue
        yearValue = awardYear.Offset(i, 0).Value
        qtrValue = awardQtr.Offset(i, 0).Value
        fullDate = QtrYearToDate(qtrValue, yearValue)
        dateHeader.Offset(i, 0).Value = fullDate
    End If
    'if cell.value contains Proposal In Progress
    If InStr(1, cell.Value, "Proposal In Progress") > 0 Then
        dollarValue = contractValue.Offset(i, 0).Value
        inProgressHeader.Offset(i, 0).Value = dollarValue
        yearValue = awardYear.Offset(i, 0).Value
        qtrValue = awardQtr.Offset(i, 0).Value
        fullDate = QtrYearToDate(qtrValue, yearValue)
        dateHeader.Offset(i, 0).Value = fullDate
    End If
    'if cell.value contains Proposal Submitted
    If InStr(1, cell.Value, "Proposal Submitted") > 0 Then
        dollarValue = contractValue.Offset(i, 0).Value
        submittedHeader.Offset(i, 0).Value = dollarValue
        yearValue = awardYear.Offset(i, 0).Value
        qtrValue = awardQtr.Offset(i, 0).Value
        fullDate = QtrYearToDate(qtrValue, yearValue)
        dateHeader.Offset(i, 0).Value = fullDate
    End If
    i = i + 1
Next cell

End If

End Sub

Function QtrYearToDate(ByVal qtrValue As Integer, ByVal yearValue As Integer) As Date
Dim fullDate As Date
Dim proxyMonth As Integer

If qtrValue = 1 Then
proxyMonth = 1
End If

If qtrValue = 2 Then
proxyMonth = 4
End If

If qtrValue = 3 Then
proxyMonth = 7
End If

If qtrValue = 4 Then
proxyMonth = 10
End If

QtrYearToDate = DateSerial(yearValue, proxyMonth, 1)

End Function
