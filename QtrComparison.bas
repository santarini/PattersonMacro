Sub QtrComparison()
Dim workingSourcePage, workingResultPage As Worksheet
Dim source_rng, AM_qtr_rng As Range
Dim testStr, year, monthRaw, month, qtr As String
Dim monthInt As Integer


'navigate to to main page
Set workingSourcePage = ActiveSheet
workingSourcePage.Activate

'define first cell in main page
Set source_rng = workingSourcePage.Range("A1")
source_rng.Select

'copy main header
Range(Selection, Selection.End(xlToRight)).Select
Selection.Copy

'create new sub work space
Sheets.Add.Name = workingSourcePage.Name & " Qtr"
Set workingResultPage = Sheets(workingSourcePage.Name & " Qtr")

'define first cell in work space
Set result_rng = workingResultPage.Range("A1")
result_rng.Select

'paste header
ActiveSheet.Paste
'offset rng
Set result_rng = workingResultPage.Range("A2")
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
'sort for Proposal Submitted
    If InStr(1, cell.Value, "Proposal Submitted") > 0 Then
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
workingSourcePage.Sort.SortFields.Clear
workingSourcePage.Sort.SortFields.Add2 Key:=Range( _
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
        If IsEmpty(cell.Offset(0, 5)) = True Then
            GoTo BlankCellError
        Else
            dateStr = cell.Offset(0, 5).Value
            
            year = Left(dateStr, InStr(dateStr, "/") - 1)
            monthRaw = Left(dateStr, InStr(dateStr, "/") + 2)
            month = Right(monthRaw, Len(monthRaw) - 5)
            
            monthInt = CInt(month)
            
            If 1 <= monthInt And monthInt <= 3 Then
            qtr = 1
            End If
            If 4 <= monthInt And monthInt <= 6 Then
            qtr = 2
            End If
            If 7 <= monthInt And monthInt <= 9 Then
            qtr = 3
            End If
            If 10 <= monthInt And monthInt <= 12 Then
            qtr = 4
            End If
        End If
    cell.Offset(0, 7).Value = year
    cell.Offset(0, 8).Value = qtr
    cell.Offset(0, 9).Value = "Actual"
BlankCellError:
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

'Delete
Columns("B:C").Select
Selection.Delete Shift:=xlToLeft
Range("A1").Select
Columns("D:D").Select
Selection.Delete Shift:=xlToLeft
Range("A1").Select

'format cells
Cells.Select
Range("Q19").Activate
Cells.EntireColumn.AutoFit

Range("A1").Select

End Sub


