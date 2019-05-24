Sub PivotTable()

Dim sheetNameStr As Variant
Dim cellCount As Integer
Dim uniqName As New Collection, a
Dim allNames() As Variant
Dim i As Long


'for each sheet ending in CWPO
Set sourceSheet = ActiveSheet

If (InStr(1, sourceSheet.Name, "CWPO") > 0) Then
    sheetNameStr = Split(sourceSheet.Name, "CWPO")
End If

'if it doeset not already exists, create new result page whose name is sourcePage.name Pivot CWPO
If sheetExists(sheetNameStr(0) & "Pivot") = False Then
    Sheets.Add.Name = sheetNameStr(0) & "Pivot"
    Set destSheet = ActiveSheet
Else
    'if it does exist re-define it
    Set destSheet = Sheets(sheetNameStr(0) & "Pivot")
End If

sourceSheet.Activate

'find cell with value "proposal status"
Set sourceRng = Cells.Find(What:="Proposal Status", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)
sourceRng.Select
Range(Selection, Selection.End(xlDown)).Select
RowCount = Selection.Rows.Count
sourceRng.Select
Range(Selection, Selection.End(xlToRight)).Select
colCount = Selection.Columns.Count
sourceRng.Resize(RowCount, colCount).Select
Set pivotSourceRange = Selection
SrcData = "'" & sourceSheet.Name & "'!" & pivotSourceRange.Address(ReferenceStyle:=xlR1C1)
sourceSheet.Range("A1").Select

'get all the unique names
'find the cell with "Dawson Capture Lead"
Cells.Find(What:="Dawson Capture Lead", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False).Offset(1, 0).Select

Set capLeadCol = Range(Selection, Selection.End(xlDown))

ReDim allNames(1 To RowCount) As Variant

i = 1
For Each cell In capLeadCol
    allNames(i) = cell.Value
i = i + 1
Next cell

On Error Resume Next
For Each a In allNames
   uniqName.Add a, a
Next

'insert pivot tables
destSheet.Activate
destSheet.Range("A1").Select

i = 1
For i = 1 To uniqName.Count
    'Cells((i * 15) - 14, 1) = uniqName(i)
    PvtDest = "'" & destSheet.Name & "'!" & destSheet.Range("A" & ((i * 15) - 13)).Address(ReferenceStyle:=xlR1C1)
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData, Version:=6).CreatePivotTable TableDestination:=PvtDest, TableName:="PivotTable" & i, DefaultVersion:=6
    destSheet.Activate
    With ActiveSheet.PivotTables("PivotTable" & i)
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable" & i).PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable" & i).RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable" & i).PivotFields("Date")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    ActiveSheet.PivotTables("PivotTable" & i).PivotFields("Date").AutoGroup
    ActiveSheet.PivotTables("PivotTable" & i).AddDataField ActiveSheet.PivotTables("PivotTable" & i).PivotFields("Planned"), "Sum of Planned", xlSum
    ActiveSheet.PivotTables("PivotTable" & i).AddDataField ActiveSheet.PivotTables("PivotTable" & i).PivotFields("Actual"), "Sum of Actual", xlSum
    ActiveSheet.PivotTables("PivotTable" & i).AddDataField ActiveSheet.PivotTables("PivotTable" & i).PivotFields("In Progress"), "Sum of In Progress", xlSum
    ActiveSheet.PivotTables("PivotTable" & i).AddDataField ActiveSheet.PivotTables("PivotTable" & i).PivotFields("Submitted"), "Sum of Submitted", xlSum
    ActiveSheet.PivotTables("PivotTable" & i).PivotFields("Dawson Capture Lead").Orientation = xlPageField
    ActiveSheet.PivotTables("PivotTable" & i).PivotFields("Dawson Capture Lead").Position = 1
    ActiveSheet.PivotTables("PivotTable" & i).PivotFields("Dawson Capture Lead").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable" & i).PivotFields("Dawson Capture Lead").CurrentPage = uniqName(i)
    ActiveSheet.PivotTables("PivotTable" & i).PivotFields("Date").PivotFilters.Add2 Type:=xlDateBetween, Value1:="12/31/2017", Value2:="1/1/2020"
    ActiveSheet.PivotTables("PivotTable" & i).PivotSelect "Years[All]", xlLabelOnly + xlFirstRow, True
    Selection.ShowDetail = True

Next i




End Sub
Function sheetExists(sheetToFind As String) As Boolean
    sheetExists = False
    For Each Sheet In Worksheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
End Function
