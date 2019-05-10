Sub createPivotTable()

Dim sourceRng, proposalColumn, pivotSourceRange As Range
Dim cellCount As Integer



'if page contains CWPO
'set as source page

'define sheet
Set sourceSheet = ActiveSheet

'get the source page name until CWPO
'get the last three columns from the data, save them as sourceDataRange
'create new result page whose name is sourcePage.name Pivot CWPO
'

'find cell with value "proposal status"
Set sourceRng = Cells.Find(What:="Proposal Status", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)

'define the proposal column range
If sourceRng.Offset(2, 0) = "" Then
    cellCount = 1
Else
    sourceRng.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    cellCount = Selection.Rows.Count
End If
    sourceRng.Select

'define pivotDataRange
Selection.End(xlToRight).Offset(0, -2).Select
Range(Selection, Selection.End(xlToRight)).Select
ActiveCell.Resize(cellCount + 1, 3).Select
Set pivotSourceRange = Selection



'define source data space
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets.Add

    Range("AZ1:BB26").Select
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Asset Mgmt CWPO!R1C52:R26C54", Version:=6).createPivotTable TableDestination:="Sheet72!R3C1", TableName:="PivotTable4", DefaultVersion:=6
    Sheets("Sheet72").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable4")
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
    With ActiveSheet.PivotTables("PivotTable4").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable4").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("Date")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable4").PivotFields("Date").AutoGroup
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables("PivotTable4").PivotFields("Planned"), "Sum of Planned", xlSum
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables("PivotTable4").PivotFields("Actual"), "Sum of Actual", xlSum
End Sub
