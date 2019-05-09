Sub createPivotTable()
'
' Macro1 Macro
'

'define source data space
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets.Add

'create pivot table
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Asset Mgmt CW!R1C1:R18C51", Version:=6).CreatePivotTable TableDestination:="Sheet30!R3C1", TableName:="PivotTable1", DefaultVersion:=6
    Sheets("Sheet30").Select
    Cells(3, 1).Select
    
'give attributes to pivot table
    With ActiveSheet.PivotTables("PivotTable1")
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
    
    With ActiveSheet.PivotTables("PivotTable1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    
    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
    ActiveWindow.SmallScroll Down:=-42
    
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Award Start Date")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Award Start Date").AutoGroup
    ActiveWindow.SmallScroll Down:=0
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables("PivotTable1").PivotFields("Contract Planned Value"), "Sum of Contract Planned Value", xlSum
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables("PivotTable1").PivotFields("Contract Funded Value"), "Sum of Contract Funded Value", xlSum
End Sub
