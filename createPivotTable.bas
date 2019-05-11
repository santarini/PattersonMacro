Sub createPivotTable()

Dim sourceRng, proposalColumn, pivotSourceRange, pivotDateRange As Range
Dim sourceSheet, destSheet As Worksheet

Dim cellCount As Integer
Dim sheetNameStr As Variant
Dim SrcData, PvtDest As String
Dim pvtCache As PivotCache
Dim pvt As PivotTable

i = 1
For Each Sheet In Worksheets
'make sure it's not the first page
If (InStr(1, Sheet.Name, "CWPO") > 0) Or (InStr(1, Sheet.Name, "PPPS") > 0) Then

Sheet.Activate

'set as source page
'define sheet
Set sourceSheet = ActiveSheet

'get the source page name until abbreviation
If (InStr(1, Sheet.Name, "CWPO") > 0) Then
sheetNameStr = Split(sourceSheet.Name, "CWPO")
End If
If (InStr(1, Sheet.Name, "PPPS") > 0) Then
sheetNameStr = Split(sourceSheet.Name, "PPPS")
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

'get the last three columns from the data, save them as sourceDataRange
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
SrcData = "'" & sourceSheet.Name & "'!" & pivotSourceRange.Address(ReferenceStyle:=xlR1C1)
If (InStr(1, Sheet.Name, "CWPO") > 0) Then
    PvtDest = "'" & destSheet.Name & "'!" & destSheet.Range("A1").Address(ReferenceStyle:=xlR1C1)
End If
If (InStr(1, Sheet.Name, "PPPS") > 0) Then
    PvtDest = "'" & destSheet.Name & "'!" & destSheet.Range("F1").Address(ReferenceStyle:=xlR1C1)
End If


'Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)
'Set pvt = pvtCache.createPivotTable(TableDestination:=PvtDest, TableName:="PivotTable1")

'define source data space

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData, Version:=6).createPivotTable TableDestination:=PvtDest, TableName:="PivotTable" & i, DefaultVersion:=6
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
    
'add data lines
    
If (InStr(1, Sheet.Name, "CWPO") > 0) Then
    ActiveSheet.PivotTables("PivotTable" & i).PivotFields("Date").AutoGroup
    ActiveSheet.PivotTables("PivotTable" & i).AddDataField ActiveSheet.PivotTables("PivotTable" & i).PivotFields("Planned"), "Sum of Planned", xlSum
    ActiveSheet.PivotTables("PivotTable" & i).AddDataField ActiveSheet.PivotTables("PivotTable" & i).PivotFields("Actual"), "Sum of Actual", xlSum
    ActiveSheet.PivotTables("PivotTable" & i).PivotFields("Date").PivotFilters.Add2 Type:=xlDateBetween, Value1:="12/31/2017", Value2:="1/1/2020"
    ActiveSheet.PivotTables("PivotTable" & i).PivotSelect "Years[All]", xlLabelOnly + xlFirstRow, True
    Selection.ShowDetail = True
End If
If (InStr(1, Sheet.Name, "PPPS") > 0) Then
    'ActiveSheet.PivotTables("PivotTable" & i).PivotSelect "Months[All]", xlLabelOnly + xlFirstRow, True
    ActiveSheet.PivotTables("PivotTable" & i).AddDataField ActiveSheet.PivotTables("PivotTable" & i).PivotFields("In Progress"), "Sum of In Progress", xlSum
    ActiveSheet.PivotTables("PivotTable" & i).AddDataField ActiveSheet.PivotTables("PivotTable" & i).PivotFields("Submitted"), "Sum of Submitted", xlSum
    ActiveSheet.PivotTables("PivotTable" & i).PivotSelect "Date[All]", xlLabelOnly + xlFirstRow, True
    Selection.Group Start:=True, End:=True, Periods:=Array(False, False, False, False, False, True, True)
End If
 
End If

sourceSheet.Activate
sourceSheet.Range("A1").Select

i = i + 1
Next Sheet

End Sub
