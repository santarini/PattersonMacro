Sub Macro1()
'
' Macro1 Macro
'

'
    Range("A1:BD131").Select
    Range("D6").Activate
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Logistics CWPO!R1C1:R131C56", Version:=6).CreatePivotTable TableDestination _
        :="Sheet8!R3C1", TableName:="PivotTable1", DefaultVersion:=6
    Sheets("Sheet8").Select
    Cells(3, 1).Select
    
    Range("A3").Select
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Planned"), "Sum of Planned", xlSum
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Actual"), "Sum of Actual", xlSum
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("In Progress"), "Sum of In Progress", xlSum
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Submitted"), "Sum of Submitted", xlSum
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Date")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Date").AutoGroup
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Dawson Capture Lead")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Dawson Capture Lead"). _
        ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Dawson Capture Lead"). _
        CurrentPage = "Bob McGhin"
End Sub
