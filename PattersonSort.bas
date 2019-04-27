Sub PattersonSort()

Dim cell, titleRng As Range
Dim AMrng





'find cell containing "Title"
    Cells.Find(What:="Title", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False).Offset(1, 0).Select
'select all rows in Title column
Range(Selection, Selection.End(xlDown)).Select
cellCount = Selection.Rows.Count
Set titleRng = Selection

MsgBox cellCount

'create tabs
'Sheets.Add.Name = "PMO Support"
'Sheets.Add.Name = "Cyber-Intel"
'Sheets.Add.Name = "Training"
'Sheets.Add.Name = "Federal Health"
'Sheets.Add.Name = "CBRNE"
'Sheets.Add.Name = "Inst Mission Spt"
Sheets.Add.Name = "Asset Mgmt"
Set AMrng = Sheets("Asset Mgmt").Range("A1")

Sheets("OpportunityDetails").Activate

For Each cell In titleRng

    If InStr(1, cell.Value, "AM -") > 0 Then
        MsgBox cell.Value
        Sheets("OpportunityDetails").Activate
        cell.Select
        Selection.End(xlToLeft).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
        Sheets("Asset Mgmt").Activate
        AMrng.Select
        ActiveSheet.Paste
        AMrng.Offset(1, 0).Select
        AMrng = Selection
    End If
Next cell

'if row contains pharse
'select entire row
'copy row
'go to tab
'paste row in tab


End Sub
