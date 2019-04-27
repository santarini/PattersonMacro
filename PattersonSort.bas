Sub PattersonSort()

Dim cell, titleRng As Range
Dim PMOrng, EMrng, IMSrng, AMrng As Range

'create tabs

Sheets.Add.Name = "PMO Support"
Set PMOrng = Sheets("PMO Support").Range("A1")
'Sheets.Add.Name = "Cyber-Intel"
'Sheets.Add.Name = "Training"
'Sheets.Add.Name = "Federal Health"
Sheets.Add.Name = "CBRNE"
Set IMSrng = Sheets("CBRNE").Range("A1")
Sheets.Add.Name = "Inst Mission Spt"
Set IMSrng = Sheets("Inst Mission Spt").Range("A1")
Sheets.Add.Name = "Asset Mgmt"
Set AMrng = Sheets("Asset Mgmt").Range("A1")

'Identify aggregate opportunity list

Sheets("OpportunityDetails").Activate

'find cell containing "Title"
    Cells.Find(What:="Title", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False).Offset(1, 0).Select
'select all rows in Title column
Range(Selection, Selection.End(xlDown)).Select
cellCount = Selection.Rows.Count
Set titleRng = Selection

MsgBox cellCount





For Each cell In titleRng
'sort for PMO
    If InStr(1, cell.Value, "PMO -") > 0 Then
        'MsgBox cell.Value
        Sheets("OpportunityDetails").Activate
        cell.Select
        Selection.End(xlToLeft).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
        Sheets("PMO Support").Activate
        PMOrng.Select
        ActiveSheet.Paste
        PMOrng.Offset(1, 0).Select
        Set PMOrng = Selection
    End If
    
'sort for EM
    If InStr(1, cell.Value, "EM-CBRNE -") > 0 Then
        'MsgBox cell.Value
        Sheets("OpportunityDetails").Activate
        cell.Select
        Selection.End(xlToLeft).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
        Sheets("CBRNE").Activate
        EMrng.Select
        ActiveSheet.Paste
        EMrng.Offset(1, 0).Select
        Set EMrng = Selection
    End If

'sort for IMS
    If InStr(1, cell.Value, "IMS -") > 0 Then
        'MsgBox cell.Value
        Sheets("OpportunityDetails").Activate
        cell.Select
        Selection.End(xlToLeft).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
        Sheets("Inst Mission Spt").Activate
        IMSrng.Select
        ActiveSheet.Paste
        IMSrng.Offset(1, 0).Select
        Set IMSrng = Selection
    End If

'sort for AM
    If InStr(1, cell.Value, "AM -") > 0 Then
        'MsgBox cell.Value
        Sheets("OpportunityDetails").Activate
        cell.Select
        Selection.End(xlToLeft).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
        Sheets("Asset Mgmt").Activate
        AMrng.Select
        ActiveSheet.Paste
        AMrng.Offset(1, 0).Select
        Set AMrng = Selection
    End If

Next cell

'if row contains pharse
'select entire row
'copy row
'go to tab
'paste row in tab


End Sub
