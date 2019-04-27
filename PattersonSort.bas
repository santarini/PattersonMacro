Sub PattersonSort()

Dim cell, titleRng As Range
Dim PMOrng, ITrng, TRAINrng, EMrng, IMSrng, AMrng As Range

Dim StartTime As Double
Dim SecondsElapsed As Double
StartTime = Timer

'create tabs

Sheets.Add.Name = "PMO Support"
Set PMOrng = Sheets("PMO Support").Range("A1")
Sheets.Add.Name = "Cyber-Intel"
Set ITrng = Sheets("Cyber-Intel").Range("A1")
Sheets.Add.Name = "Training"
Set TRAINrng = Sheets("Training").Range("A1")
Sheets.Add.Name = "Federal Health"
Set HEALTHrng = Sheets("Training").Range("A1")
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

'MsgBox cellCount





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
    
'sort for IT Cyber
    If InStr(1, cell.Value, "Health Svs - ") > 0 Then
        'MsgBox cell.Value
        Sheets("OpportunityDetails").Activate
        cell.Select
        Selection.End(xlToLeft).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
        Sheets("Cyber-Intel").Activate
        ITrng.Select
        ActiveSheet.Paste
        ITrng.Offset(1, 0).Select
        Set ITrng = Selection
    End If
    
'sort for Training
    If InStr(1, cell.Value, "Training -") > 0 Then
        'MsgBox cell.Value
        Sheets("OpportunityDetails").Activate
        cell.Select
        Selection.End(xlToLeft).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
        Sheets("Training").Activate
        TRAINrng.Select
        ActiveSheet.Paste
        TRAINrng.Offset(1, 0).Select
        Set TRAINrng = Selection
    End If

'sort for EM
    If InStr(1, cell.Value, "Training -") > 0 Then
        'MsgBox cell.Value
        Sheets("OpportunityDetails").Activate
        cell.Select
        Selection.End(xlToLeft).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
        Sheets("Training").Activate
        TRAINrng.Select
        ActiveSheet.Paste
        TRAINrng.Offset(1, 0).Select
        Set TRAINrng = Selection
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

SecondsElapsed = Round(Timer - StartTime, 2)
MsgBox cellCount & " data points successfully sorted in " & SecondsElapsed & " seconds", vbInformation


End Sub
