Sub PattersonSort()

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' (c) Makoa Santarini - https://github.com/santarini/PattersonMacro
'
' (c) DAWSON Companies
'
' Patterson Sort for VBA
'
' @class PattersonSort
' @author msantarini@dawson8a.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)

' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the <organization> nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Dim cell, titleRng As Range
Dim PMOrng, CYBERrng, TRAINrng, HEALTHrng, EMrng, IMSrng, AMrng As Range

Dim StartTime As Double
Dim SecondsElapsed As Double
StartTime = Timer

'create tabs

Sheets.Add.Name = "PMO Support"
Set PMOrng = Sheets("PMO Support").Range("A1")

Sheets.Add.Name = "Cyber-Intel"
Set CYBERrng = Sheets("Cyber-Intel").Range("A1")

Sheets.Add.Name = "Training"
Set TRAINrng = Sheets("Training").Range("A1")

Sheets.Add.Name = "Federal Health"
Set HEALTHrng = Sheets("Federal Health").Range("A1")

Sheets.Add.Name = "CBRNE"
Set EMrng = Sheets("CBRNE").Range("A1")

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

i = 0
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
        i = i + 1
    End If
    
    
'sort for Cyber
    If InStr(1, cell.Value, "IT_Cyber - ") > 0 Then
        'MsgBox cell.Value
        Sheets("OpportunityDetails").Activate
        cell.Select
        Selection.End(xlToLeft).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
        Sheets("Cyber-Intel").Activate
        CYBERrng.Select
        ActiveSheet.Paste
        CYBERrng.Offset(1, 0).Select
        Set CYBERrng = Selection
        i = i + 1
    End If
    
'sort for Training
    If InStr(1, cell.Value, "Training - ") > 0 Then
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
        i = i + 1
    End If
    
'sort for Health SVS
    If InStr(1, cell.Value, "Health Svs - ") > 0 Then
        'MsgBox cell.Value
        Sheets("OpportunityDetails").Activate
        cell.Select
        Selection.End(xlToLeft).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
        Sheets("Federal Health").Activate
        HEALTHrng.Select
        ActiveSheet.Paste
        HEALTHrng.Offset(1, 0).Select
        Set HEALTHrng = Selection
        i = i + 1
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
        i = i + 1
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
        i = i + 1
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
        i = i + 1
    End If

Next cell

SecondsElapsed = Round(Timer - StartTime, 2)
MsgBox i & " data points successfully sorted from " & cellCount & " in " & SecondsElapsed & " seconds", vbInformation

End Sub
