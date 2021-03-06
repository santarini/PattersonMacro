Sub StatusSort()

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

Dim sourceSheet, destSheet As Worksheet
Dim sourceSheetName, destSheetName As String
Dim sourceRng, proposalColumn As Range

'BEGIN OF FOR LOOP THROUGH WORKSHEETS
'for all tabs in the sheet
For Each Sheet In Worksheets
'create condition to make sure opperation avoids "OpportunityDetails" sheet
'and other conditons to avoide running opeartion on resulting sheets
'I was struggling with the or logic so I wrote 6 if statements (I realize this may not be the most efficient code)
If (InStr(1, Sheet.Name, "OpportunityDetails") = 0) Then
If (InStr(1, Sheet.Name, "CW") = 0) Then
If (InStr(1, Sheet.Name, "PO") = 0) Then
If (InStr(1, Sheet.Name, "PP") = 0) Then
If (InStr(1, Sheet.Name, "PS") = 0) Then

Sheet.Activate

'define sheet
Set sourceSheet = ActiveSheet

'get sheet name
'sourceSheet.Name

'find cell with value "proposal status"
Set sourceRng = Cells.Find(What:="Proposal Status", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)

'define the proposal column range
sourceRng.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Select
cellCount = Selection.Rows.Count
Set proposalColumn = Selection
sourceSheet.Range("A1").Select

For Each cell In proposalColumn
    'check for "Closed Won"
    If InStr(1, cell.Value, "Closed Won") > 0 Then
        'check if sheet exists
        If sheetExists(sourceSheet.Name & " CW") = False Then
            'activate the source sheet
            sourceSheet.Activate
            'take the header from the source sheet
            sourceSheet.Range("A1:AY1").Copy
            'create the dest sheet
            Sheets.Add.Name = sourceSheet.Name & " CW"
            'define the dest sheet
            Set destSheet = Sheets(sourceSheet.Name & " CW")
            'define a rng in the dest sheet
            destSheet.Range("A1").Select
            'paste the header at the rng
            ActiveSheet.Paste
        Else
            Set destSheet = Sheets(sourceSheet.Name & " CW")
        End If
        'activate source sheet
        sourceSheet.Activate
        'get cell
        cell.Activate
        'get entire row, copy
        ActiveCell.Resize(1, 70).Copy
        'go to destination sheet
        destSheet.Activate
        'select A1
        destSheet.Range("A1").Select
        'if A2 is blank select A
        If IsEmpty(Range("A2")) = True Then
            destSheet.Range("A2").Select
        Else
            'go to the bottom of the data + 1
            Selection.End(xlDown).Offset(1, 0).Select
        End If
        'paste
        ActiveSheet.Paste
        'clean up
        ActiveSheet.Range("A1").Select
    End If
    
    'check for "Pipeline Opportunity"
    If InStr(1, cell.Value, "Pipeline Opportunity") > 0 Then
        'check if sheet exists
        If sheetExists(sourceSheet.Name & " PO") = False Then
            'activate the source sheet
            sourceSheet.Activate
            'take the header from the source sheet
            sourceSheet.Range("A1:AY1").Copy
            'create the dest sheet
            Sheets.Add.Name = sourceSheet.Name & " PO"
            'define the dest sheet
            Set destSheet = Sheets(sourceSheet.Name & " PO")
            'define a rng in the dest sheet
            destSheet.Range("A1").Select
            'paste the header at the rng
            ActiveSheet.Paste
        Else
            Set destSheet = Sheets(sourceSheet.Name & " PO")
        End If
        'activate source sheet
        sourceSheet.Activate
        'get cell
        cell.Activate
        'get entire row, copy
        ActiveCell.Resize(1, 70).Copy
        'go to destination sheet
        destSheet.Activate
        'select A1
        destSheet.Range("A1").Select
        'if A2 is blank select A
        If IsEmpty(Range("A2")) = True Then
            destSheet.Range("A2").Select
        Else
            'go to the bottom of the data + 1
            Selection.End(xlDown).Offset(1, 0).Select
        End If
        'paste
        ActiveSheet.Paste
        'clean up
        ActiveSheet.Range("A1").Select
    End If
    
    'check for "Proposal In Progress"
    If InStr(1, cell.Value, "Proposal In Progress") > 0 Then
        'check if sheet exists
        If sheetExists(sourceSheet.Name & " PP") = False Then
            'activate the source sheet
            sourceSheet.Activate
            'take the header from the source sheet
            sourceSheet.Range("A1:AY1").Copy
            'create the dest sheet
            Sheets.Add.Name = sourceSheet.Name & " PP"
            'define the dest sheet
            Set destSheet = Sheets(sourceSheet.Name & " PP")
            'define a rng in the dest sheet
            destSheet.Range("A1").Select
            'paste the header at the rng
            ActiveSheet.Paste
        Else
            Set destSheet = Sheets(sourceSheet.Name & " PP")
        End If
        'activate source sheet
        sourceSheet.Activate
        'get cell
        cell.Activate
        'get entire row, copy
        ActiveCell.Resize(1, 70).Copy
        'go to destination sheet
        destSheet.Activate
        'select A1
        destSheet.Range("A1").Select
        'if A2 is blank select A
        If IsEmpty(Range("A2")) = True Then
            destSheet.Range("A2").Select
        Else
            'go to the bottom of the data + 1
            Selection.End(xlDown).Offset(1, 0).Select
        End If
        'paste
        ActiveSheet.Paste
        'clean up
        ActiveSheet.Range("A1").Select
    End If
    
    'check for "Proposal Submitted"
    If InStr(1, cell.Value, "Proposal Submitted") > 0 Then
        'check if sheet exists
        If sheetExists(sourceSheet.Name & " PS") = False Then
            'activate the source sheet
            sourceSheet.Activate
            'take the header from the source sheet
            sourceSheet.Range("A1:AY1").Copy
            'create the dest sheet
            Sheets.Add.Name = sourceSheet.Name & " PS"
            'define the dest sheet
            Set destSheet = Sheets(sourceSheet.Name & " PS")
            'define a rng in the dest sheet
            destSheet.Range("A1").Select
            'paste the header at the rng
            ActiveSheet.Paste
        Else
            Set destSheet = Sheets(sourceSheet.Name & " PS")
        End If
        'activate source sheet
        sourceSheet.Activate
        'get cell
        cell.Activate
        'get entire row, copy
        ActiveCell.Resize(1, 70).Copy
        'go to destination sheet
        destSheet.Activate
        'select A1
        destSheet.Range("A1").Select
        'if A2 is blank select A
        If IsEmpty(Range("A2")) = True Then
            destSheet.Range("A2").Select
        Else
            'go to the bottom of the data + 1
            Selection.End(xlDown).Offset(1, 0).Select
        End If
        'paste
        ActiveSheet.Paste
        'clean up
        ActiveSheet.Range("A1").Select
    End If
    
Next cell
End If
End If
End If
End If
End If
'END OF FOR LOOP THROUGH WORKSHEETS
Next Sheet

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
