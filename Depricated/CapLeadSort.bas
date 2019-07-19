Sub CapLeadSort()

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' (c) Makoa Santarini - https://github.com/santarini/medleyMacro
' (c) Tim Medley
' (c) DAWSON Companies
'
' Patterson Sort for VBA
'
' @class CapLeadSort
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
Dim sourceRng, CapLeadCol As Range
Dim cellCount As Integer
Dim captureLeadSerial As String


'for each sheet in workbook
'if sheet is not "OpportunityDetials

'set source sheet
Set sourceSheet = Sheets("Aggregate")

'find cell with "Dawson Capture Lead"
Cells.Find(What:="Dawson Capture Lead", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False).Select

'select all rows beneath it
Range(Selection, Selection.End(xlDown)).Offset(1, 0).Select
cellCount = Selection.Rows.Count
ActiveCell.Resize(cellCount - 1, 1).Select
Set CapLeadCol = Selection

For Each cell In CapLeadCol
If cell <> "" Then
captureLeadName = cell.Value
'create captureLeadSerial
captureLeadFirstInital = Left(captureLeadName, 1)
captureLeadSpaceArray = Split(captureLeadName, " ")
captureLeadSerial = captureLeadFirstInital & Left(captureLeadSpaceArray(UBound(captureLeadSpaceArray)), 4)

'check to see if sheet exists whose name is source sheet name + captureLeadSerial
'if doens't exist create it and define it
If sheetExists(captureLeadSerial) = False Then
            'activate the source sheet
            sourceSheet.Activate
            'take the header from the source sheet
            sourceSheet.Range("A1:AY1").Copy
            'create the dest sheet
            Sheets.Add.Name = captureLeadSerial
            'define the dest sheet
            Set destSheet = Sheets(captureLeadSerial)
            'define a rng in the dest sheet
            destSheet.Range("A1").Select
            'paste the header at the rng
            ActiveSheet.Paste
        Else
        'if it does exist copy entire row and paste row into that sheet
            Set destSheet = Sheets(captureLeadSerial)
        End If
        'activate source sheet
        sourceSheet.Activate
        'get cell
        cell.Offset(0, -3).Select
        'Check proposal status
        If Selection = "Closed Won" Or Selection = "Pipeline Opportunity" Or Selection = "Proposal In Progress" Or Selection = "Proposal Submitted" Or Selection = "Sources Sought-RFI In Progress" Or Selection = "Sources Sought-RFI Submitted" Then
            'get entire row, copy
            ActiveCell.Resize(1, 12).Copy
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
            'Sort by proposal we want to see
            'ActiveSheet.ListObjects("Data").Range.AutoFilter Field:=1, Criteria1:=Array("Closed Won", "Pipeline Opportunity", "Proposal In Progress", "Proposal Submitted", "Sources Sought-RFI In Progress", "Sources Sought-RFI Submitted"), Operator:=xlFilterValues
            'clean up
            ActiveSheet.Range("A1").Select
        Else
        End If
    Else
End If
Next cell


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
