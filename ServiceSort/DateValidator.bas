Sub DateValidator()

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

Dim sourceRng As Range

For Each Sheet In Worksheets
'make sure it's not the opportunity details
If Sheet.Name <> "OpportunityDetails" Then
'make sure its one of the new sheets
If (InStr(1, Sheet.Name, "CW") > 0) Or (InStr(1, Sheet.Name, "PO") > 0) Or (InStr(1, Sheet.Name, "PP") > 0) Or (InStr(1, Sheet.Name, "PS") > 0) Then
Sheet.Activate

dateValidate ("RFP Release")

dateValidate ("RFP Due")

dateValidate ("Award Start Date")

End If
End If
Next Sheet
End Sub

Function dateValidate(searchTerm As String)
Dim sourceRng, dateColumn As Range
Dim dateInCell As Date
Dim cellCount As Integer


'find the cell with value "searchTerm"
Set sourceRng = Cells.Find(What:=searchTerm, After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)

If IsEmpty(sourceRng.Offset(2, 0)) = True Then
    sourceRng.Offset(1, 0).Select
    Set dateColumn = Selection
Else
    'select column beneath it
    sourceRng.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    cellCount = Selection.Rows.Count
    Set dateColumn = Selection
End If
    
'validate the dates
For Each cell In dateColumn
    cell.Select
    If cell.Value = "" Then
        GoTo Continue
        
    Else
        dateInCell = cell.Value
        cell.Value = dateInCell
    End If
Continue:
Next cell
End Function
