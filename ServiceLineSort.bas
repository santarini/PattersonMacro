Sub ServiceLineSort()

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

'The first sort should be by the four PTS Service Lines that currently exist in the database in the Service Line column
'(Readiness & Response, National Security, Logistics, IT/Cyber), not by the prefixes in the titles.
'We now want to use what is in the system without having to do that ‘title sort’ first.
'Then we will sort by the names in the “Dawson Capture Lead” column once that initial sort is done.

Dim cell, titleRng As Range
Dim sourceRng, ServiceLineCol As Range
Dim cellCount As Integer

'define the main source page
Set sourceSheet = Sheets("OpportunityDetails")

'Naviage to source sheet
sourceSheet.Activate

'copy header from source sheet
Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Copy
Range("A1").Select

'create and define service line tabs with header and define working ranges
'Readiness & Response
Sheets.Add.Name = "ReadyResp"
Set readyRespSheet = Sheets("ReadyResp")
readyRespSheet.Activate
readyRespSheet.Range("A1").Select
ActiveSheet.Paste
Set readyRespRng = readyRespSheet.Range("A2")
readyRespRng.Select

'National Security
Sheets.Add.Name = "NatSec"
Set natSecSheet = Sheets("NatSec")
natSecSheet.Activate
natSecSheet.Range("A1").Select
natSecSheet.Paste
Set natSecRng = natSecSheet.Range("A2")
natSecRng.Select

'Logistics
Sheets.Add.Name = "Logistics"
Set logisticsSheet = Sheets("Logistics")
logisticsSheet.Activate
logisticsSheet.Range("A1").Select
logisticsSheet.Paste
Set logisticsRng = logisticsSheet.Range("A2")
logisticsRng.Select

'IT/Cyber
Sheets.Add.Name = "IT_Cyber"
Set IT_CyberSheet = Sheets("IT_Cyber")
IT_CyberSheet.Activate
IT_CyberSheet.Range("A1").Select
IT_CyberSheet.Paste
Set IT_CyberRng = IT_CyberSheet.Range("A2")
IT_CyberRng.Select

'Naviage back to source sheet
sourceSheet.Activate

'find the cell with "Service Line"
Cells.Find(What:="Service Line", After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False).Select

'select all rows beneath it
Range(Selection, Selection.End(xlDown)).Offset(1, 0).Select
'cellCount = Selection.Rows.Count
Set ServiceLineCol = Selection

For Each cell In ServiceLineCol
   'sort for Readiness & Response
    If InStr(1, cell.Value, "Readiness & Response") > 0 Then
        sourceSheet.Activate
        cell.Select
        Selection.Offset(0, -13).Select
        ActiveCell.Resize(1, 70).Copy
        readyRespSheet.Activate
        readyRespRng.Select
        ActiveSheet.Paste
        readyRespRng.Offset(1, 0).Select
        Set readyRespRng = Selection
        i = i + 1
    End If
    
    'sort for National Security
    If InStr(1, cell.Value, "National Security") > 0 Then
        sourceSheet.Activate
        cell.Select
        Selection.Offset(0, -13).Select
        ActiveCell.Resize(1, 70).Copy
        natSecSheet.Activate
        natSecRng.Select
        ActiveSheet.Paste
        natSecRng.Offset(1, 0).Select
        Set natSecRng = Selection
        i = i + 1
     End If
    
    'sort for Logistics
    If InStr(1, cell.Value, "Logistics") > 0 Then
        sourceSheet.Activate
        cell.Select
        Selection.Offset(0, -13).Select
        ActiveCell.Resize(1, 70).Copy
        logisticsSheet.Activate
        logisticsRng.Select
        ActiveSheet.Paste
        logisticsRng.Offset(1, 0).Select
        Set logisticsRng = Selection
        i = i + 1
    End If
    
    'sort for IT/Cyber
    If InStr(1, cell.Value, "IT/Cyber") > 0 Then
        sourceSheet.Activate
        cell.Select
        Selection.Offset(0, -13).Select
        ActiveCell.Resize(1, 70).Copy
        IT_CyberSheet.Activate
        IT_CyberRng.Select
        ActiveSheet.Paste
        IT_CyberRng.Offset(1, 0).Select
        Set IT_CyberRng = Selection
        i = i + 1
    End If
Next cell

'clean up
'Readiness & Response
readyRespSheet.Activate
readyRespSheet.Range("A1").Select

'National Security
natSecSheet.Activate
natSecSheet.Range("A1").Select

'Logistics
logisticsSheet.Activate
logisticsSheet.Range("A1").Select

'IT/Cyber
IT_CyberSheet.Activate
IT_CyberSheet.Range("A1").Select

'source Sheet
sourceSheet.Activate
sourceSheet.Range("A1").Select



End Sub
