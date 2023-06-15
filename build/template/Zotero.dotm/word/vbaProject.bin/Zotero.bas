Attribute VB_Name = "Zotero"
' ***** BEGIN LICENSE BLOCK *****
'
' Copyright (c) 2015  Zotero
'                     Center for History and New Media
'                     George Mason University, Fairfax, Virginia, USA
'                     http://zotero.org
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
' ***** END LICENSE BLOCK *****

Sub ZoteroInsertCitation()
CallZotero ("addCitation")
End Sub

Sub ZoteroAddNote()
CallZotero ("addNote")
End Sub

Sub ZoteroEditCitation()
CallZotero ("editCitation")
End Sub

Sub ZoteroAddEditCitation()
CallZotero ("addEditCitation")
End Sub

Sub ZoteroAddEditBibliography()
CallZotero ("addEditBibliography")
End Sub

Sub ZoteroSetDocPrefs()
CallZotero ("setDocPrefs")
End Sub

Sub ZoteroRefresh()
CallZotero ("refresh")
End Sub

Sub ZoteroRemoveCodes()
CallZotero ("removeCodes")
End Sub

Sub CallZotero(func)
Dim zoteroUrl As String
If zoteroUrl = "" Then
    zoteroUrl = "http://127.0.0.1:23119/integration/macWordCommand?"
    FileNum = FreeFile()
    On Error GoTo customUrlNotSet
        Open Application.StartupPath & Application.PathSeparator & "ZoteroUrl.txt" For Input As #FileNum
        If Not EOF(FileNum) Then
            Line Input #FileNum, DataLine
            zoteroUrl = DataLine & "/integration/macWordCommand?"
        End If
End If

customUrlNotSet:
nl$ = Chr$(10)
templateVersion$ = "2"
wordVersion$ = "MacWord2016"
onFailMessage$ = "Word could not communicate with Zotero. Please ensure that Zotero is open and try again."
wordAppPath$ = Replace(MacScript("POSIX path of (path to current application)"), " ", "%20")
#If VBA6 Then
     Dim majorVersion As Integer
     majorVersion = Split(Application.Version, ".")(0)
     If majorVersion >= 16 Then
         wordVersion$ = "MacWord16"
     End If
#End If
MacScript "try" & nl$ & "do shell script ""curl -s -o /dev/null -I -w '%{http_code}' -X GET '" & zoteroUrl & "agent=" & wordVersion$ & "&command=" & func & "&document=" & wordAppPath$ & "&templateVersion=" & templateVersion$ & "' | grep -q '200' || exit 1""" & nl$ & "on error" & nl$ & "display alert """ & onFailMessage$ & """  as critical" & nl$ & "end try"
End Sub
