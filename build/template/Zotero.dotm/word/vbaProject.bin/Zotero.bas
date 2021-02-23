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
nl$ = Chr$(10)
wordVersion$ = "MacWord2016"
pipeLocation$ = "PIPE=""$HOME/.zoteroIntegrationPipe""; if [ ! -e ""$PIPE"" ]; then PIPE=""$HOME/Library/Containers/com.microsoft.Word/Data/.zoteroIntegrationPipe""; fi"
#If VBA6 Then
     Dim majorVersion As Integer
     majorVersion = Split(Application.Version, ".")(0)
     If majorVersion >= 16 Then
         wordVersion$ = "MacWord16"
     End If
#End If
shellScript$ = pipeLocation$ & "; if [ -e ""$PIPE"" ]; then echo '" & wordVersion$ & " " & func & " " & Application.Path & Application.PathSeparator & Application.Name & ".app/' > ""$PIPE""; else exit 1; fi;"
Result$ = AppleScriptTask("Zotero.scpt", "callZotero", shellScript$)
End Sub
