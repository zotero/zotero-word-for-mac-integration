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

Sub ZoteroEditBibliography()
CallZotero ("editBibliography")
End Sub

Sub ZoteroInsertBibliography()
CallZotero ("addBibliography")
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
pipeLocation$ = "PIPE=\""/Users/Shared/.zoteroIntegrationPipe_$LOGNAME\""; if [ ! -e \""$PIPE\"" ]; then PIPE=~/.zoteroIntegrationPipe; fi"
#If VBA6 Then
    Dim majorVersion As Integer
    majorVersion = Split(Application.Version, ".")(0)
    If majorVersion >= 15 Then
        wordVersion$ = "MacWord2016"
        pipeLocation$ = "PIPE=\""$CFFIXED_USER_HOME/.zoteroIntegrationPipe\""; if [ ! -e \""$PIPE\"" ]; then PIPE=\""/Users/$USER/Library/Containers/com.microsoft.Word/Data/.zoteroIntegrationPipe\""; fi; if [ ! -e \""$PIPE\"" ]; then PIPE=.zoteroIntegrationPipe; fi"
    Else
        wordVersion$ = "MacWord2008"
    End If
#Else
    wordVersion$ = "MacWord2004"
#End If
MacScript "try" & nl$ & "do shell script """ & pipeLocation$ & "; if [ -e \""$PIPE\"" ]; then echo '" & wordVersion$ & " " & func & " "" & quoted form of POSIX path of (path to current application) & ""' > \""$PIPE\""; else exit 1; fi;""" & nl$ & "on error" & nl$ & "display alert ""Word could not communicate with Zotero. Please ensure that Zotero or Firefox is open and try again.""  as critical" & nl$ & "end try"
End Sub
