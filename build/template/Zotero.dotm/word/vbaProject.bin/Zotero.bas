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

Sub ZoteroAddAnnotation()
CallZotero ("addAnnotation")
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
        Open Application.StartupPath & Application.PathSeparator & "ZoteroPort.txt" For Input As #FileNum
        If Not EOF(FileNum) Then
            Line Input #FileNum, DataLine
            zoteroUrl = "http://127.0.0.1:" & DataLine & "/integration/macWordCommand?"
        End If
End If

customUrlNotSet:
On Error GoTo -1
nl$ = Chr$(10)
templateVersion$ = "2"
wordVersion$ = "MacWord2016"
onFailMessage$ = "Word could not communicate with Zotero. Please ensure that Zotero is open and try again."

On Error GoTo catchWordPath
    wordAppPath$ = AppleScriptTask("Zotero.scpt", "getPath", "")
    GoTo okWordPath
catchWordPath:
    wordAppPath$ = MacScript("POSIX path of (path to current application)")
    
okWordPath:
On Error GoTo -1
wordAppPath$ = Replace(wordAppPath$, " ", "%20")
#If VBA6 Then
     Dim majorVersion As Integer
     majorVersion = Split(Application.Version, ".")(0)
     If majorVersion >= 16 Then
         wordVersion$ = "MacWord16"
     End If
#End If

httpShellScript$ = "curl -s -o /dev/null -I -m 2 -w '%{http_code}' -X GET '" & zoteroUrl & "agent=" & wordVersion$ & "&command=" & func & "&document=" & wordAppPath$ & "&templateVersion=" & templateVersion$ & "' | grep -q '200' || exit 1"

On Error GoTo catchMacScriptHttp
    MacScript "try" & nl$ & "do shell script """ & httpShellScript$ & """" & nl$ & "on error" & nl$ & "display alert """ & onFailMessage$ & """  as critical" & nl$ & "end try"
Exit Sub

catchMacScriptHttp:
On Error GoTo -1
On Error GoTo catchAppleScriptTaskHttp
    Result$ = AppleScriptTask("Zotero.scpt", "callZotero", httpShellScript$)
Exit Sub
    
catchAppleScriptTaskHttp:
On Error GoTo -1
On Error GoTo catchMacScriptPipeWrite
    pipeLocation$ = "PIPE=\""$CFFIXED_USER_HOME/.zoteroIntegrationPipe\""; if [ ! -e \""$PIPE\"" ]; then PIPE=\""/Users/$USER/Library/Containers/com.microsoft.Word/Data/.zoteroIntegrationPipe\""; fi; if [ ! -e \""$PIPE\"" ]; then PIPE=.zoteroIntegrationPipe; fi"
    MacScript "try" & nl$ & "do shell script """ & pipeLocation$ & "; if [ -e \""$PIPE\"" ]; then echo '" & wordVersion$ & " " & func & " "" & POSIX path of (path to current application) & "" " & templateVersion & "' > \""$PIPE\""; else exit 1; fi;""" & nl$ & "on error" & nl$ & "display alert """ & onFailMessage$ & """  as critical" & nl$ & "end try"
Exit Sub

catchMacScriptPipeWrite:
On Error GoTo -1
    pipeLocation$ = "PIPE=""$HOME/.zoteroIntegrationPipe""; if [ ! -e ""$PIPE"" ]; then PIPE=""$HOME/Library/Containers/com.microsoft.Word/Data/.zoteroIntegrationPipe""; fi"
    shellScript$ = pipeLocation$ & "; if [ -e ""$PIPE"" ]; then echo '" & wordVersion$ & " " & func & " " & Application.Path & Application.PathSeparator & Application.Name & ".app/ " & templateVersion & "' > ""$PIPE""; else exit 1; fi;"
    Result$ = AppleScriptTask("Zotero.scpt", "callZotero", shellScript$)
End Sub
