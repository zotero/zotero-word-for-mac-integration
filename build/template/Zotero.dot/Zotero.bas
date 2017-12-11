Attribute VB_Name = "Zotero"
Sub ZoteroInsertCitation()
CallZotero ("addCitation")
End Sub

Sub ZoteroEditCitation()
CallZotero ("editCitation")
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
MacScript "try" & nl$ & "do shell script """ & pipeLocation$ & "; if [ -e \""$PIPE\"" ]; then echo '" & wordVersion$ & " " & func & " "" & quoted form of POSIX path of (path to current application) & ""' > \""$PIPE\""; else exit 1; fi;""" & nl$ & "on error" & nl$ & "display alert ""Word could not communicate with Zotero. Please ensure that Zotero is open and try again.""  as critical" & nl$ & "end try"
End Sub
