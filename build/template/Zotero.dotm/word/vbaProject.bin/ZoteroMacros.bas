Attribute VB_Name = "ZoteroMacros"
' ***** BEGIN LICENSE BLOCK *****
'
' Copyright (c) 2026  Corporation for Digital Scholarship
'                     Vienna, Virginia, USA
'                     https://www.zotero.org
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

' Document and field operations of the Zotero plugin.
'
' Word handles one Apple Event per display frame, and most operations take many events with
' Scripting Bridge, so the plugin runs its document and field operations that need Word as these
' macros instead (see resource/macroDocument.mjs). Each Zotero command is a session: the
' plugin starts it with ZoteroBegin and ends it with ZoteroComplete. The session pins the
' document, window and pane that the command works on, remembers the Word settings that it
' changes so that they can be restored, and holds the handles of the fields that the plugin has
' seen. A handle holds a Range around its field, which Word keeps up to date as the document
' changes, so handles stay valid while other fields are added or removed.
'
' The plugin runs one macro, ZoteroCall, for every operation. Macros take no arguments, so the
' plugin writes the input to document variables ZOTERO_MACRO_IN_1..n of the target document and
' reads the output from ZOTERO_MACRO_OUT_1..n. The first variable of each starts with
' "<number of variables>;". Numbers are encoded as "<number>;" and strings as "<length>:<string>",
' with lengths in UTF-16 code units.
'
' Input: <operation><arguments>, where the operation is a string with the name of the function that
'   ZoteroCall runs
' Output: <status>;<result>
'   status 0: success, followed by the result
'   status 1: error, followed by <error number>;<description>
'
' ZoteroBegin checks the protocol version. Its input and output keep their format in all protocol
' versions, so that the plugin can tell which version a template supports.
'
' A field is returned as <handle>;<note index>;<adjacent to next field>;<code><text>, where the
' code is the Zotero code without the field code prefix.
'
' ZoteroCall handles all errors, since an unhandled error shows a dialog that blocks the Apple Event
' until the user dismisses it.

Option Explicit

Private Const PROTOCOL_VERSION As Long = 1
Private Const INPUT_PREFIX As String = "ZOTERO_MACRO_IN_"
Private Const OUTPUT_PREFIX As String = "ZOTERO_MACRO_OUT_"
' Document variables hold at most 65,280 characters
Private Const CHUNK_LENGTH As Long = 60000

Private Const FIELD_PREFIX As String = " ADDIN ZOTERO_"
Private Const OTHER_FIELD_PREFIX As String = " CSL_"
Private Const FIELD_PLACEHOLDER As String = "{Citation}"
Private Const PREFS_PROPERTY As String = "ZOTERO_PREF"
Private Const MAX_PROPERTY_LENGTH As Long = 255
Private Const EXPORT_MARKER As String = "ZOTERO_TRANSFER_DOCUMENT"
Private Const OLD_EXPORT_MARKER As String = "ZOTERO_EXPORTED_DOCUMENT"
Private Const IMPORT_LINK_URL As String = "https://www.zotero.org/"
Private Const IMPORT_ITEM_PREFIX As String = "ITEM CSL_CITATION "
Private Const IMPORT_BIBL_PREFIX As String = "BIBL "
Private Const IMPORT_DOC_PREFS_PREFIX As String = "DOCUMENT_PREFERENCES "
Private Const TEMP_BOOKMARK As String = "ZOTERO_TEMP_BOOKMARK"
' Precedes each field's part of the rich text for ZoteroSetFieldTexts, followed by its number and
' another SEGMENT_MARKER
Private Const SEGMENT_MARKER As String = "ZQZSEG"
' msoPropertyTypeString
Private Const PROPERTY_TYPE_STRING As Long = 4

' The error that these macros raise themselves. Its description is the message for the user.
Private Const ERR_ZOTERO As Long = vbObjectError + 1000

' The current call
Private callDoc As Document
Private callArgs As String
Private callPos As Long

' The current session, if sDoc is set
Private sDoc As Document
Private sWindow As Window
Private sPane As Pane
Private restoreFullScreen As Boolean
Private restoreFieldCodes As Boolean
Private restoreInsertions As Boolean
Private restoreFormatChanges As Boolean
Private restoreTrackRevisions As Boolean
' The note type of the story that the last field was inserted into, where insertText inserts notes
Private insertTextIntoNote As Long

' Field handles. Handle ids are indices.
Private handleRanges() As Range
' The field code of each handle's field. Unlike the range around the field, it doesn't grow when
' text or fields are inserted right before or after the field, so its start identifies the field.
Private handleAnchors() As Range
Private handleCount As Long

Private Type FieldRecord
        handle As Long
        noteType As Long
        noteIndex As Long
        textLocation As Long
        noteLocation As Long
        codeStart As Long
        resultEnd As Long
        zoteroCode As String
        resultText As String
End Type

' ---------------------------------------------------------------------------------------------
' Entry point
' ---------------------------------------------------------------------------------------------

' Reads the input, runs the operation and writes its output, or the error
Public Sub ZoteroCall()
        Dim operation As String
        Dim output As String
        On Error GoTo fail
        Set callDoc = Nothing
        Set callDoc = FindTargetDocument()
        If callDoc Is Nothing Then Exit Sub
        callArgs = ReadVariables(callDoc, INPUT_PREFIX)
        callPos = 1
        operation = ScanString(callArgs, callPos)
        If operation <> "ZoteroBegin" And Not callDoc Is sDoc Then
                Err.Raise ERR_ZOTERO, , "The Zotero session in Word has ended. Please try again."
        End If
        Select Case operation
                Case "ZoteroBegin"
                        output = ZoteroBegin()
                Case "ZoteroCleanup"
                        output = ZoteroCleanup()
                Case "ZoteroComplete"
                        output = ZoteroComplete()
                Case "ZoteroCanInsertField"
                        output = ZoteroCanInsertField()
                Case "ZoteroCursorInField"
                        output = ZoteroCursorInField()
                Case "ZoteroGetDocumentData"
                        output = ZoteroGetDocumentData()
                Case "ZoteroSetDocumentData"
                        output = ZoteroSetDocumentData()
                Case "ZoteroInsertField"
                        output = ZoteroInsertField()
                Case "ZoteroGetFields"
                        output = ZoteroGetFields()
                Case "ZoteroGetNoteIndices"
                        output = ZoteroGetNoteIndices()
                Case "ZoteroSetBibliographyStyle"
                        output = ZoteroSetBibliographyStyle()
                Case "ZoteroExportDocument"
                        output = ZoteroExportDocument()
                Case "ZoteroImportDocument"
                        output = ZoteroImportDocument()
                Case "ZoteroInsertText"
                        output = ZoteroInsertText()
                Case "ZoteroConvertPlaceholdersToFields"
                        output = ZoteroConvertPlaceholdersToFields()
                Case "ZoteroConvert"
                        output = ZoteroConvert()
                Case "ZoteroDeleteField"
                        output = ZoteroDeleteField()
                Case "ZoteroRemoveCode"
                        output = ZoteroRemoveCode()
                Case "ZoteroSelectField"
                        output = ZoteroSelectField()
                Case "ZoteroSetText"
                        output = ZoteroSetText()
                Case "ZoteroSetCode"
                        output = ZoteroSetCode()
                Case "ZoteroSetFieldTexts"
                        output = ZoteroSetFieldTexts()
                Case "ZoteroSetFieldCodes"
                        output = ZoteroSetFieldCodes()
                Case Else
                        Err.Raise ERR_ZOTERO, , "Unknown operation " & operation
        End Select
        WriteVariables callDoc, OUTPUT_PREFIX, "0;" & output
        Exit Sub
fail:
        WriteError Err.Number, Err.Description
End Sub

' ---------------------------------------------------------------------------------------------
' Session
' ---------------------------------------------------------------------------------------------

' Starts a session for the document if the plugin uses this protocol version: turns off change
' tracking, hides field codes and leaves full screen until the session completes. A session that
' didn't complete is discarded without restoring its settings.
' Input: <protocol version of the plugin>;
' Output: <protocol version of the template>;
Private Function ZoteroBegin() As String
        On Error GoTo fail
        DiscardSession
        If ScanNumber(callArgs, callPos) <> PROTOCOL_VERSION Then
                ZoteroBegin = PROTOCOL_VERSION & ";"
                Exit Function
        End If
        Set sDoc = callDoc
        Set sWindow = sDoc.ActiveWindow
        Set sPane = sWindow.ActivePane
        ReadSettings
        If restoreTrackRevisions Then sDoc.TrackRevisions = False
        If restoreFieldCodes Then sWindow.View.ShowFieldCodes = False
        If restoreFullScreen Then sWindow.View.FullScreen = False
        ZoteroBegin = PROTOCOL_VERSION & ";"
        Exit Function
fail:
        FailBegin
End Function

' Shows what was hidden to read field codes. Called while Zotero shows a dialog, after which the
' session continues.
' Input: none
' Output: none
Private Function ZoteroCleanup() As String
        SetRevisionsShown True
End Function

' Restores everything that the session changed and ends it. Restores as much as possible even if
' something fails, and then reports the first error.
' Input: none
' Output: none
Private Function ZoteroComplete() As String
        Dim errNumber As Long
        Dim errDescription As String
        On Error GoTo fail
        RestoreSettings errNumber, errDescription
        DiscardSession
        If errNumber <> 0 Then Err.Raise errNumber, , errDescription
        Exit Function
fail:
        errNumber = Err.Number
        errDescription = Err.Description
        DiscardSession
        Err.Raise errNumber, , errDescription
End Function

' ---------------------------------------------------------------------------------------------
' Document
' ---------------------------------------------------------------------------------------------

' Input: none
' Output: <1 if a field can be inserted at the cursor, 0 otherwise>;
Private Function ZoteroCanInsertField() As String
        Dim cursorStory As Long
        cursorStory = sPane.Selection.StoryType
        If cursorStory = wdMainTextStory Or cursorStory = wdFootnotesStory Or cursorStory = wdEndnotesStory Then
                ZoteroCanInsertField = "1;"
        Else
                ZoteroCanInsertField = "0;"
        End If
End Function

' Input: none
' Output: 1;<field> if the cursor is in a Zotero field, or 0; otherwise
Private Function ZoteroCursorInField() As String
        Dim sel As Selection
        Dim f As Field
        Dim selStart As Long
        Dim selEnd As Long
        Dim fieldStart As Long
        Dim fieldEnd As Long
        PrepareRead
        Set sel = sPane.Selection
        selStart = sel.Start
        selEnd = sel.End
        If sel.Fields.Count > 0 Then
                Set f = sel.Fields(1)
                If IsZoteroCode(f.Code.Text) Then
                        ZoteroCursorInField = "1;" & EncodeField(f, GetHandlesByPosition())
                Else
                        ZoteroCursorInField = "0;"
                End If
                Exit Function
        End If
        ' Look for a field around the cursor in its paragraph
        For Each f In sel.Paragraphs(1).Range.Fields
                fieldStart = f.Result.Start
                fieldEnd = f.Result.End
                If (selStart <= fieldStart And selEnd >= fieldStart) Or (selStart <= fieldEnd And selEnd >= fieldEnd) _
                                Or (selStart >= fieldStart And selStart <= fieldEnd) Then
                        If IsZoteroCode(f.Code.Text) Then
                                ZoteroCursorInField = "1;" & EncodeField(f, GetHandlesByPosition())
                                Exit Function
                        End If
                ElseIf fieldStart > selEnd Then
                        Exit For
                End If
        Next f
        ZoteroCursorInField = "0;"
End Function

' Input: none
' Output: <document data>
Private Function ZoteroGetDocumentData() As String
        ZoteroGetDocumentData = EncodeString(ReadDocumentData(sDoc))
End Function

' Input: <document data>
' Output: none
Private Function ZoteroSetDocumentData() As String
        WriteDocumentData sDoc, ScanString(callArgs, callPos)
End Function

' Inserts a field at the cursor, in a new note if noteType is not 0 and the cursor is in the main
' text, and puts the cursor after it.
' Input: <note type>;
' Output: <field>
Private Function ZoteroInsertField() As String
        Dim noteType As Long
        Dim where As Range
        Dim f As Field
        noteType = ScanNumber(callArgs, callPos)
        Set where = sPane.Selection.Range
        Select Case where.StoryType
                Case wdFootnotesStory
                        insertTextIntoNote = 1
                Case wdEndnotesStory
                        insertTextIntoNote = 2
        End Select
        Set f = InsertFieldAt(where, noteType)
        f.Result.Text = FIELD_PLACEHOLDER
        SelectRange CollapsedAfter(f)
        ZoteroInsertField = EncodeField(f, Nothing)
End Function

' Input: none
' Output: <number of fields>; and the fields in document order
Private Function ZoteroGetFields() As String
        Dim records() As FieldRecord
        Dim total As Long
        Dim parts() As String
        Dim i As Long
        PrepareRead
        total = CollectFields(records)
        If total = 0 Then
                ZoteroGetFields = "0;"
                Exit Function
        End If
        ReDim parts(0 To total - 1)
        For i = 0 To total - 1
                With records(i)
                        parts(i) = .handle & ";" & .noteIndex & ";" & IIf(IsAdjacent(records, total, i), 1, 0) & ";" _
                                & EncodeString(.zoteroCode) & EncodeString(.resultText)
                End With
        Next i
        ZoteroGetFields = total & ";" & Join(parts, "")
End Function

' Returns the note index of every field with a handle that still exists.
' Input: none
' Output: <number of fields>; then per field <handle>;<note index>;
Private Function ZoteroGetNoteIndices() As String
        Dim i As Long
        Dim fieldCount As Long
        Dim out As String
        Dim f As Field
        For i = 1 To handleCount
                If HandleExists(i) Then
                        Set f = ResolveField(i)
                        out = out & i & ";" & GetNoteIndex(f, GetNoteType(f)) & ";"
                        fieldCount = fieldCount + 1
                End If
        Next i
        ZoteroGetNoteIndices = fieldCount & ";" & out
End Function

' Sets up the "Bibliography" style. Lengths are in twips.
' Input: <first line indent>;<body indent>;<line spacing>;<entry spacing>;<number of tab stops>;
'   then the tab stops
Private Function ZoteroSetBibliographyStyle() As String
        Dim firstLineIndentTwips As Long
        Dim bodyIndentTwips As Long
        Dim lineSpacingTwips As Long
        Dim entrySpacingTwips As Long
        Dim tabCount As Long
        Dim i As Long
        Dim bibStyle As Style
        Dim styleName As String
        firstLineIndentTwips = ScanNumber(callArgs, callPos)
        bodyIndentTwips = ScanNumber(callArgs, callPos)
        lineSpacingTwips = ScanNumber(callArgs, callPos)
        entrySpacingTwips = ScanNumber(callArgs, callPos)
        tabCount = ScanNumber(callArgs, callPos)

        On Error Resume Next
        Set bibStyle = sDoc.Styles(wdStyleBibliography)
        styleName = bibStyle.NameLocal
        On Error GoTo 0
        If Len(styleName) = 0 Then
                Set bibStyle = sDoc.Styles.Add("Bibliography", wdStyleTypeParagraph)
                bibStyle.BaseStyle = wdStyleNormal
        End If
        With bibStyle.ParagraphFormat
                .TabStops.ClearAll
                .FirstLineIndent = firstLineIndentTwips / 20
                .LeftIndent = bodyIndentTwips / 20
                .LineSpacing = lineSpacingTwips / 20
                .SpaceAfter = entrySpacingTwips / 20
                For i = 1 To tabCount
                        .TabStops.Add Position:=Fix(ScanNumber(callArgs, callPos) / 20), _
                                Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
                Next i
        End With
End Function

' Replaces Zotero fields with links that other word processors keep, and adds the document data
' and import instructions.
' Input: <import instructions>
' Output: none
Private Function ZoteroExportDocument() As String
        Dim instructions As String
        Dim noteType As Long
        Dim flds As Fields
        Dim i As Long
        Dim f As Field
        Dim zoteroCode As String
        Dim where As Range
        Dim docData As String
        instructions = ScanString(callArgs, callPos)
        PrepareRead
        For noteType = 0 To 2
                If HasStory(noteType) Then
                        Set flds = GetStoryFields(noteType)
                        ' Last first, since replacing a field also removes any fields nested in it
                        For i = flds.Count To 1 Step -1
                                Set f = flds(i)
                                If IsZoteroCode(f.Code.Text) Then
                                        zoteroCode = ExtractCode(ReadFieldCode(f))
                                        Set where = f.Code.Duplicate
                                        where.SetRange f.Code.Start - 1, f.Code.Start - 1
                                        f.Delete
                                        InsertLink where, zoteroCode
                                End If
                        Next i
                End If
        Next noteType

        docData = ReadDocumentData(sDoc)
        sDoc.Content.InsertParagraphAfter
        Set where = sDoc.Content
        where.SetRange where.End - 1, where.End - 1
        InsertLink where, IMPORT_DOC_PREFS_PREFIX & docData

        sDoc.Range(0, 0).InsertBefore EXPORT_MARKER & vbCr & vbCr & instructions & vbCr & vbCr
        DiscardHandles
End Function

' Turns the links of an exported document back into fields and restores the document data.
' Input: none
' Output: <1 if document data was found, 0 otherwise>;
Private Function ZoteroImportDocument() As String
        Dim noteType As Long
        Dim flds As Fields
        Dim f As Field
        Dim i As Long
        Dim linkText As String
        Dim foundData As Boolean
        Dim r As Range
        For noteType = 0 To 2
                If HasStory(noteType) Then
                        Set flds = GetStoryFields(noteType)
                        For i = flds.Count To 1 Step -1
                                Set f = flds(i)
                                If f.Type = wdFieldHyperlink Then
                                        linkText = f.Result.Text
                                        If Left$(linkText, Len(IMPORT_ITEM_PREFIX)) = IMPORT_ITEM_PREFIX _
                                                        Or Left$(linkText, Len(IMPORT_BIBL_PREFIX)) = IMPORT_BIBL_PREFIX Then
                                                f.Code.Text = FIELD_PREFIX & linkText & " "
                                                f.Result.Style = wdStyleDefaultParagraphFont
                                                If noteType = 1 Then
                                                        f.Result.Style = wdStyleFootnoteText
                                                ElseIf noteType = 2 Then
                                                        f.Result.Style = wdStyleEndnoteText
                                                End If
                                        ElseIf Left$(linkText, Len(IMPORT_DOC_PREFS_PREFIX)) = IMPORT_DOC_PREFS_PREFIX Then
                                                foundData = True
                                                WriteDocumentData sDoc, Mid$(linkText, Len(IMPORT_DOC_PREFS_PREFIX) + 1)
                                                f.Delete
                                        End If
                                End If
                        Next i
                End If
        Next noteType

        ' Remove the export marker, the import instructions and the empty paragraphs after them
        Set r = sDoc.Range(0, 0)
        r.MoveEnd wdParagraph, 4
        r.Delete
        DiscardHandles
        ZoteroImportDocument = IIf(foundData, "1;", "0;")
End Function

' Pastes the HTML on the clipboard at the cursor, or into a new note if the last field was inserted
' into a note that has since been deleted, and puts the cursor after it.
' Input: none
' Output: none
Private Function ZoteroInsertText() As String
        Dim target As Range
        Dim pasteRange As Range
        Dim inserted As Range
        Dim r As Range
        Dim fontName As String
        Dim fontNameOther As String
        Dim i As Long
        ' If the last field was inserted into a note and deleting it deleted the note, the cursor is
        ' back in the main text and the text goes into a new note there
        Set target = sPane.Selection.Range
        If target.StoryType = wdMainTextStory Then
                If insertTextIntoNote = 1 Then
                        Set target = sDoc.Footnotes.Add(Range:=target).Range
                ElseIf insertTextIntoNote = 2 Then
                        Set target = sDoc.Endnotes.Add(Range:=target).Range
                End If
        End If

        ' Paste in place of an empty paragraph, since pasting HTML into a paragraph with text messes up
        ' the formatting of notes, and track the pasted text with a bookmark
        target.Text = vbCr & " "
        DeleteTempBookmark
        sDoc.Bookmarks.Add TEMP_BOOKMARK, target
        On Error Resume Next
        fontName = target.Font.Name
        fontNameOther = target.Font.NameOther
        On Error GoTo 0
        Set pasteRange = target.Duplicate
        pasteRange.SetRange target.Start, target.Start + 1
        pasteRange.Paste

        Set inserted = sDoc.Bookmarks(TEMP_BOOKMARK).Range
        On Error Resume Next
        inserted.Font.Name = fontName
        inserted.Font.NameOther = fontNameOther
        On Error GoTo 0
        ' Remove the space and the paragraph mark that the paste leaves at the end
        For i = 1 To 2
                Set inserted = sDoc.Bookmarks(TEMP_BOOKMARK).Range
                Set r = inserted.Duplicate
                r.SetRange inserted.End - 1, inserted.End
                r.Delete
        Next i

        Set inserted = sDoc.Bookmarks(TEMP_BOOKMARK).Range
        Set r = inserted.Duplicate
        r.Collapse wdCollapseEnd
        SelectRange r
        DeleteTempBookmark
End Function

' Turns the placeholder links that ZoteroInsertText pasted into fields. A link in the main text
' becomes a field in a new note if noteType is not 0.
' Input: <number of placeholders>; then the placeholder ids, then <note type>;
' Output: <number of fields>; and the fields, in the order of the placeholder ids
Private Function ZoteroConvertPlaceholdersToFields() As String
        Dim placeholderCount As Long
        Dim placeholderIds() As String
        Dim anchors() As Range
        Dim anchorStories() As Long
        Dim noteType As Long
        Dim storyIndex As Long
        Dim i As Long
        Dim f As Field
        Dim placeholderId As String
        Dim where As Range
        Dim fieldCount As Long
        Dim out As String
        placeholderCount = ScanNumber(callArgs, callPos)
        If placeholderCount = 0 Then
                ZoteroConvertPlaceholdersToFields = "0;"
                Exit Function
        End If
        ReDim placeholderIds(1 To placeholderCount)
        ReDim anchors(1 To placeholderCount)
        ReDim anchorStories(1 To placeholderCount)
        For i = 1 To placeholderCount
                placeholderIds(i) = ScanString(callArgs, callPos)
        Next i
        noteType = ScanNumber(callArgs, callPos)

        ' Find the links first, since converting them changes the field collections
        For storyIndex = 0 To 2
                If HasStory(storyIndex) Then
                        For Each f In GetStoryFields(storyIndex)
                                If f.Type = wdFieldHyperlink Then
                                        placeholderId = GetPlaceholderId(f.Code.Text)
                                        If Len(placeholderId) > 0 Then
                                                For i = 1 To placeholderCount
                                                        If Left$(placeholderId, Len(placeholderIds(i))) = placeholderIds(i) Then
                                                                Set anchors(i) = GetFieldRange(f)
                                                                anchorStories(i) = storyIndex
                                                                Exit For
                                                        End If
                                                Next i
                                        End If
                                End If
                        Next f
                End If
        Next storyIndex

        For i = 1 To placeholderCount
                If Not anchors(i) Is Nothing Then
                        Set f = FieldStartingAt(anchors(i))
                        If noteType = 0 Or anchorStories(i) <> 0 Then
                                f.Code.Text = FIELD_PREFIX & "TEMP "
                                f.Result.Font.ColorIndex = wdAuto
                                f.Result.Font.Underline = wdUnderlineNone
                        Else
                                Set where = f.Code.Duplicate
                                where.SetRange f.Code.Start - 1, f.Code.Start - 1
                                f.Delete
                                Set f = InsertFieldAt(where, noteType)
                                f.Result.Text = FIELD_PLACEHOLDER
                        End If
                        out = out & EncodeField(f, Nothing)
                        fieldCount = fieldCount + 1
                End If
        Next i
        ZoteroConvertPlaceholdersToFields = fieldCount & ";" & out
End Function

' Converts fields between in-text citations, footnotes and endnotes. A note that has text besides
' its field stays a note when converting to in-text citations. The handles of all fields are
' released, since Zotero reads the fields again afterwards.
' Input: <number of fields>; then per field <handle>;<note type to convert to>;
' Output: none
Private Function ZoteroConvert() As String
        Dim fieldCount As Long
        Dim handles() As Long
        Dim toNoteTypes() As Long
        Dim noteTypes() As Long
        Dim textLocations() As Long
        Dim i As Long
        Dim f As Field
        Dim newField As Field
        Dim noteType As Long
        Dim toNoteType As Long
        Dim refRange As Range
        Dim where As Range
        Dim noteText As String
        Dim zoteroCode As String
        Dim convertedLocation As Long
        fieldCount = ScanNumber(callArgs, callPos)
        If fieldCount > 0 Then
                ReDim handles(1 To fieldCount)
                ReDim toNoteTypes(1 To fieldCount)
                ReDim noteTypes(1 To fieldCount)
                ReDim textLocations(1 To fieldCount)
        End If
        For i = 1 To fieldCount
                handles(i) = ScanNumber(callArgs, callPos)
                toNoteTypes(i) = ScanNumber(callArgs, callPos)
        Next i
        PrepareRead

        ' Where each field or its note reference is, to tell which fields share a note
        For i = 1 To fieldCount
                Set f = ResolveField(handles(i))
                noteTypes(i) = GetNoteType(f)
                textLocations(i) = GetTextLocation(f, noteTypes(i))
        Next i

        convertedLocation = -1
        For i = 1 To fieldCount
                noteType = noteTypes(i)
                toNoteType = toNoteTypes(i)
                If noteType <> 0 And textLocations(i) = convertedLocation Then GoTo nextField
                Set f = ResolveField(handles(i))
                zoteroCode = ExtractCode(f.Code.Text)

                If noteType <> 0 And toNoteType <> 0 And noteType <> toNoteType Then
                        ' Convert the whole note, with all fields in it
                        If noteType = 1 Then
                                Set refRange = f.Result.Footnotes(1).Reference
                                sDoc.Range(refRange.Start, refRange.End).Footnotes.Convert
                        Else
                                Set refRange = f.Result.Endnotes(1).Reference
                                sDoc.Range(refRange.Start, refRange.End).Endnotes.Convert
                        End If
                        convertedLocation = textLocations(i)
                        GoTo nextField
                End If
                If noteType = toNoteType Then GoTo nextField

                If noteType = 0 Then
                        ' In-text citation to a note
                        Set where = f.Code.Duplicate
                        where.SetRange f.Code.Start - 1, f.Code.Start - 1
                        Set newField = InsertFieldAt(where, toNoteType)
                        newField.Result.Text = ResolveField(handles(i)).Result.Text
                        ResolveField(handles(i)).Delete
                Else
                        ' A note to an in-text citation, unless the note has other text
                        If noteType = 1 Then
                                noteText = f.Result.Footnotes(1).Range.Text
                                Set refRange = f.Result.Footnotes(1).Reference
                        Else
                                noteText = f.Result.Endnotes(1).Range.Text
                                Set refRange = f.Result.Endnotes(1).Reference
                        End If
                        If noteText <> f.Result.Text Then GoTo nextField
                        Set where = sDoc.Range(refRange.Start, refRange.Start)
                        Set newField = InsertFieldAt(where, 0)
                        newField.Result.Text = FIELD_PLACEHOLDER
                        If noteType = 1 Then
                                ResolveField(handles(i)).Result.Footnotes(1).Delete
                        Else
                                ResolveField(handles(i)).Result.Endnotes(1).Delete
                        End If
                End If
                newField.Code.Text = FIELD_PREFIX & zoteroCode & " "
nextField:
        Next i
        DiscardHandles
End Function

' ---------------------------------------------------------------------------------------------
' Fields
' ---------------------------------------------------------------------------------------------

' Deletes a field, and its note if the field is all the note has.
' Input: <handle>;
' Output: none
Private Function ZoteroDeleteField() As String
        Dim handle As Long
        Dim f As Field
        Dim noteRange As Range
        Dim refRange As Range
        Dim sel As Selection
        Dim noteType As Long
        handle = ScanNumber(callArgs, callPos)
        Set f = ResolveField(handle)
        noteType = GetNoteType(f)
        If noteType = 1 Then
                Set noteRange = f.Result.Footnotes(1).Range
                Set refRange = f.Result.Footnotes(1).Reference
        ElseIf noteType = 2 Then
                Set noteRange = f.Result.Endnotes(1).Range
                Set refRange = f.Result.Endnotes(1).Reference
        End If
        If Not noteRange Is Nothing Then
                If noteRange.Start = f.Code.Start - 1 And noteRange.End = f.Result.End + 1 Then
                        ' Move the cursor out of the note before deleting it
                        Set sel = sPane.Selection
                        If sel.StoryType = noteRange.StoryType And sel.Start >= noteRange.Start _
                                        And sel.End <= noteRange.End Then
                                SelectRange sDoc.Range(refRange.End, refRange.End)
                        End If
                        If noteType = 1 Then
                                f.Result.Footnotes(1).Delete
                        Else
                                f.Result.Endnotes(1).Delete
                        End If
                        ReleaseHandle handle
                        Exit Function
                End If
        End If
        f.Delete
        ReleaseHandle handle
End Function

' Turns a field into plain text.
' Input: <handle>;
' Output: none
Private Function ZoteroRemoveCode() As String
        Dim handle As Long
        handle = ScanNumber(callArgs, callPos)
        ResolveField(handle).Unlink
        ReleaseHandle handle
End Function

' Selects the text of a field.
' Input: <handle>;
' Output: none
Private Function ZoteroSelectField() As String
        SelectRange ResolveField(ScanNumber(callArgs, callPos)).Result
End Function

' Sets the text of a field. Rich text is pasted from the clipboard.
' Input: <handle>;<is rich>;<is bibliography>;<text>
' Output: <resulting text>
Private Function ZoteroSetText() As String
        Dim handle As Long
        Dim isRich As Boolean
        Dim isBibliography As Boolean
        Dim newText As String
        Dim f As Field
        Dim previousScreenUpdating As Boolean
        Dim screenUpdatingChanged As Boolean
        Dim errNumber As Long
        Dim errDescription As String
        On Error GoTo fail
        handle = ScanNumber(callArgs, callPos)
        isRich = ScanNumber(callArgs, callPos) <> 0
        isBibliography = ScanNumber(callArgs, callPos) <> 0
        newText = ScanString(callArgs, callPos)

        ' Don't redraw the document for each change
        previousScreenUpdating = Application.ScreenUpdating
        Application.ScreenUpdating = False
        screenUpdatingChanged = True
        Set f = ResolveField(handle)
        If isRich Then
                SetRichText f, isBibliography, Nothing
        Else
                SetPlainText f, newText
        End If
        Set f = ResolveField(handle)
        Application.ScreenUpdating = previousScreenUpdating
        ZoteroSetText = EncodeString(f.Result.Text)
        Exit Function
fail:
        errNumber = Err.Number
        errDescription = Err.Description
        If screenUpdatingChanged Then SetScreenUpdating previousScreenUpdating
        Err.Raise errNumber, , errDescription
End Function

Private Sub SetScreenUpdating(updating As Boolean)
        On Error Resume Next
        Application.ScreenUpdating = updating
End Sub

' Sets the code of a field.
' Input: <handle>;<Zotero code>
' Output: none
Private Function ZoteroSetCode() As String
        Dim handle As Long
        handle = ScanNumber(callArgs, callPos)
        ResolveField(handle).Code.Text = FIELD_PREFIX & ScanString(callArgs, callPos) & " "
        ' Points the handle at the new code
        ResolveField handle
End Function

' Sets the texts of several fields. The rich text is on the clipboard as one RTF document, with each
' rich field's text after SEGMENT_MARKER & <number> & SEGMENT_MARKER, numbering the rich fields from
' 0 in input order, and SEGMENT_MARKER & "END" & SEGMENT_MARKER at the end. It is pasted once into a
' hidden document, from which each field's part is copied.
' Input: <number of fields>; then per field <handle>;<is rich>;<is bibliography>;<text>, where the
'   text of a rich field is empty
' Output: <number of fields>; then the resulting text of each field
Private Function ZoteroSetFieldTexts() As String
        Dim fieldCount As Long
        Dim handles() As Long
        Dim isRich() As Boolean
        Dim isBibliography() As Boolean
        Dim texts() As String
        Dim results() As String
        Dim richCount As Long
        Dim segment As Long
        Dim segmentStarts() As Long
        Dim segmentEnds() As Long
        Dim tempDoc As Document
        Dim f As Field
        Dim i As Long
        Dim previousScreenUpdating As Boolean
        Dim screenUpdatingChanged As Boolean
        Dim errNumber As Long
        Dim errDescription As String
        On Error GoTo fail
        fieldCount = ScanNumber(callArgs, callPos)
        If fieldCount = 0 Then
                ZoteroSetFieldTexts = "0;"
                Exit Function
        End If
        ReDim handles(1 To fieldCount)
        ReDim isRich(1 To fieldCount)
        ReDim isBibliography(1 To fieldCount)
        ReDim texts(1 To fieldCount)
        ReDim results(1 To fieldCount)
        For i = 1 To fieldCount
                handles(i) = ScanNumber(callArgs, callPos)
                isRich(i) = ScanNumber(callArgs, callPos) <> 0
                isBibliography(i) = ScanNumber(callArgs, callPos) <> 0
                texts(i) = ScanString(callArgs, callPos)
                If isRich(i) Then richCount = richCount + 1
        Next i

        ' Don't redraw the document for each change
        previousScreenUpdating = Application.ScreenUpdating
        Application.ScreenUpdating = False
        screenUpdatingChanged = True
        If richCount > 0 Then
                Set tempDoc = Documents.Add(Visible:=False)
                PasteSegments tempDoc, richCount, segmentStarts, segmentEnds
        End If
        For i = 1 To fieldCount
                Set f = ResolveField(handles(i))
                If isRich(i) Then
                        SetRichText f, isBibliography(i), tempDoc.Range(segmentStarts(segment), segmentEnds(segment))
                        segment = segment + 1
                Else
                        SetPlainText f, texts(i)
                End If
                results(i) = EncodeString(ResolveField(handles(i)).Result.Text)
        Next i
        DiscardDocument tempDoc
        SetScreenUpdating previousScreenUpdating
        ZoteroSetFieldTexts = fieldCount & ";" & Join(results, "")
        Exit Function
fail:
        errNumber = Err.Number
        errDescription = Err.Description
        DiscardDocument tempDoc
        If screenUpdatingChanged Then SetScreenUpdating previousScreenUpdating
        Err.Raise errNumber, , errDescription
End Function

' Sets the codes of several fields.
' Input: <number of fields>; then per field <handle>;<Zotero code>
' Output: none
Private Function ZoteroSetFieldCodes() As String
        Dim fieldCount As Long
        Dim handles() As Long
        Dim codes() As String
        Dim i As Long
        fieldCount = ScanNumber(callArgs, callPos)
        If fieldCount = 0 Then Exit Function
        ReDim handles(1 To fieldCount)
        ReDim codes(1 To fieldCount)
        For i = 1 To fieldCount
                handles(i) = ScanNumber(callArgs, callPos)
                codes(i) = ScanString(callArgs, callPos)
        Next i
        For i = 1 To fieldCount
                ResolveField(handles(i)).Code.Text = FIELD_PREFIX & codes(i) & " "
                ' Points the handle at the new code
                ResolveField handles(i)
        Next i
End Function

' ---------------------------------------------------------------------------------------------
' Session state
' ---------------------------------------------------------------------------------------------

' Reads the settings that the session may change. A setting that can't be read is left alone.
Private Sub ReadSettings()
        On Error Resume Next
        restoreFullScreen = sWindow.View.FullScreen
        restoreFieldCodes = sWindow.View.ShowFieldCodes
        restoreInsertions = sWindow.View.ShowInsertionsAndDeletions
        restoreFormatChanges = sWindow.View.ShowFormatChanges
        restoreTrackRevisions = sDoc.TrackRevisions
End Sub

' Undoes what ZoteroBegin did before it failed, and raises the error again
Private Sub FailBegin()
        Dim errNumber As Long
        Dim errDescription As String
        Dim restoreErrNumber As Long
        Dim restoreErrDescription As String
        errNumber = Err.Number
        errDescription = Err.Description
        RestoreSettings restoreErrNumber, restoreErrDescription
        DiscardSession
        Err.Raise errNumber, , errDescription
End Sub

' Restores the settings that the current session changed, as far as possible, and sets errNumber
' and errDescription to the first error
Private Sub RestoreSettings(errNumber As Long, errDescription As String)
        On Error Resume Next
        If sDoc Is Nothing Then Exit Sub
        SetRevisionsShown True
        NoteFirstError errNumber, errDescription
        If restoreTrackRevisions Then sDoc.TrackRevisions = True
        NoteFirstError errNumber, errDescription
        If restoreFieldCodes Then sWindow.View.ShowFieldCodes = True
        NoteFirstError errNumber, errDescription
        If restoreFullScreen Then sWindow.View.FullScreen = True
        NoteFirstError errNumber, errDescription
End Sub

Private Sub DiscardSession()
        Set sDoc = Nothing
        Set sWindow = Nothing
        Set sPane = Nothing
        restoreFullScreen = False
        restoreFieldCodes = False
        restoreInsertions = False
        restoreFormatChanges = False
        restoreTrackRevisions = False
        insertTextIntoNote = 0
        DiscardHandles
End Sub

Private Sub NoteFirstError(errNumber As Long, errDescription As String)
        If Err.Number <> 0 And errNumber = 0 Then
                errNumber = Err.Number
                errDescription = Err.Description
        End If
        Err.Clear
End Sub

' Hides tracked changes so that field codes and text read as they are without them
Private Sub PrepareRead()
        SetRevisionsShown False
End Sub

' Hides the tracked changes that the user shows, or shows them again. Changing these settings fails
' while the cursor is in a note, so the cursor is moved to the note's reference and back. Puts the
' cursor back and changes both settings even if one fails, and then raises the first error.
Private Sub SetRevisionsShown(shown As Boolean)
        Dim setInsertions As Boolean
        Dim setFormatChanges As Boolean
        Dim noteCursor As Range
        Dim errNumber As Long
        Dim errDescription As String
        If restoreInsertions Then setInsertions = sWindow.View.ShowInsertionsAndDeletions <> shown
        If restoreFormatChanges Then setFormatChanges = sWindow.View.ShowFormatChanges <> shown
        If Not setInsertions And Not setFormatChanges Then Exit Sub
        Set noteCursor = LeaveNote()
        On Error Resume Next
        If setInsertions Then
                sWindow.View.ShowInsertionsAndDeletions = shown
                NoteFirstError errNumber, errDescription
        End If
        If setFormatChanges Then
                sWindow.View.ShowFormatChanges = shown
                NoteFirstError errNumber, errDescription
        End If
        If Not noteCursor Is Nothing Then SelectRange noteCursor
        NoteFirstError errNumber, errDescription
        On Error GoTo 0
        If errNumber <> 0 Then Err.Raise errNumber, , errDescription
End Sub

' If the cursor is in a note, moves it to the note's reference and returns where it was, so that it
' can be put back. Tracked changes can't be hidden or shown while the cursor is in a note.
Private Function LeaveNote() As Range
        Dim sel As Selection
        Set sel = sPane.Selection
        Select Case sel.StoryType
                Case wdFootnotesStory
                        Set LeaveNote = sel.Range
                        SelectRange sel.Footnotes(1).Reference
                Case wdEndnotesStory
                        Set LeaveNote = sel.Range
                        SelectRange sel.Endnotes(1).Reference
        End Select
End Function

' Selects a range, in the session's pane if the selection there is in the same story
Private Sub SelectRange(r As Range)
        If sPane.Selection.StoryType = r.StoryType Then
                sPane.Selection.SetRange r.Start, r.End
        Else
                r.Select
        End If
End Sub

' ---------------------------------------------------------------------------------------------
' Fields and handles
' ---------------------------------------------------------------------------------------------

Private Function IsZoteroCode(fieldCode As String) As Boolean
        IsZoteroCode = InStr(fieldCode, FIELD_PREFIX) > 0 Or InStr(fieldCode, OTHER_FIELD_PREFIX) > 0
End Function

' The Zotero code within a field code
Private Function ExtractCode(fieldCode As String) As String
        Dim prefixStart As Long
        Dim codeStart As Long
        Dim zoteroCode As String
        prefixStart = InStr(fieldCode, FIELD_PREFIX)
        If prefixStart > 0 Then
                codeStart = prefixStart + Len(FIELD_PREFIX)
        Else
                prefixStart = InStr(fieldCode, OTHER_FIELD_PREFIX)
                If prefixStart = 0 Then Exit Function
                codeStart = prefixStart + Len(OTHER_FIELD_PREFIX)
        End If
        zoteroCode = Mid$(fieldCode, codeStart)
        ' Sometimes there is a space at the end of the code
        If Right$(zoteroCode, 1) = " " Then zoteroCode = Left$(zoteroCode, Len(zoteroCode) - 1)
        ExtractCode = zoteroCode
End Function

' The field code, with all caps turned off if the code is only displayed in all caps
Private Function ReadFieldCode(f As Field) As String
        Dim fieldCode As String
        fieldCode = f.Code.Text
        If fieldCode = UCase$(fieldCode) Then
                If f.Code.Font.AllCaps <> 0 Then
                        f.Code.Font.AllCaps = False
                        fieldCode = f.Code.Text
                End If
        End If
        ReadFieldCode = fieldCode
End Function

Private Function HasStory(noteType As Long) As Boolean
        If noteType = 1 Then
                HasStory = sDoc.Footnotes.Count > 0
        ElseIf noteType = 2 Then
                HasStory = sDoc.Endnotes.Count > 0
        Else
                HasStory = True
        End If
End Function

Private Function GetStoryFields(noteType As Long) As Fields
        If noteType = 1 Then
                Set GetStoryFields = sDoc.StoryRanges(wdFootnotesStory).Fields
        ElseIf noteType = 2 Then
                Set GetStoryFields = sDoc.StoryRanges(wdEndnotesStory).Fields
        Else
                Set GetStoryFields = sDoc.Fields
        End If
End Function

Private Function GetNoteType(f As Field) As Long
        Select Case f.Result.StoryType
                Case wdFootnotesStory
                        GetNoteType = 1
                Case wdEndnotesStory
                        GetNoteType = 2
                Case Else
                        GetNoteType = 0
        End Select
End Function

Private Function GetNoteIndex(f As Field, noteType As Long) As Long
        If noteType = 1 Then
                GetNoteIndex = f.Result.Footnotes(1).Index
        ElseIf noteType = 2 Then
                GetNoteIndex = f.Result.Endnotes(1).Index
        End If
End Function

' Where a field, or the reference of its note, is in the main text
Private Function GetTextLocation(f As Field, noteType As Long) As Long
        If noteType = 1 Then
                GetTextLocation = f.Result.Footnotes(1).Reference.Start
        ElseIf noteType = 2 Then
                GetTextLocation = f.Result.Endnotes(1).Reference.Start
        Else
                GetTextLocation = f.Result.Start
        End If
End Function

' A collapsed range after the end of a field
Private Function CollapsedAfter(f As Field) As Range
        Dim r As Range
        Set r = f.Result.Duplicate
        r.SetRange f.Result.End + 1, f.Result.End + 1
        Set CollapsedAfter = r
End Function

' Inserts a field. If noteType is not 0 and where is in the main text, the field goes into a new
' note there.
Private Function InsertFieldAt(where As Range, noteType As Long) As Field
        Dim r As Range
        Dim f As Field
        Set r = where.Duplicate
        If noteType <> 0 And r.StoryType = wdMainTextStory Then
                If noteType = 1 Then
                        Set r = sDoc.Footnotes.Add(Range:=r).Range
                Else
                        Set r = sDoc.Endnotes.Add(Range:=r).Range
                End If
                r.Text = ""
        End If
        ' A print date field has a result that can be set, unlike a new ADDIN field
        Set f = r.Fields.Add(Range:=r, Type:=wdFieldPrintDate, PreserveFormatting:=False)
        f.Code.Text = FIELD_PREFIX & "TEMP "
        Set InsertFieldAt = f
End Function

' Inserts a link to IMPORT_LINK_URL with the given text, which other word processors keep
Private Sub InsertLink(where As Range, linkText As String)
        Dim f As Field
        Set f = where.Fields.Add(Range:=where, Type:=wdFieldHyperlink, _
                Text:="""" & IMPORT_LINK_URL & """", PreserveFormatting:=False)
        f.Result.Text = linkText
End Sub

' The placeholder id in the code of a link to IMPORT_LINK_URL + "?" + id, or ""
Private Function GetPlaceholderId(fieldCode As String) As String
        Dim urlStart As Long
        Dim rest As String
        Dim quoteAt As Long
        urlStart = InStr(fieldCode, IMPORT_LINK_URL)
        If urlStart = 0 Then Exit Function
        rest = Mid$(fieldCode, urlStart + Len(IMPORT_LINK_URL) + 1)
        quoteAt = InStr(rest, """")
        If quoteAt > 0 Then rest = Left$(rest, quoteAt - 1)
        GetPlaceholderId = Trim$(rest)
End Function

' A field with a new or existing handle, encoded for output. New fields pass Nothing for
' existingHandles.
Private Function EncodeField(f As Field, existingHandles As Collection) As String
        Dim fieldCode As String
        fieldCode = ReadFieldCode(f)
        EncodeField = AddHandle(f, existingHandles) & ";" & GetNoteIndex(f, GetNoteType(f)) & ";0;" _
                & EncodeString(ExtractCode(fieldCode)) & EncodeString(f.Result.Text)
End Function

' Reads the Zotero fields of the document in order, creating handles for them
Private Function CollectFields(records() As FieldRecord) As Long
        Dim mainRecords() As FieldRecord
        Dim footnoteRecords() As FieldRecord
        Dim endnoteRecords() As FieldRecord
        Dim counts(0 To 2) As Long
        Dim nextIndex(0 To 2) As Long
        Dim existingHandles As Collection
        Dim total As Long
        Dim pick As Long
        Dim best As FieldRecord
        Dim i As Long

        Set existingHandles = GetHandlesByPosition()
        counts(0) = CollectStoryFields(0, mainRecords, existingHandles)
        If HasStory(1) Then counts(1) = CollectStoryFields(1, footnoteRecords, existingHandles)
        If HasStory(2) Then counts(2) = CollectStoryFields(2, endnoteRecords, existingHandles)
        total = counts(0) + counts(1) + counts(2)

        ' Merge the stories by the position of the field or its note reference
        If total > 0 Then ReDim records(0 To total - 1)
        For i = 0 To total - 1
                pick = -1
                If nextIndex(0) < counts(0) Then ConsiderRecord mainRecords(nextIndex(0)), 0, pick, best
                If nextIndex(1) < counts(1) Then ConsiderRecord footnoteRecords(nextIndex(1)), 1, pick, best
                If nextIndex(2) < counts(2) Then ConsiderRecord endnoteRecords(nextIndex(2)), 2, pick, best
                records(i) = best
                nextIndex(pick) = nextIndex(pick) + 1
        Next i
        CollectFields = total
End Function

' For the merge in CollectFields: makes a story's next record the best one if it comes first. On a
' tie, the story considered first wins.
Private Sub ConsiderRecord(candidate As FieldRecord, story As Long, pick As Long, best As FieldRecord)
        If pick <> -1 Then
                If CompareRecords(candidate, best) >= 0 Then Exit Sub
        End If
        pick = story
        best = candidate
End Sub

Private Function CollectStoryFields(noteType As Long, records() As FieldRecord, _
                existingHandles As Collection) As Long
        Dim f As Field
        Dim recordCount As Long
        Dim fieldCode As String
        ReDim records(0 To 63)
        For Each f In GetStoryFields(noteType)
                fieldCode = f.Code.Text
                If IsZoteroCode(fieldCode) Then
                        fieldCode = ReadFieldCode(f)
                        If recordCount > UBound(records) Then ReDim Preserve records(0 To 2 * recordCount - 1)
                        With records(recordCount)
                                .handle = AddHandle(f, existingHandles)
                                .noteType = noteType
                                .zoteroCode = ExtractCode(fieldCode)
                                .resultText = f.Result.Text
                                .codeStart = f.Code.Start
                                .resultEnd = f.Result.End
                                .noteLocation = f.Result.Start
                                .textLocation = GetTextLocation(f, noteType)
                                .noteIndex = GetNoteIndex(f, noteType)
                        End With
                        recordCount = recordCount + 1
                End If
        Next f
        CollectStoryFields = recordCount
End Function

Private Function CompareRecords(a As FieldRecord, b As FieldRecord) As Long
        If a.textLocation <> b.textLocation Then
                CompareRecords = IIf(a.textLocation < b.textLocation, -1, 1)
        ElseIf a.noteType <> 0 And b.noteType <> 0 And a.noteLocation <> b.noteLocation Then
                CompareRecords = IIf(a.noteLocation < b.noteLocation, -1, 1)
        End If
End Function

' Fields count as adjacent when at most one character separates them
Private Function IsAdjacent(records() As FieldRecord, total As Long, i As Long) As Boolean
        If i + 1 >= total Then Exit Function
        IsAdjacent = records(i).noteType = records(i + 1).noteType _
                And records(i).resultEnd + 2 - records(i + 1).codeStart >= 0
End Function

Private Function HandleExists(handle As Long) As Boolean
        If handle < 1 Or handle > handleCount Then Exit Function
        If handleRanges(handle) Is Nothing Then Exit Function
        If handleRanges(handle).End > handleRanges(handle).Start Then
                HandleExists = True
        Else
                ' The field has been deleted
                ReleaseHandle handle
        End If
End Function

' Live handles, keyed by story and start of their field code
Private Function GetHandlesByPosition() As Collection
        Dim handles As New Collection
        Dim i As Long
        For i = 1 To handleCount
                If HandleExists(i) Then AddToCollection handles, i, PositionKey(handleAnchors(i))
        Next i
        Set GetHandlesByPosition = handles
End Function

Private Sub AddToCollection(c As Collection, entryValue As Variant, itemKey As String)
        On Error Resume Next
        c.Add entryValue, itemKey
End Sub

Private Function GetFromCollection(c As Collection, itemKey As String) As Long
        On Error Resume Next
        GetFromCollection = c(itemKey)
End Function

' The key of a field code's position in GetHandlesByPosition
Private Function PositionKey(fieldCode As Range) As String
        PositionKey = fieldCode.StoryType & ":" & fieldCode.Start
End Function

' Returns the handle of a field, reusing the handle that already refers to it. New fields pass
' Nothing for existingHandles.
Private Function AddHandle(f As Field, existingHandles As Collection) As Long
        Dim existing As Long
        If Not existingHandles Is Nothing Then
                existing = GetFromCollection(existingHandles, PositionKey(f.Code))
                If existing > 0 Then
                        SetHandleField existing, f
                        AddHandle = existing
                        Exit Function
                End If
        End If
        If handleCount = 0 Then
                ReDim handleRanges(1 To 256)
                ReDim handleAnchors(1 To 256)
        ElseIf handleCount >= UBound(handleRanges) Then
                ReDim Preserve handleRanges(1 To 2 * handleCount)
                ReDim Preserve handleAnchors(1 To 2 * handleCount)
        End If
        handleCount = handleCount + 1
        SetHandleField handleCount, f
        AddHandle = handleCount
End Function

' The range of a field, from its start character to its end character
Private Function GetFieldRange(f As Field) As Range
        Dim r As Range
        Set r = f.Code.Duplicate
        r.SetRange f.Code.Start - 1, f.Result.End + 1
        Set GetFieldRange = r
End Function

' The field that starts where a range does
Private Function FieldStartingAt(r As Range) As Field
        Dim f As Field
        For Each f In r.Fields
                If f.Code.Start - 1 = r.Start Then
                        Set FieldStartingAt = f
                        Exit Function
                End If
        Next f
        Err.Raise ERR_ZOTERO, , "Field not found"
End Function

' Points a handle at a field
Private Sub SetHandleField(handle As Long, f As Field)
        Set handleRanges(handle) = GetFieldRange(f)
        Set handleAnchors(handle) = f.Code.Duplicate
End Sub

Private Sub ReleaseHandle(handle As Long)
        Set handleRanges(handle) = Nothing
        Set handleAnchors(handle) = Nothing
End Sub

Private Sub DiscardHandles()
        If handleCount > 0 Then
                Erase handleRanges
                Erase handleAnchors
                handleCount = 0
        End If
End Sub

' Returns the field of a handle: the field whose code starts where the handle's anchor does
Private Function ResolveField(handle As Long) As Field
        Dim f As Field
        If Not HandleExists(handle) Then Err.Raise ERR_ZOTERO, , "Field has been deleted"
        For Each f In handleRanges(handle).Fields
                If f.Code.Start = handleAnchors(handle).Start Then
                        SetHandleField handle, f
                        Set ResolveField = f
                        Exit Function
                End If
        Next f
        Err.Raise ERR_ZOTERO, , "Field has been deleted"
End Function

Private Sub SetPlainText(f As Field, newText As String)
        ' Superscripting is more likely to come from a style change than to be intended
        On Error Resume Next
        f.Result.Font.Superscript = False
        On Error GoTo 0
        f.Result.Text = newText
End Sub

' Sets the text of a field to rich text, keeping the font of the existing text: pastes the RTF on the
' clipboard if source is Nothing, and copies source otherwise
Private Sub SetRichText(f As Field, isBibliography As Boolean, source As Range)
        Dim fontSize As Single
        Dim fontName As String
        Dim fontNameOther As String
        Dim fontColorIndex As Long
        Dim adjustWordSpacing As Boolean
        Dim resultRange As Range

        On Error Resume Next
        DeleteTempBookmark
        With f.Result.Font
                fontSize = .Size
                fontName = .Name
                fontNameOther = .NameOther
                fontColorIndex = .ColorIndex
        End With
        adjustWordSpacing = Options.PasteAdjustWordSpacing
        Options.PasteAdjustWordSpacing = False
        On Error GoTo restore

        Set resultRange = f.Result
        If source Is Nothing Then
                resultRange.Paste
        Else
                resultRange.FormattedText = source.FormattedText
        End If

        On Error Resume Next
        Options.PasteAdjustWordSpacing = adjustWordSpacing
        If isBibliography Then f.Result.Style = wdStyleBibliography
        With f.Result.Font
                .Size = fontSize
                .Name = fontName
                .NameOther = fontNameOther
                .ColorIndex = fontColorIndex
        End With
        Exit Sub
restore:
        Dim errNumber As Long, errDescription As String
        errNumber = Err.Number
        errDescription = Err.Description
        On Error Resume Next
        Options.PasteAdjustWordSpacing = adjustWordSpacing
        On Error GoTo 0
        Err.Raise errNumber, , errDescription
End Sub

' Pastes the RTF document on the clipboard into tempDoc for ZoteroSetFieldTexts, and sets where the
' text of each rich field starts and ends in it
Private Sub PasteSegments(tempDoc As Document, richCount As Long, segmentStarts() As Long, _
                segmentEnds() As Long)
        Dim adjustWordSpacing As Boolean
        Dim pastedText As String
        Dim marker As String
        Dim markerPos As Long
        Dim k As Long
        Dim errNumber As Long
        Dim errDescription As String

        adjustWordSpacing = Options.PasteAdjustWordSpacing
        Options.PasteAdjustWordSpacing = False
        On Error Resume Next
        tempDoc.Content.Paste
        errNumber = Err.Number
        errDescription = Err.Description
        Options.PasteAdjustWordSpacing = adjustWordSpacing
        On Error GoTo 0
        If errNumber <> 0 Then Err.Raise errNumber, , errDescription

        ' Positions in the text are positions in the document only if it is all text
        pastedText = tempDoc.Content.Text
        If Len(pastedText) <> tempDoc.Content.End - tempDoc.Content.Start Then
                Err.Raise ERR_ZOTERO, , "The rich text has content that can't be split into fields"
        End If
        ReDim segmentStarts(0 To richCount - 1)
        ReDim segmentEnds(0 To richCount - 1)
        markerPos = 1
        For k = 0 To richCount
                If k = richCount Then
                        marker = SEGMENT_MARKER & "END" & SEGMENT_MARKER
                Else
                        marker = SEGMENT_MARKER & k & SEGMENT_MARKER
                End If
                markerPos = InStr(markerPos, pastedText, marker)
                If markerPos = 0 Then Err.Raise ERR_ZOTERO, , "The rich text is missing a segment marker"
                ' Positions in the text start at 1, and positions in the document at 0
                If k > 0 Then segmentEnds(k - 1) = tempDoc.Content.Start + markerPos - 1
                If k < richCount Then segmentStarts(k) = tempDoc.Content.Start + markerPos - 1 + Len(marker)
                markerPos = markerPos + Len(marker)
        Next k
End Sub

' Closes a document without saving it, if there is one
Private Sub DiscardDocument(d As Document)
        On Error Resume Next
        If Not d Is Nothing Then d.Close SaveChanges:=wdDoNotSaveChanges
End Sub

Private Sub DeleteTempBookmark()
        If sDoc.Bookmarks.Exists(TEMP_BOOKMARK) Then sDoc.Bookmarks(TEMP_BOOKMARK).Delete
End Sub

' ---------------------------------------------------------------------------------------------
' Document data
' ---------------------------------------------------------------------------------------------

' The document data from the document properties, or the export marker if the document has been
' exported
Private Function ReadDocumentData(doc As Document) As String
        Dim startText As String
        Dim data As String
        Dim chunk As String
        Dim i As Long
        If doc.Content.End > Len(EXPORT_MARKER) Then startText = doc.Range(0, Len(EXPORT_MARKER)).Text
        If startText = EXPORT_MARKER Or startText = OLD_EXPORT_MARKER Then
                ReadDocumentData = EXPORT_MARKER
                Exit Function
        End If
        i = 1
        Do
                chunk = GetProperty(doc, PREFS_PROPERTY & "_" & i)
                data = data & chunk
                i = i + 1
        Loop While Len(chunk) > 0
        ReadDocumentData = data
End Function

Private Sub WriteDocumentData(doc As Document, data As String)
        Dim chunkCount As Long
        Dim i As Long
        Dim propertyName As String
        chunkCount = (Len(data) + MAX_PROPERTY_LENGTH - 1) \ MAX_PROPERTY_LENGTH
        For i = 1 To chunkCount
                SetProperty doc, PREFS_PROPERTY & "_" & i, Mid$(data, (i - 1) * MAX_PROPERTY_LENGTH + 1, MAX_PROPERTY_LENGTH)
        Next i
        ' Delete properties left from longer data
        i = chunkCount + 1
        Do
                propertyName = PREFS_PROPERTY & "_" & i
                If Not PropertyExists(doc, propertyName) Then Exit Do
                If Not DeleteProperty(doc, propertyName) Then Exit Do
                i = i + 1
        Loop
End Sub

Private Function PropertyExists(doc As Document, propertyName As String) As Boolean
        Dim propertyValue As Variant
        On Error GoTo missing
        propertyValue = doc.CustomDocumentProperties(propertyName).Value
        PropertyExists = True
missing:
End Function

Private Function GetProperty(doc As Document, propertyName As String) As String
        On Error Resume Next
        GetProperty = CStr(doc.CustomDocumentProperties(propertyName).Value)
End Function

Private Sub SetProperty(doc As Document, propertyName As String, propertyValue As String)
        If PropertyExists(doc, propertyName) Then
                doc.CustomDocumentProperties(propertyName).Value = propertyValue
        Else
                doc.CustomDocumentProperties.Add Name:=propertyName, LinkToContent:=False, _
                        Type:=PROPERTY_TYPE_STRING, Value:=propertyValue
        End If
End Sub

' Deletes a property, or empties it if it can't be deleted. Returns False if neither works.
Private Function DeleteProperty(doc As Document, propertyName As String) As Boolean
        On Error Resume Next
        doc.CustomDocumentProperties(propertyName).Delete
        If Err.Number = 0 Then
                DeleteProperty = True
                Exit Function
        End If
        Err.Clear
        doc.CustomDocumentProperties(propertyName).Value = ""
        DeleteProperty = Err.Number = 0
End Function

' ---------------------------------------------------------------------------------------------
' Input and output
' ---------------------------------------------------------------------------------------------

Private Sub WriteError(errNumber As Long, errDescription As String)
        On Error Resume Next
        If callDoc Is Nothing Then Set callDoc = FindTargetDocument()
        If callDoc Is Nothing Then Exit Sub
        WriteVariables callDoc, OUTPUT_PREFIX, "1;" & errNumber & ";" & EncodeString(errDescription)
End Sub

' The document with macro input
Private Function FindTargetDocument() As Document
        Dim d As Document
        For Each d In Documents
                If VariableExists(d, INPUT_PREFIX & "1") Then
                        Set FindTargetDocument = d
                        Exit Function
                End If
        Next d
End Function

Private Function VariableExists(doc As Document, variableName As String) As Boolean
        Dim variableValue As String
        On Error GoTo missing
        variableValue = doc.Variables(variableName).Value
        VariableExists = True
        Exit Function
missing:
        VariableExists = False
End Function

' Reads the variables with a prefix and deletes them
Private Function ReadVariables(doc As Document, prefix As String) As String
        Dim first As String
        Dim pos As Long
        Dim chunkCount As Long
        Dim chunks() As String
        Dim i As Long

        first = doc.Variables(prefix & "1").Value
        doc.Variables(prefix & "1").Delete
        pos = 1
        chunkCount = ScanNumber(first, pos)
        ReDim chunks(1 To chunkCount)
        chunks(1) = Mid$(first, pos)
        For i = 2 To chunkCount
                chunks(i) = doc.Variables(prefix & i).Value
                doc.Variables(prefix & i).Delete
        Next i
        ReadVariables = Join(chunks, "")
End Function

Private Sub WriteVariables(doc As Document, prefix As String, data As String)
        Dim chunks() As String
        Dim chunkCount As Long
        Dim pos As Long
        Dim chunkLength As Long
        Dim lastChar As Long
        Dim i As Long

        ReDim chunks(1 To Len(data) \ (CHUNK_LENGTH - 16) + 2)
        pos = 1
        Do
                chunkLength = CHUNK_LENGTH - 16
                If chunkCount >= UBound(chunks) Then ReDim Preserve chunks(1 To 2 * chunkCount)
                If pos + chunkLength - 1 < Len(data) Then
                        ' Don't split a surrogate pair across variables
                        lastChar = AscW(Mid$(data, pos + chunkLength - 1, 1)) And &HFFFF&
                        If lastChar >= &HD800& And lastChar <= &HDBFF& Then chunkLength = chunkLength - 1
                End If
                chunkCount = chunkCount + 1
                chunks(chunkCount) = Mid$(data, pos, chunkLength)
                pos = pos + chunkLength
        Loop While pos <= Len(data)
        doc.Variables.Add prefix & "1", chunkCount & ";" & chunks(1)
        For i = 2 To chunkCount
                doc.Variables.Add prefix & i, chunks(i)
        Next i
End Sub

Private Function ScanNumber(data As String, pos As Long) As Long
        Dim separator As Long
        separator = InStr(pos, data, ";")
        If separator = 0 Then Err.Raise ERR_ZOTERO, , "Malformed macro input"
        ScanNumber = CLng(Mid$(data, pos, separator - pos))
        pos = separator + 1
End Function

Private Function ScanString(data As String, pos As Long) As String
        Dim separator As Long
        Dim stringLength As Long
        separator = InStr(pos, data, ":")
        If separator = 0 Then Err.Raise ERR_ZOTERO, , "Malformed macro input"
        stringLength = CLng(Mid$(data, pos, separator - pos))
        If stringLength < 0 Or separator + stringLength > Len(data) Then
                Err.Raise ERR_ZOTERO, , "Malformed macro input"
        End If
        ScanString = Mid$(data, separator + 1, stringLength)
        pos = separator + 1 + stringLength
End Function

Private Function EncodeString(s As String) As String
        EncodeString = Len(s) & ":" & s
End Function
