Attribute VB_Name = "ZoteroBatch"
' ***** BEGIN LICENSE BLOCK *****
'
' Copyright (c) 2026  Corporation for Digital Scholarship
'                     Falls Church, Virginia, USA
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

' Macros run by the Zotero plugin over Apple Events. Each Apple Event costs several milliseconds
' in large documents, while the same operations take microseconds from VBA, so these macros
' read and write all fields at once. Input comes from document variables ZOTERO_INPUT_1..n and
' output goes to ZOTERO_BATCH_1..n, with n stored in <prefix>COUNT. Strings are encoded as
' "<length>:<string>", with lengths in UTF-16 code units, and numbers as "<number>;". Fields are
' identified by note type (0 = main text, 1 = footnote, 2 = endnote) and their index within
' that story's fields.
'
' ZoteroGetFieldData output: "1;<field count>;" followed by one record per field, in story
' order: <note type>;<entry index>;<code start>;<result start>;<result end>;<text location>;
' <note index>;<code><text>
'
' ZoteroSetFieldTexts input: "1;<count>;" followed by <note type>;<entry index>;<is rich>;
' <is bibliography>;<text> per field, where rich text is pasted from the clipboard, which holds
' an RTF document with each field's text preceded by a marker (see SEGMENT_MARKER) and an end
' marker. Output: "1;<count>;" followed by <removed fields>;<resulting text> per field, where
' <removed fields> is the number of fields nested in the field that the update removed, which
' shifts the indices of later fields in the same story.
'
' ZoteroSetFieldCodes input: "1;<count>;" followed by <note type>;<entry index>;
' <expected current code>;<new code> per field. No code is changed if any field's current code
' doesn't match. Output: "1;0;" on success or "1;1;" on a mismatch.

Option Explicit

Private Const BATCH_VARIABLE_PREFIX As String = "ZOTERO_BATCH_"
Private Const INPUT_VARIABLE_PREFIX As String = "ZOTERO_INPUT_"
Private Const BATCH_FORMAT_VERSION As String = "1"
Private Const SEGMENT_MARKER As String = "ZQZSEG"
' Document variables hold at most 65,280 characters
Private Const BATCH_CHUNK_LENGTH As Long = 60000

Private batchDocument As Document
Private batchDocumentWasSaved As Boolean

Sub ZoteroGetFieldData()
	Dim doc As Document
	Dim parts() As String
	Dim partCount As Long
	Dim fieldCount As Long
	Dim data As String

	ZoteroClearFieldData

	On Error GoTo fail
	Set doc = ActiveDocument
	batchDocumentWasSaved = doc.Saved
	Set batchDocument = doc

	ReDim parts(0 To 255)
	CollectFields doc.Fields, 0, parts, partCount, fieldCount
	If doc.Footnotes.Count > 0 Then
		CollectFields doc.StoryRanges(wdFootnotesStory).Fields, 1, parts, partCount, fieldCount
	End If
	If doc.Endnotes.Count > 0 Then
		CollectFields doc.StoryRanges(wdEndnotesStory).Fields, 2, parts, partCount, fieldCount
	End If

	data = BATCH_FORMAT_VERSION & ";" & fieldCount & ";"
	If partCount > 0 Then
		ReDim Preserve parts(0 To partCount - 1)
		data = data & Join(parts, "")
	End If
	StoreData doc, data
	Exit Sub

fail:
	ZoteroClearFieldData
End Sub

Sub ZoteroSetFieldTexts()
	Dim doc As Document
	Dim tempDoc As Document
	Dim data As String
	Dim position As Long
	Dim count As Long
	Dim i As Long
	Dim noteTypes() As Long
	Dim entryIndices() As Long
	Dim isRich() As Boolean
	Dim isBibliography() As Boolean
	Dim texts() As String
	Dim richCount As Long
	Dim segmentStarts() As Long
	Dim segmentEnds() As Long
	Dim segment As Long
	Dim storyFields(0 To 2) As Variant
	Dim f As Field
	Dim fieldCount As Long
	Dim results() As String
	Dim previousScreenUpdating As Boolean

	On Error GoTo fail
	Set doc = ActiveDocument
	previousScreenUpdating = DisableScreenUpdating()
	BeginBatch doc
	data = ReadInput(doc)
	position = 1
	If ScanNumber(data, position) <> 1 Then Err.Raise vbObjectError + 1, , "Unsupported batch format"
	count = ScanNumber(data, position)
	ReDim noteTypes(0 To count - 1)
	ReDim entryIndices(0 To count - 1)
	ReDim isRich(0 To count - 1)
	ReDim isBibliography(0 To count - 1)
	ReDim texts(0 To count - 1)
	ReDim results(0 To count - 1)
	For i = 0 To count - 1
		noteTypes(i) = ScanNumber(data, position)
		entryIndices(i) = ScanNumber(data, position)
		isRich(i) = ScanNumber(data, position) <> 0
		isBibliography(i) = ScanNumber(data, position) <> 0
		texts(i) = ScanString(data, position)
		If isRich(i) Then richCount = richCount + 1
	Next i

	If richCount > 0 Then
		Set tempDoc = Documents.Add(Visible:=False)
		PasteSegments tempDoc, richCount, segmentStarts, segmentEnds
	End If

	For i = 0 To count - 1
		Set f = GetField(doc, storyFields, noteTypes(i), entryIndices(i))
		' Replacing the text removes any fields nested in this one
		fieldCount = f.Result.Fields.Count
		If isRich(i) Then
			SetRichText f, tempDoc.Range(segmentStarts(segment), segmentEnds(segment)), _
				isBibliography(i)
			segment = segment + 1
		Else
			SetPlainText f, texts(i)
		End If
		results(i) = (fieldCount - f.Result.Fields.Count) & ";" & EncodeString(f.Result.Text)
	Next i

	If Not tempDoc Is Nothing Then tempDoc.Close SaveChanges:=wdDoNotSaveChanges
	StoreData doc, BATCH_FORMAT_VERSION & ";" & count & ";" & Join(results, "")
	Application.ScreenUpdating = previousScreenUpdating
	Exit Sub

fail:
	On Error Resume Next
	If Not tempDoc Is Nothing Then tempDoc.Close SaveChanges:=wdDoNotSaveChanges
	Application.ScreenUpdating = previousScreenUpdating
End Sub

Sub ZoteroSetFieldCodes()
	Dim doc As Document
	Dim data As String
	Dim position As Long
	Dim count As Long
	Dim i As Long
	Dim noteType As Long
	Dim entryIndex As Long
	Dim expectedCode As String
	Dim storyFields(0 To 2) As Variant
	Dim flds() As Field
	Dim codes() As String
	Dim previousScreenUpdating As Boolean

	On Error GoTo fail
	Set doc = ActiveDocument
	previousScreenUpdating = DisableScreenUpdating()
	BeginBatch doc
	data = ReadInput(doc)
	position = 1
	If ScanNumber(data, position) <> 1 Then Err.Raise vbObjectError + 1, , "Unsupported batch format"
	count = ScanNumber(data, position)
	ReDim flds(0 To count - 1)
	ReDim codes(0 To count - 1)
	For i = 0 To count - 1
		noteType = ScanNumber(data, position)
		entryIndex = ScanNumber(data, position)
		expectedCode = ScanString(data, position)
		codes(i) = ScanString(data, position)
		Set flds(i) = GetField(doc, storyFields, noteType, entryIndex)
		If Len(expectedCode) > 0 Then
			If flds(i).Code.Text <> expectedCode Then
				StoreData doc, BATCH_FORMAT_VERSION & ";1;"
				Application.ScreenUpdating = previousScreenUpdating
				Exit Sub
			End If
		End If
	Next i

	For i = 0 To count - 1
		flds(i).Code.Text = codes(i)
	Next i
	StoreData doc, BATCH_FORMAT_VERSION & ";0;"
	Application.ScreenUpdating = previousScreenUpdating
	Exit Sub

fail:
	On Error Resume Next
	Application.ScreenUpdating = previousScreenUpdating
End Sub

Sub ZoteroClearFieldData()
	Dim doc As Document

	On Error Resume Next
	If batchDocument Is Nothing Then
		Set doc = ActiveDocument
	Else
		Set doc = batchDocument
	End If
	DeleteVariables doc, BATCH_VARIABLE_PREFIX
	DeleteVariables doc, INPUT_VARIABLE_PREFIX
	If Not batchDocument Is Nothing Then
		If batchDocumentWasSaved Then doc.Saved = True
		Set batchDocument = Nothing
	End If
End Sub

' Turns off screen updating, so that Word doesn't redraw after each change, and returns the
' previous setting
Private Function DisableScreenUpdating() As Boolean
	DisableScreenUpdating = Application.ScreenUpdating
	Application.ScreenUpdating = False
End Function

Private Sub BeginBatch(doc As Document)
	DeleteVariables doc, BATCH_VARIABLE_PREFIX
	batchDocumentWasSaved = False
	Set batchDocument = doc
End Sub

Private Sub DeleteVariables(doc As Document, prefix As String)
	Dim i As Long

	For i = doc.Variables.Count To 1 Step -1
		If Left$(doc.Variables(i).Name, Len(prefix)) = prefix Then
			doc.Variables(i).Delete
		End If
	Next i
End Sub

Private Sub CollectFields(flds As Fields, noteType As Long, parts() As String, partCount As Long, _
		fieldCount As Long)
	Dim f As Field
	Dim entryIndex As Long
	Dim code As String
	Dim textLocation As Long
	Dim noteIndex As Long

	For Each f In flds
		entryIndex = entryIndex + 1
		code = f.Code.Text
		If InStr(code, " ADDIN ZOTERO_") > 0 Or InStr(code, " CSL_") > 0 Then
			' An all-caps code may just be displayed in all caps, so turn that off and reread it
			If code = UCase$(code) Then
				If f.Code.Font.AllCaps <> 0 Then
					f.Code.Font.AllCaps = False
					code = f.Code.Text
				End If
			End If
			If noteType = 1 Then
				textLocation = f.Result.Footnotes(1).Reference.Start
				noteIndex = f.Result.Footnotes(1).Index
			ElseIf noteType = 2 Then
				textLocation = f.Result.Endnotes(1).Reference.Start
				noteIndex = f.Result.Endnotes(1).Index
			Else
				textLocation = f.Result.Start
				noteIndex = 0
			End If

			If partCount > UBound(parts) Then
				ReDim Preserve parts(0 To 2 * (UBound(parts) + 1) - 1)
			End If
			parts(partCount) = noteType & ";" & entryIndex & ";" & f.Code.Start & ";" _
				& f.Result.Start & ";" & f.Result.End & ";" & textLocation & ";" & noteIndex & ";" _
				& EncodeString(code) & EncodeString(f.Result.Text)
			partCount = partCount + 1
			fieldCount = fieldCount + 1
		End If
	Next f
End Sub

' Returns a field by its index in its story at the time of the first call for that story.
' Fields(i) walks the story to find each field, so each story's fields are listed once.
Private Function GetField(doc As Document, storyFields() As Variant, noteType As Long, _
		entryIndex As Long) As Field
	If IsEmpty(storyFields(noteType)) Then
		storyFields(noteType) = ListFields(GetStoryFields(doc, noteType))
	End If
	Set GetField = storyFields(noteType)(entryIndex)
End Function

Private Function GetStoryFields(doc As Document, noteType As Long) As Fields
	If noteType = 1 Then
		Set GetStoryFields = doc.StoryRanges(wdFootnotesStory).Fields
	ElseIf noteType = 2 Then
		Set GetStoryFields = doc.StoryRanges(wdEndnotesStory).Fields
	Else
		Set GetStoryFields = doc.Fields
	End If
End Function

Private Function ListFields(flds As Fields) As Variant
	Dim result() As Field
	Dim f As Field
	Dim i As Long

	ReDim result(0 To flds.Count)
	For Each f In flds
		i = i + 1
		Set result(i) = f
	Next f
	ListFields = result
End Function

' Pastes the RTF document on the clipboard into tempDoc and finds the text following each marker
Private Sub PasteSegments(tempDoc As Document, richCount As Long, segmentStarts() As Long, _
		segmentEnds() As Long)
	Dim adjustWordSpacing As Boolean
	Dim text As String
	Dim marker As String
	Dim markerPosition As Long
	Dim k As Long

	adjustWordSpacing = Options.PasteAdjustWordSpacing
	Options.PasteAdjustWordSpacing = False
	tempDoc.Content.Paste
	Options.PasteAdjustWordSpacing = adjustWordSpacing

	' Character positions match string positions only if the content is plain text
	text = tempDoc.Content.Text
	If Len(text) <> tempDoc.Content.End - tempDoc.Content.Start Then
		Err.Raise vbObjectError + 2, , "Unexpected content in pasted text"
	End If

	ReDim segmentStarts(0 To richCount - 1)
	ReDim segmentEnds(0 To richCount - 1)
	markerPosition = 1
	For k = 0 To richCount
		If k = richCount Then
			marker = SEGMENT_MARKER & "END" & SEGMENT_MARKER
		Else
			marker = SEGMENT_MARKER & k & SEGMENT_MARKER
		End If
		markerPosition = InStr(markerPosition, text, marker)
		If markerPosition = 0 Then Err.Raise vbObjectError + 3, , "Missing segment marker"
		' String positions are 1-based and character positions 0-based
		If k > 0 Then segmentEnds(k - 1) = markerPosition - 1
		If k < richCount Then segmentStarts(k) = markerPosition - 1 + Len(marker)
		markerPosition = markerPosition + Len(marker)
	Next k
End Sub

Private Sub SetRichText(f As Field, source As Range, isBibliography As Boolean)
	Dim fontSize As Single
	Dim fontName As String
	Dim fontNameOther As String
	Dim fontColorIndex As Long

	' Pasted text takes on the formatting of the existing text
	On Error Resume Next
	With f.Result.Font
		fontSize = .Size
		fontName = .Name
		fontNameOther = .NameOther
		fontColorIndex = .ColorIndex
	End With
	On Error GoTo 0

	f.Result.FormattedText = source.FormattedText

	On Error Resume Next
	If isBibliography Then f.Result.Style = wdStyleBibliography
	With f.Result.Font
		.Size = fontSize
		.Name = fontName
		.NameOther = fontNameOther
		.ColorIndex = fontColorIndex
	End With
	On Error GoTo 0
End Sub

Private Sub SetPlainText(f As Field, text As String)
	' Superscripting is more likely to come from a style change than to be intended
	On Error Resume Next
	f.Result.Font.Superscript = False
	On Error GoTo 0
	f.Result.Text = text
End Sub

Private Function ReadInput(doc As Document) As String
	Dim chunkCount As Long
	Dim chunks() As String
	Dim i As Long

	chunkCount = CLng(doc.Variables(INPUT_VARIABLE_PREFIX & "COUNT").Value)
	ReDim chunks(1 To chunkCount)
	For i = 1 To chunkCount
		chunks(i) = doc.Variables(INPUT_VARIABLE_PREFIX & i).Value
	Next i
	ReadInput = Join(chunks, "")
	DeleteVariables doc, INPUT_VARIABLE_PREFIX
End Function

Private Function ScanNumber(data As String, position As Long) As Long
	Dim separator As Long

	separator = InStr(position, data, ";")
	If separator = 0 Then Err.Raise vbObjectError + 4, , "Malformed batch input"
	ScanNumber = CLng(Mid$(data, position, separator - position))
	position = separator + 1
End Function

Private Function ScanString(data As String, position As Long) As String
	Dim separator As Long
	Dim length As Long

	separator = InStr(position, data, ":")
	If separator = 0 Then Err.Raise vbObjectError + 4, , "Malformed batch input"
	length = CLng(Mid$(data, position, separator - position))
	ScanString = Mid$(data, separator + 1, length)
	position = separator + 1 + length
End Function

Private Function EncodeString(s As String) As String
	EncodeString = Len(s) & ":" & s
End Function

Private Sub StoreData(doc As Document, data As String)
	Dim position As Long
	Dim chunkLength As Long
	Dim chunkCount As Long
	Dim lastChar As Long

	position = 1
	Do While position <= Len(data)
		chunkLength = BATCH_CHUNK_LENGTH
		If position + chunkLength - 1 < Len(data) Then
			' Don't split a surrogate pair across chunks
			lastChar = AscW(Mid$(data, position + chunkLength - 1, 1)) And &HFFFF&
			If lastChar >= &HD800& And lastChar <= &HDBFF& Then chunkLength = chunkLength - 1
		End If
		chunkCount = chunkCount + 1
		doc.Variables.Add BATCH_VARIABLE_PREFIX & chunkCount, Mid$(data, position, chunkLength)
		position = position + chunkLength
	Loop
	doc.Variables.Add BATCH_VARIABLE_PREFIX & "COUNT", CStr(chunkCount)
End Sub
