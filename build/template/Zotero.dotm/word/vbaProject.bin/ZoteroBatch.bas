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

' Macros run by the Zotero plugin over Apple Events. Reading fields one Apple Event at a time
' costs several milliseconds per event, while the same reads take microseconds from VBA, so
' ZoteroGetFieldData reads every Zotero field in the active document and stores the results in
' document variables for the plugin to read with a few events.
'
' Data format: "<version>;<field count>;" followed by one record per field, in story order:
' "<note type>;<entry index>;<code start>;<result start>;<result end>;<text location>;
' <note index>;<code length>:<code><text length>:<text>", where lengths are in UTF-16 code units.
' The data is split across ZOTERO_BATCH_1..n, with n stored in ZOTERO_BATCH_COUNT.

Option Explicit

Private Const BATCH_VARIABLE_PREFIX As String = "ZOTERO_BATCH_"
Private Const BATCH_FORMAT_VERSION As String = "1"
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

Sub ZoteroClearFieldData()
	Dim doc As Document
	Dim i As Long

	On Error Resume Next
	If batchDocument Is Nothing Then
		Set doc = ActiveDocument
	Else
		Set doc = batchDocument
	End If
	For i = doc.Variables.Count To 1 Step -1
		If Left$(doc.Variables(i).Name, Len(BATCH_VARIABLE_PREFIX)) = BATCH_VARIABLE_PREFIX Then
			doc.Variables(i).Delete
		End If
	Next i
	If Not batchDocument Is Nothing Then
		If batchDocumentWasSaved Then doc.Saved = True
		Set batchDocument = Nothing
	End If
End Sub

Private Sub CollectFields(flds As Fields, noteType As Long, parts() As String, partCount As Long, _
		fieldCount As Long)
	Dim f As Field
	Dim entryIndex As Long
	Dim code As String
	Dim text As String
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
			text = f.Result.Text

			If partCount > UBound(parts) Then
				ReDim Preserve parts(0 To 2 * (UBound(parts) + 1) - 1)
			End If
			parts(partCount) = noteType & ";" & entryIndex & ";" & f.Code.Start & ";" _
				& f.Result.Start & ";" & f.Result.End & ";" & textLocation & ";" & noteIndex & ";" _
				& Len(code) & ":" & code & Len(text) & ":" & text
			partCount = partCount + 1
			fieldCount = fieldCount + 1
		End If
	Next f
End Sub

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
