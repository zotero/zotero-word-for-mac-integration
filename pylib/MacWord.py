#!/usr/bin/python

"""
    ***** BEGIN LICENSE BLOCK *****
	
	Copyright (c) 2009  Zotero
	                    Center for History and New Media
						George Mason University, Fairfax, Virginia, USA
						http://zotero.org
	
	This program is free software: you can redistribute it and/or modify
	it under the terms of the GNU General Public License as published by
	the Free Software Foundation, either version 3 of the License, or
	(at your option) any later version.
	
	This program is distributed in the hope that it will be useful,
	but WITHOUT ANY WARRANTY; without even the implied warranty of
	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
	GNU General Public License for more details.
	
	You should have received a copy of the GNU General Public License
	along with this program.  If not, see <http://www.gnu.org/licenses/>.
    
    ***** END LICENSE BLOCK *****
"""

DEBUG = True

# NOTES FOR FUTURE MODULES:
# The Document class must be defined, and must return a Field class when
# appropriate. Other classes (FieldConversion, Properties, Clipboard) are
# specific to this implementation. The Document and Field classes must provide
# the public methods defined here.

PREFS_PROPERTIES = [u'ZOTERO_PREF', u'CITE_PREF']
BOOKMARK_REFERENCE_PROPERTIES = [u'ZOTERO_BREF', u'CITE_BREF']

FIELD_PLACEHOLDER = '{Citation}'
FIELD_PREFIXES = [u' ADDIN ZOTERO_', u' ADDIN CITE_']
BOOKMARK_PREFIXES = [u'ZOTERO_', u'CITE_']
RTF_TEMP_BOOKMARK = u'ZOTERO_TEMP_BOOKMARK'
SAVE_PROPERTIES = [u'font_size', u'name', u'other_name', u'color']

MAX_PROPERTY_LENGTH = 255
MAX_BOOKMARK_LENGTH = 50

import random, string, copy, os, tempfile, sys, traceback, xpcom.server, subprocess

from appscript import *
from xpcom import components, ServerException, nsError

DIALOG_ICON_STOP = components.interfaces.zoteroIntegrationDocument.DIALOG_ICON_STOP;
DIALOG_ICON_NOTICE = components.interfaces.zoteroIntegrationDocument.DIALOG_ICON_NOTICE;
DIALOG_ICON_CAUTION = components.interfaces.zoteroIntegrationDocument.DIALOG_ICON_CAUTION;
DIALOG_BUTTONS_OK = components.interfaces.zoteroIntegrationDocument.DIALOG_BUTTONS_OK;
DIALOG_BUTTONS_OK_CANCEL = components.interfaces.zoteroIntegrationDocument.DIALOG_BUTTONS_OK_CANCEL;
DIALOG_BUTTONS_YES_NO = components.interfaces.zoteroIntegrationDocument.DIALOG_BUTTONS_YES_NO;
DIALOG_BUTTONS_YES_NO_CANCEL = components.interfaces.zoteroIntegrationDocument.DIALOG_BUTTONS_YES_NO_CANCEL;
NOTE_FOOTNOTE = components.interfaces.zoteroIntegrationDocument.NOTE_FOOTNOTE;
NOTE_ENDNOTE = components.interfaces.zoteroIntegrationDocument.NOTE_ENDNOTE;

class XPCOMEnumerator:
	_com_interfaces_ = [components.interfaces.nsISimpleEnumerator]
	_reg_clsid_ = "{01e0a337-8c76-47c9-8c16-1f8f0a27ca3e}"
	_reg_contractid_ = "@zotero.org/Zotero/integration/enumerator?agent=MacWord;1"
	_reg_desc_ = "MacWord nsISimpleEnumerator"
	
	def __init__(self, iterator):
		self.iterator = iterator
		self.nextItem = False
	
	def getNext(self):
		if self.nextItem:
			ret = self.nextItem
		else:
			# shouldn't actually ever happen according to the spec
			self.hasMoreElements()
			ret = self.nextItem
		
		self.nextItem = False
		return ret
	
	def hasMoreElements(self):
		if self.nextItem:
			return True

		try:
			self.nextItem = self.iterator.next()
			return True
		except StopIteration:
			return False

class ExceptionAlreadyDisplayed(Exception):
	pass

class Document:
	_com_interfaces_ = [components.interfaces.zoteroIntegrationDocument]
	_reg_clsid_ = "{b8189090-48bd-11de-8a39-0800200c9a66}"
	_reg_contractid_ = "@zotero.org/Zotero/integration/document?agent=MacWord;1"
	_reg_desc_ = "MacWord Document"
	
	SUPPORT_FOOTNOTES = True
	SUPPORT_ENDNOTES = True
	
	def __init__(self, appRef):
		self.asApp = appRef
		self.asDoc = self.asApp.active_document
		self.properties = Properties(self)
		self.tempFile = TempFile()
		
		# Show appropriate error if there is no document to create, or if VBA is not installed
		if self.asDoc.get() == k.missing_value:
			if self.asApp.active_window.get() == k.missing_value:
				self.displayAlert("Zotero could not find an open document. Please open a "+ \
					"document and try again.",
					DIALOG_ICON_STOP, DIALOG_BUTTONS_OK)
				raise ExceptionAlreadyDisplayed()
			else:
				self.displayAlert('Zotero could not perform this action. Please ensure that a '+ \
					'document is open. If you have performed a custom installation of Office, '+ \
					'you may need to run the installer again, ensuring that '+ \
					'"Visual Basic for Applications" is selected.',
					DIALOG_ICON_STOP, DIALOG_BUTTONS_OK)
				raise ExceptionAlreadyDisplayed()
		
		# Capture full screen mode setting and exit
		try:
			self.restoreFullScreenMode = self.asApp.active_window.view.full_screen.get()
		except:
			self.restoreFullScreenMode = False
		if self.restoreFullScreenMode == True:
			self.asApp.active_window.view.full_screen.set(False)
			# Only re-enter full screen mode once the activate method has been called
			self.haveReactivated = False
			
	
	def displayAlert(self, text, icon=1, buttons=0):
		"""Displays a dialog"""
		icons = ['critical', 'informational', 'warning'];
		
		script = 'Tell application "Microsoft Word" to display alert("'+ \
			text.replace('\\', '\\\\').replace('"', '\\"')+'") as '+icons[icon]
		
		if buttons == DIALOG_BUTTONS_OK_CANCEL:
			buttons = ["Cancel", "OK"]
		elif buttons == DIALOG_BUTTONS_YES_NO:
			buttons = ["No", "Yes"]
		elif buttons == DIALOG_BUTTONS_YES_NO_CANCEL:
			buttons = ["Cancel", "No", "Yes"]
		else:
			buttons = ["OK"]
		
		script += ' buttons {"'+ \
			 '", "'.join([butt.replace('\\', '\\\\').replace('"', '\\"') for butt in buttons])+ \
			'"}'
		
		p = subprocess.Popen('/usr/bin/osascript', stdin=subprocess.PIPE, stdout=subprocess.PIPE)
		p.stdin.write(script.encode("utf-8", "replace"))
		p.stdin.close()
		output = p.stdout.read()
		p.stdout.close()
		
		if buttons:
			for i in range(len(buttons)):
				if output[:-1] == "button returned:"+buttons[i]:
					return i
		return True
	
	def activate(self):
		"""Brings this document to the foreground (if necessary.)"""
		self.asApp.activate()
		self.haveReactivated = True
	
	def canInsertField(self, fieldType):
		"""Determines whether a field can be inserted at the current position."""
		if self.asDoc.active_window.view.view_type.get() == k.WordNote_view:
			self.displayAlert("Zotero cannot insert a citation here because Word does not "+ \
				"support inserting fields in Notebook Layout.", 
				DIALOG_ICON_STOP, DIALOG_BUTTONS_OK)
			raise ExceptionAlreadyDisplayed()
		
		position = self.asApp.selection.story_type.get()
		return (fieldType != "Bookmark" and (position == k.footnotes_story or position == k.endnotes_story)) \
			or position == k.main_text_story
	
	def cursorInField(self, fieldType):
		"""Determines whether the cursor is in a mark.
		
		Returns the mark if it is, or false if it is not."""
		field = None
		selection = self.asApp.selection
		
		if fieldType == "Field":
			fields = selection.fields.get()
			if fields:
				field = fields[0]
			else:
				# See if there are any fields in current paragraph
				fields = selection.paragraphs[1].text_object.fields.get()
				if fields:
					# Check if they are in the selection
					selectionStart = selection.selection_start.get();
					selectionEnd = selection.selection_end.get();
					for test_field in fields:
						fieldStart = test_field.result_range.start_of_content.get()
						fieldEnd = test_field.result_range.end_of_content.get()
						# Check whether selection intersects with field. There is actually a
						# dictionary command that's supposed to do this, but it always returns true
						# when called from appscript in Word 2004, even though, oddly enough, it 
						# works when called from AppleScript
						if(selectionStart <= fieldStart and selectionEnd >= fieldStart) or \
						   (selectionStart <= fieldEnd and selectionEnd >= fieldEnd) or \
						   (selectionStart >= fieldStart and selectionStart <= fieldEnd):
							field = test_field
							break
						# No need to keep checking if we're already looking at fields past the end
						# of the selection
						elif fieldStart > selectionEnd:
							break
			
			# See if this is a Zotero field
			if field:
				codeRange = field.field_code
				rawCode = codeRange.content.get()
				for prefix in FIELD_PREFIXES:
					if rawCode.startswith(prefix):
						return Field(self, field, None, codeRange, rawCode)
		elif fieldType == "Bookmark":
			bookmarks = selection.bookmarks.get()
			if bookmarks == k.missing_value:
				return None
			for bookmark in bookmarks:
				bookmarkName = bookmark.name.get()
				for prefix in BOOKMARK_PREFIXES:
					if bookmarkName.startswith(prefix):
						return Bookmark(self, bookmark, None, bookmarkName)
		
	def getDocumentData(self):
		"""Retrieves persistent data from this document. Returns a string."""
		for prop in PREFS_PROPERTIES:
			data = self.properties.getProperty(prop);
			if data != "":
				return data
		return ""
	
	def setDocumentData(self, documentData):
		"""Sets persistent data for this document. documentData is a string."""
		self.properties.setProperty(PREFS_PROPERTIES[0], documentData)
	
	def insertField(self, fieldType, noteType, rangeToInsert=None, bookmarkName=None):
		"""Makes a field at the selection. The "rangeToInsert" and "bookmarkName"
		arguments are specific to this implementation (both are used by convert
		below) and need not be implemented."""
		
		if rangeToInsert:
			where = rangeToInsert
		else:
			if fieldType == "Bookmark" and not noteType:
				self.asApp.selection.content.set(FIELD_PLACEHOLDER)
			where = self.asApp.selection.text_object
		
		# Make footnote or endnote if desired (and cursor is in main text)
		storyType = where.story_type.get()
		if noteType and storyType == k.main_text_story:
			apNoteType = k.footnote
			if noteType == NOTE_ENDNOTE:
				apNoteType = k.endnote
			newNote = self.asApp.make(new=apNoteType, at=self.asDoc, with_properties={k.text_object:where})
			where = newNote.text_object
			where.content.set("")
			
			newEnd = newNote.note_reference.end_of_content.get();
			self.asApp.selection.selection_start.set(newEnd);
			self.asApp.selection.selection_end.set(newEnd);
		
		# Make note
		if fieldType == "Field":		# Fields
			# If creating a field in a footnote or endnote, Word might lose the
			# reference. Take preventative measures.
			if storyType != k.main_text_story:
				if storyType == k.footnotes_story:
					note = self.asDoc.footnotes[where.footnotes[1].entry_index.get()]
				elif storyType == k.endnotes_story:
					note = self.asDoc.endnotes[where.endnotes[1].entry_index.get()]
				fieldStart = where.start_of_content.get()
			
			# Create as a print date field, since otherwise there is no way to
			# add any content to the result range. Unfortunately, "make" doesn't actually return
			# the new field like it should :(
			self.asApp.make(new=k.field, at=where, with_properties={k.field_type:k.field_print_date})
				
			# Word doesn't return a working reference to new fields, so we need to find them
			if storyType != k.main_text_story:	# User created field in note
				try:
					field = where.fields[1]
					# Use to test if we actually have the field
					contentStart = field.field_code.start_of_content.get()
				except:
					contentStart = None
				# Range might not work anymore, in which case we need to find
				# the new field.
				if not contentStart or contentStart == k.missing_value:
					for noteField in note.text_object.fields.get():
						if noteField.field_code.start_of_content.get() == fieldStart+1:
							field = noteField
							break
			elif noteType:						# We created a field in note
				field = where.fields[1]
			else:								# Need to find the field in document text. Luckily
												# we know the offset of the range where it was
												# created.
				field = self.asDoc.fields[self.asDoc.create_range(start=where.start_of_content.get()-1, end_=where.end_of_content.get()+1).fields[1].entry_index.get()]
			
			field = Field(self, field)
			
			if not rangeToInsert:
				field.setText(FIELD_PLACEHOLDER, False)
		elif fieldType == "Bookmark":	# Bookmarks
			# Let convert() use a pre-defined bookmark name; otherwise,
			# generate one
			if bookmarkName == None:
				bookmarkName = BOOKMARK_REFERENCE_PROPERTIES[0]+"_"+''. \
				join([random.choice("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789") for i in range(12)])
			
			field = self.asApp.make(new=k.bookmark, at=where,
				with_properties={k.text_object:where,
				k.name:bookmarkName})
			field = Bookmark(self, field)
			
			# Select past bookmark
			newEnd = field.field.end_of_bookmark.get();
			self.asApp.selection.selection_start.set(newEnd);
			self.asApp.selection.selection_end.set(newEnd);
		
		if not rangeToInsert:
			# Add placeholder text and set code only if no rangeToInsert
			# specified (if a rangeToInsert was specified, this function
			# was called from convert())
			field.setCode("")
		return field
	
	def getFields(self, fieldType):
		"""Gets a sorted list of all fields from this document."""
		fields = []
		if fieldType == "Field":		# Fields
			# Get fields from document
			collections = self._getCollections()
			
			# Append field objects
			for noteType, collectionFields in collections.items():
				# 4X speed improvement from iterating by index instead of
				# using collection.get()
				try:
					maxField = collectionFields[-1].entry_index.get()
				except:
					continue
				if maxField == k.missing_value:
					continue
				
				for i in range(1, maxField+1):
					field = collectionFields[i]
					codeRange = field.field_code
					rawCode = codeRange.content.get()
					if rawCode != k.missing_value:
						for prefix in FIELD_PREFIXES:
							if rawCode.startswith(prefix):
								fields.append(Field(self, field, noteType, codeRange, rawCode))
		elif fieldType == "Bookmark":	# Bookmarks
			getBookmarks = self.asDoc.bookmarks.get()
			if getBookmarks:
				for bookmark in getBookmarks:
					bookmarkName = bookmark.name.get()
					if bookmarkName != k.missing_value:
						for prefix in BOOKMARK_PREFIXES:
							if bookmarkName.startswith(prefix):
								fields.append(Bookmark(self, bookmark, None, bookmarkName))
		
		fields.sort()
		return XPCOMEnumerator(fields.__iter__())
	
	def convert(self, fieldEnumerator, toFieldType, noteTypes):
		"""Convert "fields" to a different fieldType, or a different noteType.
		This function is unfortunately complicated by AppleScript's tendency
		to reference fields by index and poor API support."""
		
		# TODO: Move punctuation around magically
		# Get all fields
		fields = []
		while fieldEnumerator.hasMoreElements():
			fields.append(xpcom.server.UnwrapObject(fieldEnumerator.getNext()))
		
		# Holds the py-appscript object corresponding to Word's field array
		collections = self._getCollections()
		conversions = []
		offsets = {0:0, 1:0, 2:0}
		
		# Fields are referenced by index, so we save the fields for each
		# portion of the document separately
		for i in range(len(fields)):
			field = fields[i]
			if hasattr(field, "bookmarkName"):		 # Bookmark
				fromFieldType = "Bookmark"
				ref = field
			else:									 # Field
				fromFieldType = "Field"
				ref = field.field.entry_index.get()
			
			conversions.append(FieldConversion(ref, fields[i].textLocation,
				fromFieldType, field.noteType, noteTypes[i]))
		
		conversions.sort()
		
		# Loop through fields in different sections separately, to keep track of
		# indices
		while conversions:
			cnv = conversions.pop(0)
			
			# Get field and other relevant information
			if cnv.fromFieldType == "Field":		# Field; Index may have changed
				asField = collections[cnv.fromNoteType][cnv.ref+offsets[cnv.fromNoteType]]
				fieldRange = asField.result_range
			else:						# Bookmark; Persistent
				asField = cnv.ref.field
				fieldRange = cnv.ref.fieldRange
			
			# First handle converting between footnote, endnote, and inline
			note = None
			if cnv.fromNoteType == NOTE_FOOTNOTE:
				note = self.asDoc.footnotes[fieldRange.footnotes[1].entry_index.get()]
			elif cnv.fromNoteType == NOTE_ENDNOTE:
				note = self.asDoc.endnotes[fieldRange.endnotes[1].entry_index.get()]
			
			# Word has special facilities for converting footnotes/endnotes
			if (cnv.toNoteType == NOTE_ENDNOTE and cnv.fromNoteType == NOTE_FOOTNOTE) \
					or (cnv.toNoteType == NOTE_FOOTNOTE and cnv.fromNoteType == NOTE_ENDNOTE):
				# Collect all fields in this same footnote (in case there
				# are multiple)
				noteCnvs = []
				nNoteCnv = 0
				while nNoteCnv < len(conversions) \
						and conversions[nNoteCnv].textLocation == cnv.textLocation:
					noteCnvs.append(conversions[nNoteCnv])
					nNoteCnv += 1
				
				# Keep data needed to re-create bookmarks, because the
				# conversion will delete them
				if cnv.fromFieldType == "Bookmark":
					oldNoteStart = note.text_object.start_of_content.get()
					allCnvs = [cnv]+noteCnvs
					bookmarkStarts = []
					bookmarkEnds = []
					bookmarkNames = []
					for noteCnv in allCnvs:
						bookmarkStarts.append(noteCnv.ref.field.start_of_bookmark.get())
						bookmarkEnds.append(noteCnv.ref.field.end_of_bookmark.get())
						bookmarkNames.append(noteCnv.ref.bookmarkName)
				
				# Convert fields and get the mark
				note = self._noteSwap(cnv.fromNoteType, note)
				
				if cnv.fromFieldType == "Field":
					offsets[cnv.fromNoteType] -= nNoteCnv+1
					cnv.fromNoteType = cnv.toNoteType
					
					if toFieldType == "Field":
						del conversions[0:nNoteCnv]
					elif toFieldType == "Bookmark":
						# Need to update conversions so that notes can be
						# converted to bookmarks later
						asField = note.text_object.fields[1]
						fieldRange = asField.result_range
						
						offsets[cnv.fromNoteType] += nNoteCnv+1
						for noteCnv in noteCnvs:
							noteCnv.ref += asField.entry_index.get()-cnv.ref-offsets[cnv.fromNoteType]
							noteCnv.fromNoteType = cnv.toNoteType
				elif cnv.fromFieldType == "Bookmark":
					startDiff = note.text_object.start_of_content.get()-oldNoteStart
					for i in range(len(allCnvs)):
						allCnvs[i].fromNoteType = cnv.toNoteType
						# Regenerate bookmarks
						newField = self.insertField(1, 0, note.text_object, bookmarkNames[i])
						newField.field.start_of_bookmark.set(bookmarkStarts[i]+startDiff)
						newField.field.end_of_bookmark.set(bookmarkEnds[i]+startDiff)
			
			# If converting a non-empty note to inline, don't do anything to the
			# note; only play with the fields
			inlineNoteField = note \
				and note.text_object.content.get() != fieldRange.content.get()
			
			# Skip further processing if it's unnecessary
			if cnv.fromFieldType == toFieldType and \
				(cnv.fromNoteType == cnv.toNoteType or inlineNoteField):
				continue
			
			if cnv.fromFieldType == "Field":
				rawCode = asField.field_code.content.get()
				for prefix in FIELD_PREFIXES:
					if(rawCode.startswith(prefix)):
						code = rawCode[len(prefix):-1]
				
				if not cnv.fromNoteType:	# Convert from in-text citation
					# Delete field, but preserve range
					fieldStart = asField.field_code.start_of_content.get()
					asField.delete()
					where = self.asDoc.create_range(start=fieldStart-1, end_=fieldStart-1)
					# Insert new field
					newField = self.insertField(toFieldType, cnv.toNoteType, where)
				elif cnv.fromNoteType == cnv.toNoteType or inlineNoteField:
					# Convert fields inside a note to bookmarks
					newField = self.insertField(toFieldType, 0, fieldRange)
					newField.field.start_of_bookmark.set(newField.field.end_of_bookmark.get()+1)
					asField.delete()
				else:
					# Conversion between note and inline; need to generate new field
					# at note reference, rather than in note
					newField = self.insertField(toFieldType, cnv.toNoteType, note.note_reference.collapse_range())
					# Delete old content
					note.note_reference.content.set("")
				
				offsets[cnv.fromNoteType] -= 1
				newField.setCode(code)
			elif cnv.fromFieldType == "Bookmark":
				# Find appropriate range to overwrite depending on type of
				# citation we're convering from
				if cnv.fromNoteType and not cnv.toNoteType \
						and not inlineNoteField:		# Note to in-text
					# Delete old content and restore range
					referenceStart = note.note_reference.start_of_content.get()
					note.note_reference.content.set("")
					where = self.asDoc.create_range(start=referenceStart, end_=referenceStart)
				elif not cnv.fromNoteType and cnv.toNoteType \
						and not inlineNoteField:		# In-text to note
					# Delete old content and restore range
					bookmarkStart = asField.start_of_bookmark.get()
					fieldRange.content.set("")
					where = self.asDoc.create_range(start=bookmarkStart, end_=bookmarkStart)
				else:									# In-text to in-text
					where = fieldRange
				
				# Create bookmark or field
				if toFieldType == "Field":				# Field
					code = cnv.ref.getCode()
					newField = self.insertField(toFieldType, cnv.toNoteType, where)
					newField.setCode(code)
					
					# Word 2004: If the bookmark still exists, delete it.
					try:
						fieldRange.content.set("")
					except:
						pass
				elif toFieldType == 1:				# Bookmark
					self.insertField(toFieldType, cnv.toNoteType, where, cnv.ref.bookmarkName)
			
			# Increment offsets appropriately
			if toFieldType == "Field":
				offsets[cnv.toNoteType] += 1
	
	def setBibliographyStyle(self, firstLineIndent, bodyIndent, lineSpacing, entrySpacing, tabStops):
		"""Sets the style for the bibliography."""
		# Get or make bibliography style
		try:
			bibStyle = self.asDoc.Word_styles["Bibliography"].get()
			existingTabStops = bibStyle.paragraph_format.tab_stops.get()
			if existingTabStops != k.missing_value:
				[existingTabStop.clear() for existingTabStop in existingTabStops];
		except:
			bibStyle = self.asApp.make(at=self.asDoc, new=k.Word_style, with_properties={ \
				k.name_local:"Bibliography", k.style_type:k.style_type_paragraph, \
				k.base_style:k.style_normal})
		
		# Set properties
		bibStyle.paragraph_format.first_line_indent.set(firstLineIndent/20)
		bibStyle.paragraph_format.paragraph_format_left_indent.set(bodyIndent/20)
		bibStyle.paragraph_format.line_spacing.set(lineSpacing/20)
		bibStyle.paragraph_format.space_after.set(entrySpacing/20)
		[self.asApp.make(at=bibStyle.paragraph_format, new=k.tab_stop, \
			with_properties={k.alignment:k.align_tab_left, \
			k.tab_leader:k.tab_leader_spaces, k.tab_stop_position:tabStop/20}) \
			for tabStop in tabStops]		
	
	def cleanup(self):
		"""Run on exit to clean up anything we played with..."""
		# Restore full screen mode if necessary
		if self.restoreFullScreenMode == True and self.asApp.frontmost.get() and self.haveReactivated:
			self.asDoc.active_window.view.full_screen.set(True)
		# Delete temporary file
		if self.tempFile.path:
			os.unlink(self.tempFile.path)
	
	def _getCollections(self):
		"""Gets the contents of the main body, footnote, and endnote collections. Specific to this
		implementation."""
		collections = {0:self.asDoc.fields,
			1:self.asDoc.get_story_range(story_type=k.footnotes_story).fields,
			2:self.asDoc.get_story_range(story_type=k.endnotes_story).fields}
		return collections
	
	def _noteSwap(self, fromNoteType, note):
		"""Converts a footnote to an endnote, or vice versa. Specific to this implementation."""
		# Get a range containing the note reference, so we don't lose the note
		rfRange = self.asDoc.create_range(
			start=note.note_reference.start_of_content.get(),
			end_=note.note_reference.end_of_content.get())
		
		# Do the conversion
		if fromNoteType == NOTE_FOOTNOTE:
			note.text_object.footnote_options.footnote_convert()
			note = self.asDoc.endnotes[rfRange.endnotes[1].entry_index.get()]
		elif fromNoteType == NOTE_ENDNOTE:	
			note.text_object.endnote_options.endnote_convert()
			note = self.asDoc.footnotes[rfRange.footnotes[1].entry_index.get()]
		
		return note

class Document_2004(Document):
	def __extractFieldsFromNotes(self, notes):
		"""Extracts fields from a given collection. Not included in original MacWord module."""
		try:
			maxField = notes[-1].entry_index.get()
		except:
			return []
		
		if maxField == k.missing_value or maxField == 0:
			return []
		
		# pad with "False" to mimic appscript behavior
		fields = [False]
		for i in range(1, maxField+1):
			fields.extend(notes[i].text_object.fields.get())
		
		return fields
	
	def _getCollections(self):
		"""Overload __getCollections to fix issues with get story range in Word 2004."""
		collections = {0:self.asDoc.fields,
			1:self.__extractFieldsFromNotes(self.asDoc.footnotes),
			2:self.__extractFieldsFromNotes(self.asDoc.endnotes)}
		
		return collections
	
	def _noteSwap(self, fromNoteType, note):
		"""Overload _noteSwap because footnote_convert() and endnote_convert() crash Word"""
		# Get a range containing the note reference, so we don't lose the note
		rfRange = note.note_reference.move_end_of_range()
		createRange = note.note_reference.collapse_range(direction=k.collapse_end)
		
		# Create a new note of the appropriate type
		if fromNoteType == NOTE_FOOTNOTE:
			self.asDoc.make(new=k.endnote, at=createRange)
			newNote = self.asDoc.endnotes[rfRange.endnotes[1].entry_index.get()]
		elif fromNoteType == NOTE_ENDNOTE:	
			self.asDoc.make(new=k.footnote, at=createRange)
			newNote = self.asDoc.footnotes[rfRange.footnotes[1].entry_index.get()]
		
		# Transfer the text content
		newNote.text_object.formatted_text.set(note.text_object.formatted_text)
		note.delete()
		
		return newNote

class Field:
	_com_interfaces_ = [components.interfaces.zoteroIntegrationField]
	_reg_clsid_ = "{b3581f3c-1b53-4514-926f-14df94df5042}"
	_reg_contractid_ = "@zotero.org/Zotero/integration/field?agent=MacWord&type=Field;1"
	_reg_desc_ = "MacWord Field"
	
	def __init__(self, wpDoc, field, noteType=None, codeRange=None, rawCode=None):
		self.wpDoc = wpDoc
		self.field = field
		self.rawCode = rawCode
		if codeRange:
			self.fieldRange = codeRange
		else:
			self.fieldRange = field.field_code;
		self.noteType = noteType
		self._getTextLocation()
	
	def __cmp__(x, y):
		c = cmp(x.textLocation, y.textLocation)
		if c != 0:
			return c
		
		# If in a note, compare field positions inside the note
		if x.note and y.note:
			if not x.noteLocation:
				x.noteLocation = x.fieldRange.start_of_content.get()
			if not y.noteLocation:
				y.noteLocation = y.fieldRange.start_of_content.get()
			return cmp(x.noteLocation, y.noteLocation)
		
		return 0
	
	def delete(self):
		"""Deletes this field"""
		if self.note \
			and self.note.text_object.start_of_content.get() == self.fieldRange.start_of_content.get()-1 \
			and self.note.text_object.end_of_content.get() == self.displayFieldRange.end_of_content.get()+1:
			self.note.delete()
		else:
			self.field.delete()
	
	def removeCode(self):
		"""Removes field, but maintains its contents"""
		self.field.unlink()
	
	def select(self):
		"""Selects this field"""
		self.displayFieldRange.select()
	
	def setText(self, string, isRich, deleteBM=True):
		"""Sets the text inside this field to a specified pseudo-RTF formatted
		   text string. deleteBM is specific to this implementation."""
		if isRich:
			# Make sure temp bookmark is gone
			try:
				self.wpDoc.asDoc.bookmarks[RTF_TEMP_BOOKMARK].delete()
			except:
				pass
			
			# Save properties
			savedProperties = [getattr(self.displayFieldRange.font_object, prop).get() \
				for prop in SAVE_PROPERTIES]
			
			# Write RTF to a file
			hfsPath = self.wpDoc.tempFile.write("{\\rtf {\\bkmkstart "+RTF_TEMP_BOOKMARK+"}"+ \
				string[6:-1]+"{\\bkmkend "+RTF_TEMP_BOOKMARK+"}}")
			self.displayFieldRange.content.set("")
			self.wpDoc.asApp.insert_file(at=self.displayFieldRange, \
				file_name=self.wpDoc.tempFile.hfsPath, file_range=RTF_TEMP_BOOKMARK, \
				confirm_conversions=False)
			self.displayFieldRange.bookmarks[RTF_TEMP_BOOKMARK].delete()
			
			# Set style
			if self.getCode().startswith("BIBL"):
				try:
					# oddly, it wants a string and not a style object
					self.displayFieldRange.style.set("Bibliography")
				except:
					pass
			
			# Set properties back to saved
			[getattr(self.displayFieldRange.font_object, SAVE_PROPERTIES[i]).set(savedProperties[i]) \
				for i in range(len(SAVE_PROPERTIES)) if savedProperties[i] != k.missing_value]
		else:
			# Just set content of text object
			self.displayFieldRange.content.set(string)
	
	def setCode(self, code):
		"""Sets some (non-user-readable) code to accompany this field"""
		self.rawCode = FIELD_PREFIXES[0]+code+" "
		self.field.field_code.content.set(self.rawCode)
	
	def getCode(self):
		"""Returns the current code"""
		if not self.rawCode:
			self.rawCode = self.field.field_code.content.get()
		for prefix in FIELD_PREFIXES:
			if(self.rawCode.startswith(prefix)):
				if self.rawCode[-1] == " ":
					return self.rawCode[len(prefix):-1]
				else:
					return self.rawCode[len(prefix):]
	
	def equals(self, field):
		"""Returns true if field and this field refer to the same field"""
		return xpcom.server.UnwrapObject(field) == self
	
	def getNoteIndex(self):
		"""Returns the index of the note in which this field resides"""
		if self.note:
			return self.note.entry_index.get()
		return 0
	
	def _getTextLocation(self):
		"""Adds note and textLocation properties to this instance. This is
		protected and not private, so that the Bookmark class can get at it, but
		is never used by Zotero.py"""
		# Try to avoid the AppleEvent call here if possible
		if self.noteType == None:
			storyType = self.fieldRange.story_type.get()
			if storyType == k.footnotes_story:
				self.noteType = NOTE_FOOTNOTE
			elif storyType == k.endnotes_story:
				self.noteType = NOTE_ENDNOTE
			else:
				self.noteType = 0
		
		# Save note and textLocation
		if self.noteType == NOTE_FOOTNOTE:
			self.note = self.fieldRange.footnotes[1]
			self.textLocation = self.note.note_reference.start_of_content.get()
			self.noteLocation = None
		elif self.noteType == NOTE_ENDNOTE:
			self.note = self.fieldRange.endnotes[1]
			self.textLocation = self.note.note_reference.start_of_content.get()
			self.noteLocation = None
		else:
			self.note = None
			self.textLocation = self.fieldRange.start_of_content.get()
	
	@property
	def displayFieldRange(self):
		return self.field.result_range

class Bookmark(Field):
	_reg_clsid_ = "{92771357-7c15-4a49-b963-7cd91886b06b}"
	_reg_contractid_ = "@zotero.org/Zotero/integration/field?agent=MacWord&type=Bookmark;1"
	_reg_desc_ = "MacWord Bookmark"
	
	def __init__(self, wpDoc, field, noteType=None, bookmarkName=None):
		self.wpDoc = wpDoc
		
		# Save useful information
		if bookmarkName:
			self.bookmarkName = bookmarkName
		else:
			self.bookmarkName = field.name.get()
			if self.bookmarkName == k.missing_value:
				# This is not the clean way of doing things, but apparently in Office 2004, it is
				# sometimes the only way of doing things
				fieldName = repr(field)
				for prop in BOOKMARK_REFERENCE_PROPERTIES:
					nameIndex = fieldName.find(prop)
					if nameIndex != -1:
						self.bookmarkName = fieldName[nameIndex:fieldName.rindex("'")]
						break
			
		# Re-index by field name
		self.field = wpDoc.asDoc.bookmarks[self.bookmarkName]
		
		self.fieldRange = self.displayFieldRange = self.field.text_object
		self.noteType = noteType
		self._getTextLocation()
	
	def delete(self):
		"""Deletes this bookmark, and any properties"""
		if self.note and self.note.text_object.content.get() == self.fieldRange.content.get():
			self.note.delete()
		else:
			self.fieldRange.content.set("")
		
		self.wpDoc.properties.setProperty(self.bookmarkName, "")
	
	def removeCode(self):
		"""Removes bookmark, but maintains its contents"""
		self.field.delete()
	
	def setCode(self, code):
		"""Sets some (non-user-readable) code to accompany this bookmark"""
		
		# Set the property corresponding to the name
		self.wpDoc.properties.setProperty(self.bookmarkName, BOOKMARK_PREFIXES[0]+code)
	
	def setText(self, string, isRich):
		# Because ranges don't work in footnotes and endnotes, this is
		# unfortunately complex
		#
		# WHAT DOESN'T WORK:
		# - Directly setting content of self.fieldRange (deletes bookmark)
		# - Setting content of self.fieldRange and re-generating bookmark
		#   (range no longer works once content is changed)
		# - Adding content to end of bookmark, then deleting content at
		#   beginning (cannot use custom-defined ranges in footnotes)
		
		if isRich:
			# Make sure temp bookmark is gone
			try:
				self.wpDoc.asDoc.bookmarks[RTF_TEMP_BOOKMARK].delete()
			except:
				pass
			
			# Check whether cursor is at end of citation
			selection = self.wpDoc.asApp.selection
			selectionAtEnd = selection.selection_end.get() == self.fieldRange.end_of_content.get()
			
			# Save properties
			savedProperties = [getattr(self.displayFieldRange.font_object, prop).get() \
				for prop in SAVE_PROPERTIES]
			
			# Rename bookmark
			tempBookmark = self.wpDoc.asDoc.make(new=k.bookmark, at=self.fieldRange,
				with_properties={k.name:RTF_TEMP_BOOKMARK})
			self.field.delete()
			tempBookmark = self.wpDoc.asDoc.bookmarks[RTF_TEMP_BOOKMARK]
			
			# Insert RTF
			hfsPath = self.wpDoc.tempFile.write("{\\rtf {\\bkmkstart "+self.bookmarkName+"}"+ \
				string[6:-1]+"{\\bkmkend "+self.bookmarkName+"}}")
			self.wpDoc.asApp.insert_file(at=tempBookmark.text_object, \
				file_name=self.wpDoc.tempFile.hfsPath, file_range=self.bookmarkName, \
				confirm_conversions=False)
			
			# Delete old bookmark
			tempBookmark.text_object.content.set("")
			
			try:
				self.wpDoc.asDoc.bookmarks[RTF_TEMP_BOOKMARK].delete()
			except:
				pass
			
			# Set style
			if self.getCode().startswith("BIBL"):
				try:
					# oddly, it wants a string and not a style object
					self.displayFieldRange.style.set("Bibliography")
				except:
					pass
			
			# Set properties back to saved
			[getattr(self.displayFieldRange.font_object, SAVE_PROPERTIES[i]).set(savedProperties[i]) \
				for i in range(len(SAVE_PROPERTIES)) if savedProperties[i] != k.missing_value]
			
			# If selection was at end of mark, put it there again
			if selectionAtEnd:
				selection.selection_start.set(self.fieldRange.end_of_content.get())
				selection.font_object.reset()
		else:
			# Find a reference point in the appropriate story
			if self.noteType == NOTE_FOOTNOTE:
				rfRange = self.wpDoc.asDoc.footnotes[self.note.entry_index.get()].text_object
			elif self.noteType == NOTE_ENDNOTE:
				rfRange = self.wpDoc.asDoc.endnotes[self.note.entry_index.get()].text_object
			else:
				rfRange = self.wpDoc.asDoc.text_object
			
			# Get some data about this field
			oldStart = self.field.start_of_bookmark.get()
			oldEnd = self.field.end_of_bookmark.get()
			oldStoryEnd = rfRange.end_of_content.get()
			
			# Set text (deletes bookmark)
			Field.setText(self, string, False)
			
			# Re-generate bookmark (an empty bookmark wouldn't have gotten erased)
			if oldStart != oldEnd:
				self.wpDoc.asDoc.make(new=k.bookmark, at=rfRange,
					with_properties={k.name:self.bookmarkName})
			
			# Move around text
			self.field.start_of_bookmark.set(oldStart)
			self.field.end_of_bookmark.set(oldEnd + \
			rfRange.end_of_content.get() - oldStoryEnd)
		
	def getCode(self):
		"""Returns the current code"""
		rawCode = self.wpDoc.properties.getProperty(self.bookmarkName)
		for prefix in BOOKMARK_PREFIXES:
			if rawCode.startswith(prefix):
				return rawCode[len(prefix):]
		
class Properties:
	def __init__(self, wpDoc):
		self.wpDoc = wpDoc
		self.asProperties = wpDoc.asDoc.custom_document_properties
	
	def getProperty(self, property):
		"""Gets a document-specific property."""
		i = 0
		propertyValue = ""
		
		while True:
			i = i + 1
			try:
				propertyValue += self.asProperties[property+"_"+repr(i)].value.get()
			except:
				break
		
		return propertyValue
	
	def setProperty(self, propertyName, value):
		"""Sets a document-specific property."""
		i = 0
		
		# update existing property values or add new ones
		while value != "":
			i = i + 1
			docPropertyName = propertyName+u'_'+repr(i)
			docPropertyValue = value[0:MAX_PROPERTY_LENGTH]
			value = value[MAX_PROPERTY_LENGTH:]
			
			exists = False
			try:
				property = self.asProperties[docPropertyName]
				if property.exists():
					property.value.set(docPropertyValue)
					exists = True
			except:
				pass
			
			if not exists:
				# document property does not already exist; add one
				self.wpDoc.asApp.make(new=k.custom_document_property,
					at=self.wpDoc.asDoc.end,
					with_properties={k.name:docPropertyName, k.value:docPropertyValue})
		
		# delete unnecessary fields
		while True:
			i = i + 1
			try:
				property = self.asProperties[propertyName+u'_'+repr(i)]
				if property.exists():
					property.delete()
				else:
					break
			except:
				break

class FieldConversion:
	def __init__(self, ref, textLocation, fromFieldType, fromNoteType, toNoteType):
		self.ref = ref
		self.textLocation = textLocation
		self.fromFieldType = fromFieldType
		self.fromNoteType = fromNoteType
		self.toNoteType = toNoteType
	
	def __cmp__(x, y):
		c = cmp(x.textLocation, y.textLocation)
		if c != 0:
			return c
		return cmp(x.ref, y.ref)

class TempFile:
	path = None
	hfsPath = None
	def write(self, string):
		if not self.path:
			(file, self.path) = tempfile.mkstemp('.rtf')
			file = os.fdopen(file, "w")
			
			# Also store HFS path
			import Carbon.CF, Carbon.CoreFoundation
			self.hfsPath = Carbon.CF.toCF(self.path). \
				CFURLCreateWithFileSystemPath(Carbon.CoreFoundation.kCFURLPOSIXPathStyle, True). \
				CFURLCopyFileSystemPath(Carbon.CoreFoundation.kCFURLHFSPathStyle).toPython()
		else:
			file = open(self.path, "w")
		
		file.write(string.encode("macroman", "replace"))
		file.close()

def main(doc):
	z = Zotero(doc)
	try:
		try:
			getattr(z, sys.argv[1])()
		except:
			info = sys.exc_info()
			exception = traceback.format_exception(info[0], info[1], info[2])
			doc.displayAlert("An unexpected error occured while performing this operation:\n\n"+\
				"\n"+''.join(exception[-4:]), 0)
	finally:
		z.cleanup()
