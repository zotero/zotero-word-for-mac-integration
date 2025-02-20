/*
	***** BEGIN LICENSE BLOCK *****

	Copyright (c) 2011  Zotero
	Center for History and New Media
	George Mason University, Fairfax, Virginia, USA
	http://zotero.org

	Zotero is free software: you can redistribute it and/or modify
	it under the terms of the GNU Affero General Public License as published by
	the Free Software Foundation, either version 3 of the License, or
	(at your option) any later version.

	Zotero is distributed in the hope that it will be useful,
	but WITHOUT ANY WARRANTY; without even the implied warranty of
	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
	GNU Affero General Public License for more details.

	You should have received a copy of the GNU Affero General Public License
	along with Zotero.  If not, see <http://www.gnu.org/licenses/>.

	***** END LICENSE BLOCK *****
 */

#include "zoteroMacWordIntegration.h"

NSString* EXPORTED_DOCUMENT_MARKER[] = {@"ZOTERO_TRANSFER_DOCUMENT", @"ZOTERO_EXPORTED_DOCUMENT", NULL};

// Prototypes for file-local objects and functions
void freeFieldList(listNode_t* fieldList, bool freeFields);
statusCode getFieldCollections(document_t *doc, NSArray** fieldCollections);
statusCode extractFieldsFromNotes(SBElementArray* notes, 
								  NSArray** returnValue);
statusCode noteSwap(document_t *doc, WordFootnote* sbNote,
					unsigned short noteType,
					WordFootnote **returnValue);
statusCode setFieldAdjacency(field_t *fieldA, field_t *fieldB);

// Frees a document struct.
void freeDocument(document_t* doc) {
	[doc->lock lock];
	
	// Restore full screen mode
	if(doc->restoreFullScreenMode) {
		IGNORING_SB_ERRORS_BEGIN
		[doc->sbView setFullScreen:YES];
		IGNORING_SB_ERRORS_END
	}
	
	// Free allocated fields
	freeFieldList(doc->allocatedFieldsStart, true);
	
	// Free allocated field lists
	listNode_t* nextNode = doc->allocatedFieldListsStart;
	while(nextNode) {
		listNode_t* currentNode = nextNode;
		freeFieldList(currentNode->value, false);
		nextNode = currentNode->next;
		free(currentNode);
	}
	
	// Free document
	[doc->sbApp release];
	[doc->sbDoc release];
	[doc->sbView release];
	[doc->sbProperties release];
	[doc->lock unlock];
	[doc->lock release];
	free(doc->wordPath);
	free(doc);
}

// Free a field list. The second parameter determines whether the fields inside
// the list are freed, or just the list itself.
void freeFieldList(listNode_t* fieldList, bool freeFields) {
	listNode_t* nextNode = fieldList;
	while(nextNode) {
		listNode_t* currentNode = nextNode;
		if(freeFields) {
			freeField((field_t*) currentNode->value);
		}
		nextNode = currentNode->next;
		free(currentNode);
	}
}

// Activates Word to deal with a document
statusCode activate(document_t *doc) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->lock lock];
	
	[doc->sbApp activate];
	CHECK_STATUS_LOCKED(doc)
	RETURN_STATUS_LOCKED(doc, STATUS_OK)
	HANDLE_EXCEPTIONS_END
}

// Displays an alert within Word
statusCode displayAlert(document_t *doc, char const dialogText[],
						unsigned short icon, unsigned short buttons,
						unsigned short *returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->lock lock];
	
	// Make button list descriptor
	NSArray* buttonArray;
	if(buttons == DIALOG_BUTTONS_OK_CANCEL) {
		buttonArray = [NSArray arrayWithObjects:@"Cancel", @"OK", nil];
	} else if(buttons == DIALOG_BUTTONS_YES_NO) {
		buttonArray = [NSArray arrayWithObjects:@"No", @"Yes", nil];
	} else if(buttons == DIALOG_BUTTONS_YES_NO_CANCEL) {
		buttonArray = [NSArray arrayWithObjects:@"Cancel", @"No", @"Yes",
					   nil];
	} else {
		buttonArray = [NSArray arrayWithObjects:@"OK", nil];
	}
	
	NSAppleEventDescriptor *buttonList = [NSAppleEventDescriptor
										  listDescriptor];
	for(NSString* buttonName in buttonArray) {
		[buttonList insertDescriptor:
		 [NSAppleEventDescriptor descriptorWithString:buttonName] atIndex:0];
	}
	
	// Get icon
	NSAppleEventDescriptor *dialogType;
	if(icon == DIALOG_ICON_CAUTION) {
		dialogType = [NSAppleEventDescriptor descriptorWithEnumCode:'warN'];
	} else if(icon == DIALOG_ICON_NOTICE) {
		dialogType = [NSAppleEventDescriptor descriptorWithEnumCode:'infA'];
	} else {
		dialogType = [NSAppleEventDescriptor descriptorWithEnumCode:'criT'];
	}
	
	// Display dialog
	NSString* dialogTextNS = [NSString stringWithUTF8String:dialogText];
	NSDictionary *response = [doc->sbApp sendEvent:'syso' id:'disA'
						  parameters:'----', dialogTextNS,
							  'as A', dialogType,
							  'btns', buttonList, nil];
	CHECK_STATUS_LOCKED(doc);
	
	// Check button clicked
	if(returnValue != NULL) {
		NSString *buttonPressed = [response objectForKey:
								   [NSNumber numberWithInt:(int) 'bhit']];
		*returnValue = [buttonArray indexOfObject:buttonPressed];
	}
	
	RETURN_STATUS_LOCKED(doc, STATUS_OK);
	HANDLE_EXCEPTIONS_END
}

// Disables Track Changes settings so that we can read field codes
statusCode prepareReadFieldCode(document_t *doc) {
	[doc->lock lock];
	
	// Hmm. I thought I had seen this before somewhere
	// https://github.com/zotero/zotero-word-for-windows-integration/blob/dc9e67fe7a3547def56b0efcb04544a1762ccbe5/build/zoteroWinWordIntegration/document.cpp#L285-L306
	// Seems like the same codebase also means the same bugs https://twitter.com/Schwieb/status/954037656677072896
	// (Cannot disable track changes while cursor in note)
	if(doc->statusInsertionsAndDeletions) {
		ENSURE_OK_LOCKED(doc, moveCursorOutOfNote(doc));
		[doc->sbView setShowInsertionsAndDeletions:NO];
		CHECK_STATUS_LOCKED(doc)
		doc->statusInsertionsAndDeletions = NO;
	}
	
	if(doc->statusFormatChanges) {
		ENSURE_OK_LOCKED(doc, moveCursorOutOfNote(doc));
		[doc->sbView setShowFormatChanges:NO];
		CHECK_STATUS_LOCKED(doc)
		doc->statusFormatChanges = NO;
	}

	RETURN_STATUS_LOCKED(doc, STATUS_OK)
}

// Determines whether it is possible to insert a field at the current cursor
// position
statusCode canInsertField(document_t *doc, const char fieldType[],
						  bool* returnCode) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->lock lock];
	
	ENSURE_OK_LOCKED(doc, restoreCursor(doc))

	WordWdStoryType position = [[doc->sbApp selection] storyType];
	CHECK_STATUS_LOCKED(doc)
	*returnCode = (strcmp(fieldType, "Bookmark") != 0
				   && (position == WordWdStoryTypeFootnotesStory
					   || position == WordWdStoryTypeEndnotesStory))
					   || position == WordWdStoryTypeMainTextStory;
	RETURN_STATUS_LOCKED(doc, STATUS_OK)
	HANDLE_EXCEPTIONS_END
}

// Determines whether the cursor is in a field. Returns the a field struct if
// it is, or NULL if it is not.
statusCode cursorInField(document_t *doc, const char fieldType[],
						 field_t **returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->lock lock];
	
	WordSelectionObject* sbSelection = [doc->sbApp selection];
	CHECK_STATUS_LOCKED(doc)
	
	ENSURE_OK_LOCKED(doc, restoreCursor(doc))
	
	if(strcmp(fieldType, "Field") == 0) {
		NSInteger fieldCount;
		
		NSArray* sbFields = [[sbSelection fields] get];
		if(errorHasOccurred()) {
			clearError();
			fieldCount = 0;
		} else {
			fieldCount = [sbFields count];
		}
		
		if(fieldCount) {
			statusCode status = initField(doc, (WordField*) [sbFields
															objectAtIndex:0],
										 -1,  -1, NO, returnValue);
			RETURN_STATUS_LOCKED(doc, status);
		}
	
		sbFields = [[[[[sbSelection paragraphs] objectAtIndex:0] textObject]
				  fields] get];
		CHECK_STATUS_LOCKED(doc)
		if([sbFields count]) {
			// Check if fields are in the selection
			NSInteger selectionStart = [sbSelection selectionStart];
			CHECK_STATUS_LOCKED(doc)
			NSInteger selectionEnd = [sbSelection selectionEnd];
			CHECK_STATUS_LOCKED(doc)
			for(WordField* sbTestField in sbFields) {
				NSInteger fieldStart = [[sbTestField resultRange]
										startOfContent];
				CHECK_STATUS_LOCKED(doc)
				NSInteger fieldEnd = [[sbTestField resultRange] endOfContent];
				CHECK_STATUS_LOCKED(doc)
				
				// Check whether selection intersects with the field.
				if((selectionStart <= fieldStart
					&& selectionEnd >= fieldStart)
				   || (selectionStart <= fieldEnd
					   && selectionEnd >= fieldEnd)
				   || (selectionStart >= fieldStart
					   && selectionStart <= fieldEnd)) {
					   // Field in selection
					   statusCode status = initField(doc, sbTestField, -1, -1,
													 NO, returnValue);
					   if(status || *returnValue) RETURN_STATUS_LOCKED(doc,
																	   status)
				   } else if(fieldStart > selectionEnd) {
					   // Cursor already past end of selection
					   break;
				   }
			}
		}
	} else if(strcmp(fieldType, "Bookmark") == 0) {
		SBElementArray *sbBookmarks = [sbSelection bookmarks];
		CHECK_STATUS_LOCKED(doc)
		
		if(![doc->sbDoc isEqual:[NSNull null]]) {
			for(WordBookmark* sbBookmark in sbBookmarks) {
				statusCode status = initBookmark(doc, sbBookmark, -1, nil, NO,
												 returnValue);
				if(status || *returnValue) RETURN_STATUS_LOCKED(doc, status);
			}
			CHECK_STATUS_LOCKED(doc)
		}
	}
	
	*returnValue = NULL;
	RETURN_STATUS_LOCKED(doc, STATUS_OK)
	HANDLE_EXCEPTIONS_END
}

// Gets document data
statusCode getDocumentData(document_t *doc, char **returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->lock lock];
	
	WordTextRange *range = [doc->sbDoc createRangeStart:0 end:[EXPORTED_DOCUMENT_MARKER[0] length]];
	NSString *text = [range content];
	if (errorHasOccurred()) {
		clearError();
	}
	else if (text != nil) {
		for (unsigned short i = 0; EXPORTED_DOCUMENT_MARKER[i] != NULL; i++) {
			if ([text rangeOfString:EXPORTED_DOCUMENT_MARKER[i]].location == 0) {
				*returnValue = copyNSString(EXPORTED_DOCUMENT_MARKER[0]);
				RETURN_STATUS_LOCKED(doc, STATUS_OK);
			}
		}
	}
	
	NSString* returnString;
	statusCode status = getProperty(doc, PREFS_PROPERTY, &returnString);
	ENSURE_OK_LOCKED(doc, status)
	*returnValue = copyNSString(returnString);
	
	RETURN_STATUS_LOCKED(doc, STATUS_OK)
	HANDLE_EXCEPTIONS_END
}

// Sets document data
statusCode setDocumentData(document_t *doc, const char documentData[]) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->lock lock];
	
	NSString* propertyValue = [NSString stringWithUTF8String:documentData];
	RETURN_STATUS_LOCKED(doc, (setProperty(doc, PREFS_PROPERTY, propertyValue)));
	HANDLE_EXCEPTIONS_END
}

// Makes a field at the selection.
statusCode insertField(document_t *doc, const char fieldType[],
					   unsigned short noteType, field_t **returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->lock lock];
	
	WordTextRange* sbWhere = [[doc->sbApp selection] textObject];
	CHECK_STATUS_LOCKED(doc)
	
	WordWdStoryType storyType = [sbWhere storyType];
	if (storyType == WordWdStoryTypeFootnotesStory) {
		doc->insertTextIntoNote = NOTE_FOOTNOTE;
	}
	else if (storyType == WordWdStoryTypeEndnotesStory) {
		doc->insertTextIntoNote = NOTE_ENDNOTE;
	}
	
	if(strcmp(fieldType, "Bookmark") == 0 && !noteType) {
		[[doc->sbApp selection] setContent:@ FIELD_PLACEHOLDER];
		CHECK_STATUS_LOCKED(doc)
	}
	
	statusCode status = insertFieldRaw(doc, fieldType, noteType, sbWhere,
									   nil, returnValue);
	if(!status && strcmp(fieldType, "Bookmark") != 0) {
		ENSURE_OK_LOCKED(doc, setText(*returnValue, FIELD_PLACEHOLDER, false));
	}
	storeCursorLocation(doc);
	RETURN_STATUS_LOCKED(doc, status)
	HANDLE_EXCEPTIONS_END
}

statusCode setFieldAdjacency(field_t *fieldA, field_t *fieldB) {
	HANDLE_EXCEPTIONS_BEGIN
	// Mac Word appears to maintain some fake 1-char wide character between field codes in Word.
	fieldA->adjacent = fieldA->noteType == fieldB->noteType && [fieldA->sbContentRange endOfContent] + 2 - [fieldB->sbCodeRange startOfContent] >= 0;
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Gets fields
statusCode getFields(document_t *doc, const char fieldType[],
					 listNode_t** returnNode) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->lock lock];
	
	listNode_t* fieldListStart = NULL;
	listNode_t* fieldListEnd = NULL;
	
	if(strcmp(fieldType, "Field") == 0) {
		// Get fields in main body, footnotes, and endnotes
		NSArray* fieldCollections[3];
		statusCode status = getFieldCollections(doc, fieldCollections);
		ENSURE_OK_LOCKED(doc, status)
		
		long currentFieldIndices[] = {0, 0, 0};
		field_t* currentFields[] = {NULL, NULL, NULL};
		
		// Get numbers of fields and first fields
		unsigned long fieldCollectionSizes[3];
		for(unsigned short noteType=0; noteType<3; noteType++) {
			if(noteType == 0 || doc->wordVersion == 2004) {
				fieldCollectionSizes[noteType] = [fieldCollections[noteType]
												  count];
			} else {
				// This is necessary because the count selector always returns
				// 0 for Word 2008
				SBElementArray* sbFieldCollection = (SBElementArray*)
					fieldCollections[noteType];
				fieldCollectionSizes[noteType] =
					getEntryIndex(doc, [sbFieldCollection objectAtLocation:
										[NSNumber numberWithInt:-1]]);
			}
			
			if(errorHasOccurred()) {
				// There aren't actually any fields
				clearError();
				fieldCollectionSizes[noteType] = 0;
			} else {
				// Get first fields
				while(currentFields[noteType] == NULL &&
					  currentFieldIndices[noteType] <
					  fieldCollectionSizes[noteType]) {
					
					statusCode status = initField(doc, [fieldCollections[noteType]
						objectAtIndex:(currentFieldIndices[noteType])], noteType,
						currentFieldIndices[noteType]+1, NO,
												  &currentFields[noteType]);
					ENSURE_OK_LOCKED(doc, status)
					currentFieldIndices[noteType]++;
				}
			}
		}
		
		while(true) {
			BOOL isNextField = NO;
			
			// Figure out which node should come next
			unsigned short noteTypeA;
			for(noteTypeA=0; noteTypeA<3; noteTypeA++) {
				// Check if there is a field.
				if(!currentFields[noteTypeA]) continue;
				
				// We have a field. For now, assume this will be the next field.
				isNextField = YES;
				
				for(unsigned short noteTypeB = 0; noteTypeB<3; noteTypeB++) {
					if(noteTypeB == noteTypeA || !currentFields[noteTypeB]) {
						continue;
					}
					
					// If there is another field before this one, then obviously
					// this is not the next field.
					short returnValue;
					statusCode status = compareFields(currentFields[noteTypeA],
													  currentFields[noteTypeB],
													  &returnValue);
					ENSURE_OK_LOCKED(doc, status)
					
					if(returnValue > 0) {
						isNextField = NO;
						break;
					}
				}
				
				if(isNextField) {
					// Add as next field
					break;
				}
			}
			if(!isNextField) break;
			
			if (fieldListEnd) {
				ENSURE_OK_LOCKED(doc, setFieldAdjacency((field_t *) fieldListEnd->value, currentFields[noteTypeA]));
			}
			
			// Add node to linked list
			addValueToList(currentFields[noteTypeA], &fieldListStart,
						   &fieldListEnd);
			
			// Get next node of type
			currentFields[noteTypeA] = NULL;
			while(currentFields[noteTypeA] == NULL &&
				  currentFieldIndices[noteTypeA] <
				  fieldCollectionSizes[noteTypeA]) {
				
				statusCode status = initField(doc,
											  [fieldCollections[noteTypeA]
					objectAtIndex:(currentFieldIndices[noteTypeA])], noteTypeA,
					currentFieldIndices[noteTypeA]+1, NO,
					&currentFields[noteTypeA]);
				ENSURE_OK_LOCKED(doc, status)
				currentFieldIndices[noteTypeA]++;
			}
		}
	} else if(strcmp(fieldType, "Bookmark") == 0) {
		[doc->sbView setShowBookmarks:YES];
		SBElementArray *sbBookmarks = [doc->sbDoc bookmarks];
		CHECK_STATUS_LOCKED(doc)
		// We are going to allocate enough memory for all of the bookmarks to be
		// fields. This may not end up being the case, but we need arrays and
		// not linked lists to be able to use the 
		field_t** fields = (field_t**) malloc(sizeof(field_t*) * 
											  [sbBookmarks count]);
		CHECK_STATUS_LOCKED(doc)
		unsigned long nFields = 0;
		
		// Generate field structures
		for(WordBookmark* sbBookmark in sbBookmarks) {
			statusCode status = initBookmark(doc, sbBookmark, -1, nil, NO,
											 &fields[nFields]);
			ENSURE_OK_LOCKED(doc, status)
			if(fields[nFields]) nFields++;
		}
		
		// Sort
		statusCode status = STATUS_OK;
		qsort_r(fields, nFields, sizeof(field_t *), &status,
				&compareFieldsQsort);
		CHECK_STATUS_LOCKED(doc)
		
		// Generate the linked list
		for(unsigned long i=0; i<nFields; i++) {
			addValueToList(fields[i], &fieldListStart, &fieldListEnd);
		}
	} else {
		DIE(([NSString stringWithFormat:@"Unknown field type \"%s\"",
			fieldType]))
	}
	
	if(fieldListStart) addValueToList(fieldListStart,
									  &(doc->allocatedFieldListsStart),
									  &(doc->allocatedFieldListsEnd));
	
	*returnNode = fieldListStart;
	RETURN_STATUS_LOCKED(doc, STATUS_OK)
	HANDLE_EXCEPTIONS_END
}

statusCode convert(document_t *doc, field_t* fields[], unsigned long nFields,
				   const char toFieldType[], unsigned short toNoteTypes[]) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->lock lock];
	
	statusCode status = prepareReadFieldCode(doc);
	ENSURE_OK_LOCKED(doc, status)
	
	// Get field collections, keep track of offsets
	NSArray* fieldCollections[3];
	status = getFieldCollections(doc, fieldCollections);
	ENSURE_OK_LOCKED(doc, status)
	NSInteger offsets[] = {0, 0, 0};
	
	// Make sure text locations are set
	for(long i=0; i<nFields; i++) {
		ensureTextLocationSet(fields[i]);
	}
	
	// Get to field type
	BOOL toField = NO;
	BOOL toBookmark = NO;
	if(strcmp(toFieldType, "Field") == 0) {
		toField = YES;
	} else if(strcmp(toFieldType, "Bookmark") == 0) {
		toBookmark = YES;
	} else {
		DIE(([NSString stringWithFormat:@"Unknown field type \"%s\"",
			  toFieldType]))
	}
	
	// Loop through fields in reverse order
	for(unsigned long i = 0; i<nFields; i++) {
		field_t* field = fields[i];
		
		unsigned short toNoteType = toNoteTypes[i];
		
		if(field->sbField) {
			[field->sbField release];
			[field->sbContentRange release];
			
			field->sbField = [fieldCollections[field->noteType] objectAtIndex:
							  (field->entryIndex+offsets[field->noteType]-1)];
			CHECK_STATUS_LOCKED(doc)
			field->sbContentRange = [field->sbField resultRange];
			CHECK_STATUS_LOCKED(doc)
			
			[field->sbField retain];
			[field->sbContentRange retain];
		}
		
		// First, handle converting between footnote, endnote, and inline
		WordFootnote* sbNote;
		if(field->noteType == NOTE_FOOTNOTE) {
			sbNote = [[field->sbContentRange footnotes] objectAtIndex:0];
			CHECK_STATUS_LOCKED(doc)
		} else if(field->noteType == NOTE_ENDNOTE) {
			sbNote = [[field->sbContentRange endnotes] objectAtIndex:0];
			CHECK_STATUS_LOCKED(doc)
		} else {
			sbNote = nil;
		}
		
		// Word has special facilities for converting footnotes/endnotes
		if((toNoteType == NOTE_ENDNOTE && field->noteType == NOTE_FOOTNOTE) ||
		   (toNoteType == NOTE_FOOTNOTE && field->noteType == NOTE_ENDNOTE)) {
			// Check which fields are in this note
			unsigned long nFieldsInNote = 0;
			while(nFieldsInNote < nFields-i &&
				  fields[nFieldsInNote+i]->textLocation
				  == field->textLocation) {
				nFieldsInNote++;
			}
			
			// Keep data needed to re-create bookmarks, because the conversion
			// will delete them
			NSInteger* bookmarkStarts;
			NSInteger* bookmarkEnds;
			NSInteger oldNoteStart;
			if(field->sbBookmark) {
				oldNoteStart = [[sbNote textObject] startOfContent];
				CHECK_STATUS_LOCKED(doc)
				bookmarkStarts = (NSInteger*) malloc(sizeof(NSInteger)
													 * nFieldsInNote);
				bookmarkEnds = (NSInteger*) malloc(sizeof(NSInteger)
													 * nFieldsInNote);
				for(unsigned long j=0; j<nFieldsInNote; j++) {
					bookmarkStarts[j] = [fields[i+j]->sbBookmark
										 startOfBookmark];
					CHECK_STATUS_LOCKED(doc)
					bookmarkEnds[j] = [fields[i+j]->sbBookmark endOfBookmark];
					CHECK_STATUS_LOCKED(doc)
				}
			}
			
			// Convert fields and get the mark
			WordFootnote *sbNewNote;
			status = noteSwap(doc, sbNote, field->noteType, &sbNewNote);
			ENSURE_OK_LOCKED(doc, status)
			sbNote = sbNewNote;
			
			// Find the fields again
			if(field->sbField) {
				offsets[field->noteType] -= nFieldsInNote;
				offsets[toNoteType] += nFieldsInNote;
				
				if(toField) {
					// We don't need to convert the other entries in this note.
					i += nFieldsInNote-1;
					continue;
				}  else {
					// Need to update conversions
					[field->sbField release];
					[field->sbContentRange release];
					
					field->sbField = [[[sbNote textObject] fields]
									  objectAtIndex:0];
					CHECK_STATUS_LOCKED(doc)
					field->sbContentRange = [field->sbField resultRange];
					CHECK_STATUS_LOCKED(doc)
					NSInteger entryIndex = getEntryIndex(doc, field->sbField);
					CHECK_STATUS_LOCKED(doc)
					
					[field->sbField retain];
					[field->sbContentRange retain];
					
					NSUInteger oldEntryIndex = field->entryIndex;
					for(unsigned long j=0; j<nFieldsInNote; j++) {
						fields[i+j]->entryIndex = entryIndex
							+ (fields[i+j]->entryIndex - oldEntryIndex)
							- offsets[toNoteType];
						CHECK_STATUS_LOCKED(doc)
						fields[i+j]->noteType = toNoteType;
						CHECK_STATUS_LOCKED(doc)
					}
				}
			} else if(field->sbBookmark) {
				WordTextRange* sbNoteTextObject = [sbNote textObject];
				CHECK_STATUS_LOCKED(doc)
				NSInteger startDifference = [sbNoteTextObject startOfContent]
				- oldNoteStart;
				CHECK_STATUS_LOCKED(doc)
				for(unsigned long j=0; j<nFieldsInNote; j++) {
					fields[i+j]->noteType = toNoteType;
					status = insertFieldRaw(doc, "Bookmark", 0,
											sbNoteTextObject,
											fields[i+j]->bookmarkNameNS, NULL);
					ENSURE_OK_LOCKED(doc, status)
					[fields[i+j]->sbBookmark
					 setStartOfBookmark:bookmarkStarts[j]+startDifference];
					CHECK_STATUS_LOCKED(doc)
					[fields[i+j]->sbBookmark
					 setEndOfBookmark:bookmarkEnds[j]+startDifference];
					CHECK_STATUS_LOCKED(doc)
				}
			}
		}
		
		// Check whether the field is not the entire content of the note. If
		// this is the case, then we don't need to convert fields to notes
		BOOL inlineNoteField = sbNote && [[[sbNote textObject] content]
										   isNotEqualTo:[field->sbContentRange
														 content]];
		CHECK_STATUS_LOCKED(doc)
		
		// Skip further processing if it's unnecessary
		if(((field->sbField && toField) || (field->sbBookmark && toBookmark))
		   && (field->noteType == toNoteType || inlineNoteField)) {
			continue;
		}
		
		// Convert between field types and between inline and footnote/endnote
		// citations
		if(field->sbField) {
			field_t* newField;
			if(!(field->noteType)) {
				// Convert an in-text citation to note or bookmark
				NSInteger fieldStart = [[field->sbField fieldCode]
										startOfContent];
				CHECK_STATUS_LOCKED(doc)
				if (strcmp(toFieldType, "Bookmark") == 0) {
					[field->sbField delete];
					CHECK_STATUS_LOCKED(doc)
				}
				WordTextRange* sbWhere = [doc->sbDoc
										  createRangeStart:fieldStart-1
										  end:fieldStart-1];
				CHECK_STATUS_LOCKED(doc)
				status = insertFieldRaw(doc, toFieldType, toNoteType, sbWhere,
										nil, &newField);
				ENSURE_OK_LOCKED(doc, status)
				if (toBookmark) {
					ENSURE_OK_LOCKED(doc, setText(newField, FIELD_PLACEHOLDER, false));
				}
				else {
					// Copy text for in-text to note conversions
					[newField->sbContentRange setContent:[field->sbContentRange content]];
					CHECK_STATUS_LOCKED(doc)
					[field->sbField delete];
					CHECK_STATUS_LOCKED(doc)
				}
			} else if(field->noteType == toNoteType || inlineNoteField) {
				// Convert fields inside a note to bookmarks
				status = insertFieldRaw(doc, toFieldType, 0,
										field->sbContentRange, nil, &newField);
				ENSURE_OK_LOCKED(doc, status)
				[newField->sbBookmark
				 setStartOfBookmark:([newField->sbBookmark endOfBookmark]+1)];
				CHECK_STATUS_LOCKED(doc)
				[field->sbField delete];
				CHECK_STATUS_LOCKED(doc)
			} else {
				// Convert note to inline
				status = insertFieldRaw(doc, toFieldType, toNoteType,
										[sbNote noteReference], nil,
										&newField);
				ENSURE_OK_LOCKED(doc, status)
				ENSURE_OK_LOCKED(doc, setText(newField, FIELD_PLACEHOLDER, false));
				CHECK_STATUS_LOCKED(doc)
				// Delete old note
				[[sbNote noteReference] setContent:@""];
			}
			
			offsets[field->noteType]--;
			setCode(newField, field->code);
		} else {
			// Find appropriate range to overwrite depending on type of citation
			// we're converting from
			WordTextRange* sbWhere;
			if(field->noteType && !toNoteType && !inlineNoteField) {
				// Convert note to in-text
				// Delete old footnote reference and put range where it was
				NSInteger referenceStart = [[sbNote noteReference]
											startOfContent];
				CHECK_STATUS_LOCKED(doc)
				[[sbNote noteReference] setContent:@""];
				CHECK_STATUS_LOCKED(doc)
				sbWhere = [doc->sbDoc createRangeStart:referenceStart
												   end:referenceStart];
				CHECK_STATUS_LOCKED(doc)
			} else if(!(field->noteType) && toNoteType && !inlineNoteField) {
				// Convert in-text to note
				// Delete old in-text citation and put footnote reference
				// where it was
				NSInteger bookmarkStart = [field->sbBookmark startOfBookmark];
				CHECK_STATUS_LOCKED(doc)
				[field->sbContentRange setContent:@""];
				CHECK_STATUS_LOCKED(doc)
				sbWhere = [doc->sbDoc createRangeStart:bookmarkStart
												   end:bookmarkStart];
				CHECK_STATUS_LOCKED(doc)
			} else {
				// Convert in-text to in-text
				sbWhere = field->sbContentRange;
			}
			
			// Create bookmark or field
			if(toField) {
				field_t* newField;
				status = insertFieldRaw(doc, toFieldType, toNoteType, sbWhere,
										nil, &newField);
				ENSURE_OK_LOCKED(doc, status)
				setCode(newField, field->code);
				
				// Word 2004: If the bookmark still exists, delete it
				IGNORING_SB_ERRORS_BEGIN
				[field->sbContentRange setContent:@""];
				IGNORING_SB_ERRORS_END
			} else {
				insertFieldRaw(doc, toFieldType, toNoteType, sbWhere,
							   field->bookmarkNameNS, NULL);
			}
		}
		
		// Increment offsets appropriately
		if(toField) {
			offsets[toNoteType]++;
		}
	};
	
	RETURN_STATUS_LOCKED(doc, STATUS_OK)
	HANDLE_EXCEPTIONS_END
}

// Sets the "Bibliography" style
statusCode setBibliographyStyle(document_t* doc, long firstLineIndent, 
								long bodyIndent, unsigned long lineSpacing,
								unsigned long entrySpacing, long tabStops[],
								unsigned long tabStopCount) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->lock lock];
	
	WordWordStyle* sbBibliographyStyle;
	NSString* sanityCheck = nil;
	
	sbBibliographyStyle = [[doc->sbDoc WordStyles]
						   objectWithID:@(WordWdBuiltinStyleStyleBibliography)];
	if(!errorHasOccurred()) {
		sanityCheck = [sbBibliographyStyle nameLocal];
		if(sanityCheck && !errorHasOccurred()) {
			SBElementArray* sbTabStops = [[sbBibliographyStyle paragraphFormat]
										  tabStops];
			[sbTabStops removeAllObjects];
		}
	}
	
	if(!sanityCheck || errorHasOccurred()) {
		clearError();
		
		// tell application "Microsoft Word" to make new Word style at %@
		// with properties {name local:"Bibliography",
		// style type:style type paragraph, base style:style normal}
		NSAppleEventDescriptor* reco = [NSAppleEventDescriptor recordDescriptor];
		[reco setDescriptor:[NSAppleEventDescriptor descriptorWithString:
							 @"Bibliography"]
				 forKeyword:'2499'];
		[reco setDescriptor:[NSAppleEventDescriptor
							 descriptorWithEnumCode:WordWdStyleTypeStyleTypeParagraph]
				 forKeyword:'2504'];
		[reco setDescriptor:[NSAppleEventDescriptor
							 descriptorWithEnumCode:WordWdBuiltinStyleStyleNormal]
				 forKeyword:'2501'];
		sbBibliographyStyle = [doc->sbApp sendEvent:'core' id:'crel'
										 parameters:'kocl',
		 [NSAppleEventDescriptor descriptorWithTypeCode:'w173'], 'insh',
							   doc->sbDoc, 'prdt', reco, nil];
		CHECK_STATUS_LOCKED(doc)
	}
	
	WordParagraphFormat* sbParagraphFormat = [sbBibliographyStyle
											  paragraphFormat];
	CHECK_STATUS_LOCKED(doc)
	[sbParagraphFormat setFirstLineIndent:((double) firstLineIndent)/20];
	CHECK_STATUS_LOCKED(doc)
	[sbParagraphFormat setParagraphFormatLeftIndent:((double) bodyIndent)/20];
	CHECK_STATUS_LOCKED(doc)
	[sbParagraphFormat setLineSpacing:((double) lineSpacing)/20];
	CHECK_STATUS_LOCKED(doc)
	[sbParagraphFormat setSpaceAfter:((double) entrySpacing)/20];
	CHECK_STATUS_LOCKED(doc)
	
	for(unsigned long i=0; i<tabStopCount; i++) {
		// tell application "Microsoft Word" to make new tab stop at %@
		// with properties {alignment:align tab left,
		// tab leader: tab leader spaces, tab stop position:%@}
		NSAppleEventDescriptor* reco = [NSAppleEventDescriptor recordDescriptor];
		[reco setDescriptor:[NSAppleEventDescriptor
							 descriptorWithEnumCode:WordWdTabAlignmentAlignTabLeft]
				 forKeyword:'1918'];
		[reco setDescriptor:[NSAppleEventDescriptor
							 descriptorWithEnumCode:WordWdTabLeaderTabLeaderSpaces]
				 forKeyword:'2007'];
		[reco setDescriptor:[NSAppleEventDescriptor descriptorWithInt32:
							 ((int32_t) (((double) tabStops[i])/20))]
				 forKeyword:'2009'];
		
		[doc->sbApp sendEvent:'core' id:'crel' parameters:'kocl',
		 [NSAppleEventDescriptor descriptorWithTypeCode:'w135'], 'insh',
		 [sbBibliographyStyle paragraphFormat], 'prdt', reco, nil];
	}
	
	RETURN_STATUS_LOCKED(doc, STATUS_OK)
	HANDLE_EXCEPTIONS_END
}

statusCode exportDocument(document_t *doc,
						  const char fieldType[],
						  const char importInstructions[]) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->lock lock];
	
	// Export fields
	listNode_t *fields;
	ENSURE_OK(getFields(doc, fieldType, &fields));
	while (fields != NULL) {
		field_t *field = (field_t *)fields->value;
		// If you insert in sbField.resultRange it doesn't remove the field code
		// and upon unlinking removes the hyperlink too.
		// sbField.resultRange depends on the existance of sbField, so if we in any way unlink
		// the field the resultRange stops working too.
		// Inserting at char 1 replaces the field. This is the only way to make this code work.
		WordTextRange *insertRange = [[[field->sbField resultRange] characters] objectAtIndex:0];
		CHECK_STATUS_LOCKED(doc);
		[doc->sbApp createNewFieldTextRange:insertRange fieldType:WordWdFieldTypeFieldHyperlink fieldText:IMPORT_LINK_URL preserveFormatting:NO];
		CHECK_STATUS_LOCKED(doc);
		[[field->sbField resultRange] setContent:[NSString stringWithUTF8String:field->code]];
		CHECK_STATUS_LOCKED(doc);
		fields = fields->next;
	}

	// Document data
	char *docData;
	ENSURE_OK_LOCKED(doc, getDocumentData(doc, &docData));
	NSString *importDocData = IMPORT_DOC_PREFS_PREFIX;
	importDocData = [importDocData stringByAppendingString:[NSString stringWithUTF8String:docData]];
	free(docData);
	WordTextRange *range = [doc->sbDoc textObject];
	range = [range collapseRangeDirection:WordWdCollapseDirectionCollapseEnd];
	[doc->sbApp insertParagraphAt:range];
	range = [[doc->sbDoc textObject] collapseRangeDirection:WordWdCollapseDirectionCollapseEnd];
	[doc->sbApp createNewFieldTextRange:range fieldType:WordWdFieldTypeFieldHyperlink fieldText:IMPORT_LINK_URL preserveFormatting:NO];
	WordField *insertedHyperlink = [[doc->sbDoc fields] lastObject];
	[[insertedHyperlink resultRange] setContent:importDocData];
	CHECK_STATUS_LOCKED(doc)

	// Import instructions
	range = [doc->sbDoc textObject];
	range = [range collapseRangeDirection:WordWdCollapseDirectionCollapseStart];
	[doc->sbApp insertParagraphAt:range];
	range = [range collapseRangeDirection:WordWdCollapseDirectionCollapseStart];
	[doc->sbApp insertParagraphAt:range];
	[doc->sbApp insertText:[NSString stringWithUTF8String:importInstructions] at:range];
	// Export marker
	range = [range collapseRangeDirection:WordWdCollapseDirectionCollapseStart];
	[doc->sbApp insertParagraphAt:range];
	range = [range collapseRangeDirection:WordWdCollapseDirectionCollapseStart];
	[doc->sbApp insertParagraphAt:range];
	[doc->sbApp insertText:EXPORTED_DOCUMENT_MARKER[0] at:range];
	CHECK_STATUS_LOCKED(doc)
	
	// Don't attempt to restore cursor location. It doesn't work since the document size changes
	// and it doesn't make sense either.
	doc->shouldRestoreCursor = NO;
	
	RETURN_STATUS_LOCKED(doc, STATUS_OK);
	HANDLE_EXCEPTIONS_END
}

statusCode importDocument(document_t *doc, const char fieldType[], bool *returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->lock lock];
	*returnValue = false;
	
	NSArray* fieldCollections[3];
	statusCode status = getFieldCollections(doc, fieldCollections);
	ENSURE_OK_LOCKED(doc, status)
	
	// Get numbers of links
	for (unsigned short noteType = 0; noteType < 3; noteType++) {
		unsigned long fieldCollectionSize;
		if (noteType == 0) {
			fieldCollectionSize = [fieldCollections[noteType]
											  count];
		} else {
			// This is necessary because the count selector always returns
			// 0 for Word 2008
			SBElementArray* sbfieldCollection = (SBElementArray *)	fieldCollections[noteType];
			fieldCollectionSize =
				getEntryIndex(doc, [sbfieldCollection objectAtLocation:
								[NSNumber numberWithInt:-1]]);
		}
		
		if (errorHasOccurred()) {
			// There aren't any links
			clearError();
			fieldCollectionSize = 0;
			continue;
		}
		
		for (long i = fieldCollectionSize-1; i >= 0; i--) {
			WordField *link = [fieldCollections[noteType] objectAtIndex:(i)];
			CHECK_STATUS_LOCKED(doc);
			WordWdFieldType sbFieldType = [link fieldType];
			if (sbFieldType != WordWdFieldTypeFieldHyperlink) {
				continue;
			}
			CHECK_STATUS_LOCKED(doc);
			WordTextRange *linkRange = [link resultRange];
			CHECK_STATUS_LOCKED(doc);
			NSString *linkText = [linkRange content];
			CHECK_STATUS_LOCKED(doc);
			if ([linkText rangeOfString:IMPORT_ITEM_PREFIX].location == 0
					|| [linkText rangeOfString:IMPORT_BIBL_PREFIX].location == 0) {
				NSString* rawCode = [NSString stringWithFormat:@"%@%@ ",
									 MAIN_FIELD_PREFIX,
									 linkText];
				[[link fieldCode] setContent:rawCode];
				[linkRange setStyle:WordWdBuiltinStyleStyleDefaultParagraphFont];
				CHECK_STATUS_LOCKED(doc);
				if (noteType == 1) {
					[linkRange setStyle:WordWdBuiltinStyleStyleFootnoteText];
					CHECK_STATUS_LOCKED(doc);
				}
				else if (noteType == 2) {
					[linkRange setStyle:WordWdBuiltinStyleStyleEndnoteText];
					CHECK_STATUS_LOCKED(doc);
				}
			}
			else if ([linkText rangeOfString:IMPORT_DOC_PREFS_PREFIX].location == 0) {
				*returnValue = true;
				const char* docPrefs = [[linkText substringFromIndex:
										 [IMPORT_DOC_PREFS_PREFIX length]] UTF8String];
				ENSURE_OK_LOCKED(doc, setDocumentData(doc, docPrefs));
				[link delete];
				CHECK_STATUS_LOCKED(doc);
			}
		}
	}
	
	// Remove 3 paragraphs: export marker, empty paragraph, import instructions and another empty paragraph
	WordTextRange *range = [doc->sbDoc textObject];
	range = [range collapseRangeDirection:WordWdCollapseDirectionCollapseStart];
	range = [range moveEndOfRangeBy:WordWdUnitsAParagraphItem count:4];
	[range setContent:@""];
	CHECK_STATUS_LOCKED(doc);
	
	// Don't attempt to restore cursor location. It doesn't work since the document size changes
	// and it doesn't make sense either.
	doc->shouldRestoreCursor = NO;
	
	RETURN_STATUS_LOCKED(doc, STATUS_OK);
	HANDLE_EXCEPTIONS_END
}

statusCode insertText(document_t *doc, const char htmlString[]) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->lock lock];
	
	if (doc->insertTextIntoNote != 0) {
		// Create new note
		NSAppleEventDescriptor* noteTypeCode;
		if (doc->insertTextIntoNote == NOTE_FOOTNOTE) {
			noteTypeCode = [NSAppleEventDescriptor
							descriptorWithTypeCode:'w156'];
		} else {
			noteTypeCode = [NSAppleEventDescriptor
							descriptorWithTypeCode:'w157'];
		}
		// make new footnote/endnote at sbDoc
		WordFootnote *sbNote = [doc->sbApp sendEvent:'core' id:'crel' parameters:'kocl',
				  noteTypeCode, 'insh', doc->sbDoc, nil];
		CHECK_STATUS_LOCKED(doc)
		[[sbNote textObject] sendEvent:'misc' id:'slct' parameters:'\00\00\00\00', nil];
		CHECK_STATUS_LOCKED(doc)
	}
	
	// Insert empty paragraph before rich text since inserting it messes up footnote formatting
	WordTextRange* selectionRange = [[doc->sbApp selection] textObject];
	CHECK_STATUS_LOCKED(doc)
	[[doc->sbApp selection] setContent:@"\n "];
	CHECK_STATUS_LOCKED(doc)
	
	IGNORING_SB_ERRORS_BEGIN
	// Make sure temp bookmark is gone
	[[[doc->sbDoc bookmarks] objectWithName:@ RTF_TEMP_BOOKMARK]
	 delete];
	IGNORING_SB_ERRORS_END
	IGNORING_SB_ERRORS_BEGIN
	// Save properties
	WordFont* font = [selectionRange fontObject];
	NSString* oldFontName = [font name];
	NSString* oldFontOtherName = [font otherName];
	IGNORING_SB_ERRORS_END
	
	// Insert a temp bookmark into which we'll insert the HTML
	// We need to insert into a bookmark so that we have a text range
	// which corresponds to the inserted HTML so we can change fonts, etc.
	ENSURE_OK_LOCKED(doc, insertFieldRaw(doc, "Bookmark", 0,
									   selectionRange,
									   @ RTF_TEMP_BOOKMARK, nil))
	WordBookmark *tempBookmark = [[doc->sbDoc bookmarks]
								  objectWithName:@ RTF_TEMP_BOOKMARK];
	CHECK_STATUS_LOCKED(doc)
	WordTextRange *bookmarkRange = [tempBookmark textObject];
	CHECK_STATUS_LOCKED(doc)
	[[doc->sbApp selection] setSelectionEnd: [[doc->sbApp selection] selectionEnd]-1];
	CHECK_STATUS_LOCKED(doc)
	
	// Write HTML to a clipboard
	storePasteboardItems();
	replacePasteboardContentsWithHTML(htmlString);
	
	// Paste HTML
	[selectionRange pasteObject];
	// Restore clipboard contents and only then check for errors
	restorePasteboardContents();
	CHECK_STATUS_LOCKED(doc)
	
	// Set style
	IGNORING_SB_ERRORS_BEGIN
	// Set properties back to saved
	font = [bookmarkRange fontObject];
	[font setName:oldFontName];
	[font setOtherName:oldFontOtherName];
	IGNORING_SB_ERRORS_END
	
	// Remove the previously added empty paragraph
	[bookmarkRange sendEvent:'misc' id:'slct' parameters:'\00\00\00\00', nil];
	[[doc->sbApp selection] setSelectionStart: [[doc->sbApp selection] selectionEnd]-1];
	[[doc->sbApp selection] setContent:@""];
	CHECK_STATUS_LOCKED(doc)
	
	// Put the cursor at the end of inserted text
	[bookmarkRange sendEvent:'misc' id:'slct' parameters:'\00\00\00\00', nil];
	[[doc->sbApp selection] setSelectionStart: [[doc->sbApp selection] selectionEnd]-1];
	[[doc->sbApp selection] setContent:@""];
	CHECK_STATUS_LOCKED(doc)
	storeCursorLocation(doc);
	doc->cursorMoved = NO;
	
	// Remove the temp bookmark
	[[[doc->sbDoc bookmarks] objectWithName:@ RTF_TEMP_BOOKMARK]
	 delete];
	CHECK_STATUS_LOCKED(doc)
	
	RETURN_STATUS_LOCKED(doc, STATUS_OK)
	HANDLE_EXCEPTIONS_END
}

statusCode convertPlaceholdersToFields(document_t *doc, const char* placeholders[],
									   const unsigned long nPlaceholders, const unsigned short noteType,
									   const char fieldType[], listNode_t** returnNode) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->lock lock];
	
	listNode_t* fieldListStart = NULL;
	listNode_t* fieldListEnd = NULL;
	
	NSArray* fieldCollections[3];
	statusCode status = getFieldCollections(doc, fieldCollections);
	ENSURE_OK_LOCKED(doc, status)
	NSMutableArray *links = [NSMutableArray array];
	
	// Get numbers of links
	for (unsigned short noteType = 0; noteType < 3; noteType++) {
		unsigned long fieldCollectionSize;
		if (noteType == 0) {
			fieldCollectionSize = [fieldCollections[noteType]
								   count];
		} else {
			// This is necessary because the count selector always returns
			// 0 for Word 2008
			SBElementArray* sbfieldCollection = (SBElementArray *)	fieldCollections[noteType];
			fieldCollectionSize =
			getEntryIndex(doc, [sbfieldCollection objectAtLocation:
								[NSNumber numberWithInt:-1]]);
		}
		
		if (errorHasOccurred()) {
			// There aren't any links
			clearError();
			fieldCollectionSize = 0;
			continue;
		}
		
		for (long i = fieldCollectionSize-1; i >= 0; i--) {
			WordField *link = [fieldCollections[noteType] objectAtIndex:(i)];
			CHECK_STATUS_LOCKED(doc);
			WordWdFieldType sbFieldType = [link fieldType];
			if (sbFieldType != WordWdFieldTypeFieldHyperlink) {
				continue;
			}
			CHECK_STATUS_LOCKED(doc);
			WordTextRange *linkRange = [link resultRange];
			CHECK_STATUS_LOCKED(doc);
			NSString *linkUrl = [[[linkRange hyperlinkObjects] objectAtIndex:0] hyperlinkAddress];
			if (linkUrl == nil) continue;
			CHECK_STATUS_LOCKED(doc);
			NSString *placeholder = [linkUrl substringFromIndex:[IMPORT_LINK_URL length]+1];
			for (long j = 0; j < nPlaceholders; j++) {
				if ([placeholder rangeOfString:[NSString stringWithUTF8String:placeholders[j]]].location != 0) {
					continue;
				}
				[links addObject:link];
				break;
			}
		}
	}
	NSArray *sortedLinks = [links sortedArrayUsingComparator:^NSComparisonResult(id a, id b) {
		NSString *aLinkUrl = [[[[(WordField *)a resultRange] hyperlinkObjects] objectAtIndex:0] hyperlinkAddress];
		NSString *bLinkUrl = [[[[(WordField *)a resultRange] hyperlinkObjects] objectAtIndex:0] hyperlinkAddress];
		NSString *aPlaceholder = [aLinkUrl substringFromIndex:[IMPORT_LINK_URL length]+1];
		NSString *bPlaceholder = [bLinkUrl substringFromIndex:[IMPORT_LINK_URL length]+1];
		for (long i = 0; i < nPlaceholders; i++) {
			NSString *placeholder = [NSString stringWithUTF8String:placeholders[i]];
			if ([aPlaceholder rangeOfString:placeholder].location == 0) return -1;
			else if ([bPlaceholder rangeOfString:placeholder].location == 0) return 1;
		}
		return 0;
	}];
	
	for (id object in sortedLinks) {
		WordField *link = object;
		field_t *newField;
		WordTextRange *linkRange = [link resultRange];
		CHECK_STATUS_LOCKED(doc)
		WordWdStoryType storyType = [linkRange storyType];
		CHECK_STATUS_LOCKED(doc)
		
		if ((noteType == 0 || storyType != WordWdStoryTypeMainTextStory) && strcmp(fieldType, "Field") == 0) {
			NSString* rawCode = [NSString stringWithFormat:@"%@%@ ",
								 MAIN_FIELD_PREFIX,
								 @"TEMP"];
			[[link fieldCode] setContent:rawCode];
			WordFont *font = [linkRange fontObject];
			[font setColorIndex:WordWdColorIndexAuto];
			[font setUnderline:WordWdUnderlineUnderlineNone];
			
			CHECK_STATUS_LOCKED(doc)
			ENSURE_OK_LOCKED(doc, initField(doc, link, -1, -1, false, &newField))
		}
		else {
			// The field code in MacWord is stored sequentially as `fieldCode` range followed by `resultRange` range
			// and we want a cursor just preceding the field here.
			WordTextRange *codeRange = [link fieldCode];
			CHECK_STATUS_LOCKED(doc)
			WordTextRange *insertRange = [doc->sbDoc createRangeStart:[codeRange startOfContent]-1 end:[codeRange startOfContent]-1];
			CHECK_STATUS_LOCKED(doc)
			[link delete];
			CHECK_STATUS_LOCKED(doc)
			ENSURE_OK_LOCKED(doc, insertFieldRaw(doc, fieldType, noteType, insertRange, nil, &newField));
			ENSURE_OK_LOCKED(doc, setCode(newField, "TEMP"))
		}
		addValueToList(newField, &fieldListStart, &fieldListEnd);
	}
	
	*returnNode = fieldListStart;
	
	RETURN_STATUS_LOCKED(doc, STATUS_OK)
	HANDLE_EXCEPTIONS_END
}

// Run on exit/dialog display in Zotero to restore document to previous state
statusCode cleanup(document_t *doc) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->lock lock];
	
	if(doc->restoreInsertionsAndDeletions
	   && !doc->statusInsertionsAndDeletions) {
		// See comment in prepareReadFieldCode() for explanation
		ENSURE_OK_LOCKED(doc, moveCursorOutOfNote(doc));
		[doc->sbView setShowInsertionsAndDeletions:YES];
		CHECK_STATUS_LOCKED(doc)
		doc->statusInsertionsAndDeletions = YES;
	}
	if(doc->restoreFormatChanges && !doc->statusFormatChanges) {
		ENSURE_OK_LOCKED(doc, moveCursorOutOfNote(doc));
		[doc->sbView setShowFormatChanges:YES];
		CHECK_STATUS_LOCKED(doc)
		doc->statusFormatChanges = YES;
	}
	
	RETURN_STATUS_LOCKED(doc, restoreCursor(doc))
	HANDLE_EXCEPTIONS_END
}

// Run on exit to clean up anything we played with
statusCode complete(document_t *doc) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->lock lock];
	
	if (doc->restoreTrackRevisions) {
		[doc->sbDoc setTrackRevisions:YES];
		CHECK_STATUS_LOCKED(doc);
	}
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Get NSArrays representing fields in different parts of the document.
statusCode getFieldCollections(document_t *doc, NSArray** fieldCollections) {
	fieldCollections[0] = [doc->sbDoc fields];
	CHECK_STATUS;
	
	if(doc->wordVersion == 2004) {
		ENSURE_OK(extractFieldsFromNotes([doc->sbDoc footnotes],
										 &fieldCollections[1]));
		ENSURE_OK(extractFieldsFromNotes([doc->sbDoc endnotes],
										 &fieldCollections[2]));
	} else {
		fieldCollections[1] = [[doc->sbDoc
								getStoryRangeStoryType:WordWdStoryTypeFootnotesStory]
							   fields];
		CHECK_STATUS;
		fieldCollections[2] = [[doc->sbDoc
								getStoryRangeStoryType:WordWdStoryTypeEndnotesStory]
							   fields];
		CHECK_STATUS;
	}
	
	return STATUS_OK;
}

// Get NSArrays representing fields from a given collection. Included only for 
// Word 2004
statusCode extractFieldsFromNotes(SBElementArray* sbNotes, 
								  NSArray** returnValue) {
	NSUInteger nNotes = [sbNotes count];
	CHECK_STATUS
	
	NSMutableArray* sbFields = [NSMutableArray arrayWithCapacity:nNotes];
	for(WordFootnote* sbNote in sbNotes) {
		NSArray* sbNoteFields = [[[sbNote textObject] fields] get];
		CHECK_STATUS
		
		if(!sbNoteFields) continue;
		
		[sbFields addObjectsFromArray:sbNoteFields];
	}
	
	*returnValue = sbFields;
	
	return STATUS_OK;
}

// Swap footnotes and endnotes
statusCode noteSwap(document_t *doc, WordFootnote* sbNote,
					unsigned short noteType,
					WordFootnote **returnValue) {
	if(doc->wordVersion == 2004) {
		// footnote_convert() and endnote_convert() crash Word 2004
		WordTextRange* sbReferenceRange = [[sbNote noteReference]
							moveEndOfRangeBy:WordWdUnitsACharacterItem count:0];
		CHECK_STATUS
		WordTextRange* sbCreateRange = [[sbNote noteReference] 
										collapseRangeDirection:WordWdCollapseDirectionCollapseEnd];
		CHECK_STATUS
		
		if(noteType == NOTE_FOOTNOTE) {
			// make new endnote at active document
			[doc->sbApp sendEvent:'core' id:'crel'
					   parameters:'kocl', [NSAppleEventDescriptor
										   descriptorWithTypeCode:'w157'],
			 'insh', sbCreateRange, nil];
			CHECK_STATUS
			*returnValue = [[doc->sbDoc endnotes]
							objectAtIndex:(getEntryIndex(doc,
														 [[sbReferenceRange
														   endnotes]
														  objectAtIndex:0])-1)];
			CHECK_STATUS
		} else if(noteType == NOTE_ENDNOTE) {
			// make new footnote at active document
			[doc->sbApp sendEvent:'core' id:'crel'
					   parameters:'kocl', [NSAppleEventDescriptor
										   descriptorWithTypeCode:'w156'],
			 'insh', sbCreateRange, nil];
			CHECK_STATUS
			*returnValue = [[doc->sbDoc footnotes]
							objectAtIndex:(getEntryIndex(doc,
														 [[sbReferenceRange
														   footnotes]
														  objectAtIndex:0])-1)];
			CHECK_STATUS
		} else {
			DIE(@"Invalid note type")
		}
		
		// Transfer text content
		[[*returnValue textObject] 
		 setFormattedText:[[sbNote textObject] formattedText]];
		CHECK_STATUS
		
		// Delete old note
		[sbNote delete];
	} else {
		WordTextRange* sbNoteReference = [sbNote noteReference];
		CHECK_STATUS
		WordTextRange* sbReferenceRange = [doc->sbDoc 
										   createRangeStart:[sbNoteReference 
															 startOfContent]
										   end:[sbNoteReference endOfContent]];
		CHECK_STATUS
		
		if(noteType == NOTE_FOOTNOTE) {
			[[[sbNote textObject] footnoteOptions] footnoteConvert];
			CHECK_STATUS
			*returnValue = [[doc->sbDoc endnotes]
							objectAtIndex:(getEntryIndex(doc,
														 [[sbReferenceRange
														   endnotes]
														  objectAtIndex:0])-1)];
			CHECK_STATUS
		} else if(noteType == NOTE_ENDNOTE) {
			[[[sbNote textObject] endnoteOptions] endnoteConvert];
			CHECK_STATUS
			*returnValue = [[doc->sbDoc footnotes]
							objectAtIndex:(getEntryIndex(doc,
														 [[sbReferenceRange
														   footnotes]
														  objectAtIndex:0])-1)];
			CHECK_STATUS
		} else {
			DIE(@"Invalid note type")
		}
	}
	
	return STATUS_OK;
}

// Gets a property from the document. Remember to free the result!
statusCode getProperty(document_t *doc, NSString* propertyName,
					   NSString **returnValue) {
	NSInteger i = 1;
	NSMutableString* stringComponents = [NSMutableString
										 stringWithCapacity:512];
	NSString* propertyValue = nil;
	do {
		NSString* currentPropertyName = [NSString stringWithFormat:@"%@_%ld",
										 propertyName, (long) i];
		IGNORING_SB_ERRORS_BEGIN
		WordCustomDocumentProperty* property = [doc->sbProperties
												objectWithName:
												currentPropertyName];
		propertyValue = [property value];
		IGNORING_SB_ERRORS_END
		if(propertyValue) {
			[stringComponents appendString:propertyValue];
			i++;
		}
	} while([propertyValue length]);
	*returnValue = stringComponents;
	
	return STATUS_OK;
}

// Stores a property in the document.
statusCode setProperty(document_t *doc, NSString* propertyName,
					   NSString* propertyValue) {
	NSUInteger propertyValueLength = [propertyValue length];
	NSUInteger numberOfProperties = ceil(((float) propertyValueLength)
										 /MAX_PROPERTY_LENGTH);
	
	// Set fields with value
	NSMutableString* scriptToRun = nil;
	for(NSUInteger i=0; i<numberOfProperties; i++) {
		NSString* currentPropertyName = [NSString stringWithFormat:@"%@_%lu",
										 propertyName, (unsigned long) i+1];
		
		NSString* currentPropertyValue;
		if(i == numberOfProperties-1) {
			currentPropertyValue = [propertyValue substringFromIndex:
									i*MAX_PROPERTY_LENGTH];
		} else {
			currentPropertyValue = [propertyValue substringWithRange:
									NSMakeRange(i*MAX_PROPERTY_LENGTH,
												MAX_PROPERTY_LENGTH)];
		}
		
		WordCustomDocumentProperty* property = [doc->sbProperties
												objectWithName:
												currentPropertyName];
		BOOL exists = [property exists];
		
		if(!errorHasOccurred() && exists) {
			[property setValue:currentPropertyValue];
		} else {
			clearError();
			
			if(doc->wordVersion == 2004) {
				// In Word 2004, we need to insert the custom document property
				// at the end of the active document. Otherwise, it doesn't
				// work.
				if(!scriptToRun) {
					scriptToRun = [NSMutableString stringWithCapacity:
								   512*numberOfProperties+256];
					[scriptToRun appendFormat:
					 @"tell application \"Microsoft Word\"\n"];
				}
				
				[scriptToRun appendFormat:@"make new custom document property "
				 "at end of active document with properties {name:\"%@\", "
				 "value:\"%@\"}\n",
				 [[currentPropertyName
				   stringByReplacingOccurrencesOfString:@"\\"
				   withString:@"\\\\"]
				  stringByReplacingOccurrencesOfString:@"\""
				  withString:@"\\\""],
				 [[currentPropertyValue
				   stringByReplacingOccurrencesOfString:@"\\"
				   withString:@"\\\\"]
				  stringByReplacingOccurrencesOfString:@"\""
				  withString:@"\\\""]];
			} else {
				// make new custom document property at active document with
				// properties {name:currentPropertyName,
				// value:currentPropertyValue}
				NSAppleEventDescriptor *rd = [NSAppleEventDescriptor
											  recordDescriptor];
				[rd setDescriptor:[NSAppleEventDescriptor
								   descriptorWithString:currentPropertyName]
					   forKeyword:'pnam'];
				[rd setDescriptor:[NSAppleEventDescriptor
								   descriptorWithString:currentPropertyValue]
					   forKeyword:'DPVu'];
				[doc->sbApp sendEvent:'core' id:'crel' parameters:'kocl',
				 @"mCDP", 'insh', doc->sbDoc, 'prdt', rd, nil];
				CHECK_STATUS
			}
		}
	}
	
	// Word 2004: run AppleScript developed above
	if(scriptToRun) {
		[scriptToRun appendString:@"end tell"];
		NSAppleScript* script = [[NSAppleScript alloc]
								 initWithSource:scriptToRun];
		NSAppleEventDescriptor *result = [script executeAndReturnError:nil];
		[script release];
		if(!result) DIE(@"Error running AppleScript");
	}
	
	// Delete extra fields
	NSUInteger i = numberOfProperties+1;
	while(true) {
		NSString *currentPropertyName = [NSString stringWithFormat:@"%@_%lu",
										 propertyName, (unsigned long) i++];
		
		IGNORING_SB_ERRORS_BEGIN
		WordCustomDocumentProperty* property = [doc->sbProperties
												objectWithName:
												currentPropertyName];
		IGNORING_SB_ERRORS_END
		
		if(property && [property exists]) {
			[property delete];
			if(errorHasOccurred()) {
				clearError();
				// Try to set property to empty string, since sometimes delete
				// doesn't seem to work.
				[property setValue:@""];
				if(errorHasOccurred()) {
					clearError();
					break;
				}
			}
		} else {
			break;
		}
	}
	
	return STATUS_OK;
}

// Makes a field at the specified range. Uses the specified bookmark name for
// a bookmark.
statusCode insertFieldRaw(document_t *doc, const char fieldType[],
						  unsigned short noteType, WordTextRange *sbWhere,
						  NSString* bookmarkName, field_t** returnValue) {
	WordWdStoryType storyType = [sbWhere storyType];
	CHECK_STATUS
	
	WordFootnote* sbNote = nil;
	NSInteger fieldStart;
	if(noteType && storyType == WordWdStoryTypeMainTextStory) {
		// Create new note
		NSAppleEventDescriptor* noteTypeCode;
		if(noteType == NOTE_FOOTNOTE) {
			noteTypeCode = [NSAppleEventDescriptor
							descriptorWithTypeCode:'w156'];
		} else if(noteType == NOTE_ENDNOTE) {
			noteTypeCode = [NSAppleEventDescriptor
							descriptorWithTypeCode:'w157'];
		} else {
			DIE(@"Invalid field type");
		}
		
		// Select range, because otherwise we won't insert at the right
		// place. If we could pass properties {text object:sbWhere} to 
		// the make command, that would also take care of this, but alas,
		// we can't construct them with Scripting Bridge.
		[sbWhere sendEvent:'misc' id:'slct' parameters:'\00\00\00\00', nil];
		
		// make new footnote/endnote at sbDoc
		sbNote = [doc->sbApp sendEvent:'core' id:'crel' parameters:'kocl',
				  noteTypeCode, 'insh', doc->sbDoc, nil];
		
		// Clear range content
		sbWhere = [sbNote textObject];
		CHECK_STATUS
		[sbWhere setContent:@""];
		CHECK_STATUS
		
		// Move selection end past new footnote if necessary
		WordSelectionObject *sbSelection = [doc->sbApp selection];
		CHECK_STATUS
		WordTextRange *sbNoteReference = [sbNote noteReference];
		CHECK_STATUS
		[sbSelection setSelectionStart:[sbNoteReference endOfContent]];
		CHECK_STATUS
		[sbSelection setSelectionEnd:[sbNoteReference endOfContent]];
		CHECK_STATUS
		
		sbWhere = [sbNote textObject];
	} else if(storyType == WordWdStoryTypeFootnotesStory) {
		sbNote = [[sbWhere footnotes] objectAtIndex:0];
		noteType = NOTE_FOOTNOTE;
	} else if(storyType == WordWdStoryTypeEndnotesStory) {
		sbNote = [[sbWhere endnotes] objectAtIndex:0];
		noteType = NOTE_ENDNOTE;
	}
	
	// Keep track of field start, so that we can find the field in the note
	// after we lose it due to a Word bug
	if(sbNote) fieldStart = [sbWhere startOfContent];
	
	if(strcmp(fieldType, "Field") == 0) {
		WordField* sbField;
		// field
		NSAppleEventDescriptor* w170 = [NSAppleEventDescriptor
										descriptorWithTypeCode:'w170'];
		// field type:field print date
		NSAppleEventDescriptor* wFtP = [NSAppleEventDescriptor
										recordDescriptor];
		[wFtP setDescriptor:[NSAppleEventDescriptor
							 descriptorWithEnumCode:WordWdFieldTypeFieldPrintDate]
				 forKeyword:'wFtP'];
		
		// The above version works for most users of 16.9+, but some report
		// the field being inserted at cursor and not at the textRange,
		// which then breaks further actions so we move the cursor to the insertion point
		if (doc->wordVersion >= 16 && doc->wordVersion < 2000) {
			[sbWhere sendEvent:'misc' id:'slct' parameters:'\00\00\00\00', nil];
			doc->cursorMoved = YES;
			CHECK_STATUS
		}
		
		// Create as a print date field, since otherwise there is no way to
		// add any content to the result range.
		// make new field at sbWhere with properties {field type:field print
		// date}
		sbField = [doc->sbApp sendEvent:'core' id:'crel'
										parameters:'kocl',
							  w170, 'insh', sbWhere, 'prdt', wFtP, nil];
		CHECK_STATUS
		
		// We have to figure out where the field got put, because Word returns
		// an invalid reference, and Microsoft has not fixed this for 7 years.
		// If we had access to the underlying event response, it would contain
		// the entry index, but we don't, so we have to re-find the field.
		NSInteger entryIndex;
		if(sbNote) {
			// Field is within a note. First, see if we can get a valid 
			// reference easily.
			BOOL rangeNoLongerWorks = NO;
			sbField = [[sbWhere fields] objectAtIndex:0];
			
			// Range might no longer work
			rangeNoLongerWorks = errorHasOccurred() || doc->wordVersion == 2004
				|| [sbField isEqual:[NSNull null]];
			if(!rangeNoLongerWorks) {
				// It's possible the reference will fail if we try to use it
				getEntryIndex(doc, sbField);
				rangeNoLongerWorks = errorHasOccurred()
				|| [sbField isEqual:[NSNull null]];
			}
			
			if(rangeNoLongerWorks) {
				clearError();
				
				// Loop through fields in note text object until we find this
				// one
				NSArray* sbNoteFields = [[[sbNote textObject] fields] get];
				CHECK_STATUS
				for(WordField* sbNoteField in sbNoteFields) {
					if([[sbNoteField fieldCode] startOfContent]
					   == fieldStart+1) {
						sbField = sbNoteField;
					}
					CHECK_STATUS
				}
				CHECK_STATUS
			}
			
			entryIndex = getEntryIndex(doc, sbField);
			CHECK_STATUS
		} else {
			// Need to find the field within the document text. Luckily, we know
			// where we created it.
			NSInteger whereStart = [sbWhere startOfContent] - 1;
			NSInteger whereEnd = MIN([sbWhere endOfContent] + 1, [sbWhere storyLength]);
			WordTextRange* tmpRange = [doc->sbDoc createRangeStart:whereStart
															   end:whereEnd];
			NSArray* sbFields = [[tmpRange fields] get];
			NSUInteger fieldsInRange = [sbFields count];
			if(!fieldsInRange) DIE(@"Field reference lost")
			
			WordField* fieldInRange = nil;
			
			// Hack to fix https://github.com/zotero/zotero-word-for-mac-integration/issues/5
			if(fieldsInRange != 1) {
				IGNORING_SB_ERRORS_BEGIN
				for(WordField* sbTestField in sbFields) {
					if([[sbTestField resultRange] endOfContent]+1 == whereStart) {
						fieldInRange = sbTestField;
						break;
					}
				}
				IGNORING_SB_ERRORS_END
			}
			
			if(fieldInRange == nil) {
				fieldInRange = [sbFields objectAtIndex:0];
				CHECK_STATUS
			}
			
			entryIndex = getEntryIndex(doc, fieldInRange);
			CHECK_STATUS
			sbField = [[doc->sbDoc fields] objectAtIndex:(entryIndex-1)];
			CHECK_STATUS
		}
		
		if(returnValue) {
			ENSURE_OK(initField(doc, sbField, noteType, entryIndex, YES,
							 returnValue));
			if (noteType == 0) {
				doc->restoreFieldIdx = getEntryIndex(doc, (*returnValue)->sbField);
			}
			CHECK_STATUS
			return STATUS_OK;
		}
		
		return STATUS_OK;
	} else if(strcmp(fieldType, "Bookmark") == 0 ) {
		NSString* useBookmarkName;
		if(bookmarkName) {
			useBookmarkName = bookmarkName;
		} else {
			useBookmarkName = [NSString stringWithFormat:@"%@_%@",
							   BOOKMARK_REFERENCE_PROPERTY,
							   generateRandomString(21)];
		}
		
		// bookmark
		NSAppleEventDescriptor* w110 = [NSAppleEventDescriptor
										descriptorWithTypeCode:'w110'];
		// {name:"%@"}
		NSAppleEventDescriptor* reco = [NSAppleEventDescriptor
										recordDescriptor];
		[reco setDescriptor:[NSAppleEventDescriptor
							 descriptorWithString:useBookmarkName]
				 forKeyword:'pnam'];
		// make new bookmark at %@ with properties {name:%@}
		WordBookmark* sbBookmark = [doc->sbApp sendEvent:'core' id:'crel'
											  parameters:'kocl',
									w110, 'insh', sbWhere, 'prdt', reco, nil];
		
		if(returnValue) {
			return initBookmark(doc, sbBookmark, noteType, useBookmarkName, YES,
								returnValue);
		}
		
		return STATUS_OK;
	}
	
	DIE(([NSString stringWithFormat:@"Unimplemented field type %s", fieldType]))
}

// Word has "smart paste" settings that add spacing to pastes we perform
// so we have to temporarily disable these
statusCode pasteStupid(document_t *doc, WordTextRange *range) {
	bool setting = [[doc->sbApp settings] pasteAdjustWordSpacing];
	[[doc->sbApp settings] setPasteAdjustWordSpacing:NO];
	[range pasteObject];
	[[doc->sbApp settings] setPasteAdjustWordSpacing:setting];
	return STATUS_OK;
}

// Adds a field to a field list
void addValueToList(void* value, listNode_t** listStart,
					listNode_t** listEnd) {
	listNode_t* node = (listNode_t*) malloc(sizeof(listNode_t));
	node->value = value;
	node->next = NULL;
	
	if(*listEnd) {
		(*listEnd)->next = node;
		*listEnd = node;
	} else {
		*listStart = *listEnd = node;
	}
}
