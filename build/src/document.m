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

// Prototypes for file-local objects and functions
void freeFieldList(listNode_t* fieldList, bool freeFields);
statusCode getFieldCollections(document_t *doc, NSArray** fieldCollections);
statusCode extractFieldsFromNotes(SBElementArray* notes, 
								  NSArray** returnValue);
statusCode noteSwap(document_t *doc, WordFootnote* sbNote,
					unsigned short noteType,
					WordFootnote **returnValue);
@interface FieldGetter : NSObject {
	document_t* document;
	char* fieldType;
	void (*onProgress)(int progress);
	listNode_t** returnNode;
}

- (FieldGetter*) initWithDocument:(document_t*)aDocument
						fieldType:(char*)aFieldType
					   onProgress:(void (*)(int progress))aOnProgress
					   returnNode:(listNode_t**)aReturnNode;
- (void) getFields;
- (void) progress:(NSNumber*)progressValue;
@end

// This object handles acquiring fields on a separate thread
@implementation FieldGetter

- (FieldGetter*) initWithDocument:(document_t*)aDocument
						fieldType:(char*)aFieldType
					   onProgress:(void (*)(int progress))aOnProgress
					   returnNode:(listNode_t**)aReturnNode {
	document = aDocument;
	fieldType = aFieldType;
	onProgress = aOnProgress;
	returnNode = aReturnNode;
	return self;
}

// Gets fields. Call this from off the main thread.
- (void) getFields {
	NSNumber* progressValue;
	statusCode status = getFields(document, fieldType, returnNode);
	if(status) {
		progressValue = [NSNumber numberWithInt:-1];
	} else {
		progressValue = [NSNumber numberWithInt:100];
	}
	
	[self performSelectorOnMainThread:@selector(progress:)
						   withObject:progressValue waitUntilDone:NO];
}

// Dispatches the progress handler, which is a function pointer to a JavaScript
// function.
- (void) progress:(NSNumber*)progressValue {
	(*onProgress)([progressValue intValue]);
}
@end

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

	if([doc->sbView viewType] == WordE202WordNoteView) {
		displayAlert(doc,
					 "Zotero cannot insert a citation here because Word does "
					 "not support inserting fields in Notebook Layout.",
					 DIALOG_ICON_STOP, DIALOG_BUTTONS_OK, NULL);
		RETURN_STATUS_LOCKED(doc, STATUS_EXCEPTION_ALREADY_DISPLAYED)
	}
	CHECK_STATUS_LOCKED(doc)
	
	WordE160 position = [[doc->sbApp selection] storyType];
	CHECK_STATUS_LOCKED(doc)
	*returnCode = (strcmp(fieldType, "Bookmark") != 0
				   && (position == WordE160FootnotesStory
					   || position == WordE160EndnotesStory))
					   || position == WordE160MainTextStory;
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
	
	WordTextRange *range = [doc->sbDoc createRangeStart:0 end:[EXPORTED_DOCUMENT_MARKER length]];
	NSString *text = [range content];
	if (errorHasOccurred()) {
		clearError();
	}
	else if (text != nil && [text rangeOfString:EXPORTED_DOCUMENT_MARKER].location == 0) {
		*returnValue = copyNSString(EXPORTED_DOCUMENT_MARKER);
		RETURN_STATUS_LOCKED(doc, STATUS_OK);
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
	
	if(strcmp(fieldType, "Bookmark") == 0 && !noteType) {
		[[doc->sbApp selection] setContent:@ FIELD_PLACEHOLDER];
		CHECK_STATUS_LOCKED(doc)
	}
	
	statusCode status = insertFieldRaw(doc, fieldType, noteType, sbWhere,
									   nil, returnValue);
	if(!status && strcmp(fieldType, "Bookmark") != 0) {
		setText(*returnValue, FIELD_PLACEHOLDER, false);
	}
	RETURN_STATUS_LOCKED(doc, status)
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

statusCode getFieldsAsync(document_t *doc, const char fieldType[],
						  listNode_t** returnNode,
						  void (*onProgress)(int progress)) {
	HANDLE_EXCEPTIONS_BEGIN
	size_t bufferSize = strlen(fieldType)+1;
	char* fieldTypeCopy = (char*) malloc(bufferSize);
	memcpy(fieldTypeCopy, fieldType, bufferSize);
	FieldGetter* fieldGetter = [[FieldGetter alloc]
								initWithDocument:doc
								fieldType:fieldTypeCopy
								onProgress:onProgress
								returnNode:returnNode];
	
	[fieldGetter performSelectorInBackground:@selector(getFields)
								  withObject:nil];
	
	return STATUS_OK;
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
				// Copy text for in-text to note conversions
				if (strcmp(toFieldType, "Bookmark") != 0) {
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
										[[sbNote noteReference]
										 collapseRangeDirection:
										 WordE132CollapseStart], nil,
										&newField);
				ENSURE_OK_LOCKED(doc, status)
				[newField->sbContentRange setContent:[field->sbContentRange content]];
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
						   objectWithName:@"Bibliography"];
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
							 descriptorWithEnumCode:WordE128StyleTypeParagraph]
				 forKeyword:'2504'];
		[reco setDescriptor:[NSAppleEventDescriptor
							 descriptorWithEnumCode:WordE184StyleNormal]
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
							 descriptorWithEnumCode:WordE145AlignTabLeft]
				 forKeyword:'1918'];
		[reco setDescriptor:[NSAppleEventDescriptor
							 descriptorWithEnumCode:WordE170TabLeaderSpaces]
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
		[doc->sbApp createNewFieldTextRange:insertRange fieldType:WordE183FieldHyperlink fieldText:IMPORT_LINK_URL preserveFormatting:NO];
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
	range = [range collapseRangeDirection:WordE132CollapseEnd];
	[doc->sbApp insertParagraphAt:range];
	range = [range collapseRangeDirection:WordE132CollapseEnd];
	[doc->sbApp insertParagraphAt:range];
	range = [[doc->sbDoc textObject] collapseRangeDirection:WordE132CollapseEnd];
	[doc->sbApp createNewFieldTextRange:range fieldType:WordE183FieldHyperlink fieldText:IMPORT_LINK_URL preserveFormatting:NO];
	WordField *insertedHyperlink = [[doc->sbDoc fields] lastObject];
	[[insertedHyperlink resultRange] setContent:importDocData];
	CHECK_STATUS_LOCKED(doc)

	// Import instructions
	range = [doc->sbDoc textObject];
	range = [range collapseRangeDirection:WordE132CollapseStart];
	[doc->sbApp insertParagraphAt:range];
	range = [range collapseRangeDirection:WordE132CollapseStart];
	[doc->sbApp insertParagraphAt:range];
	[doc->sbApp insertText:[NSString stringWithUTF8String:importInstructions] at:range];
	range = [range collapseRangeDirection:WordE132CollapseStart];
	[doc->sbApp insertParagraphAt:range];
	[doc->sbApp insertText:EXPORTED_DOCUMENT_MARKER at:range];
	CHECK_STATUS_LOCKED(doc)
	
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
			WordE183 sbFieldType = [link fieldType];
			if (sbFieldType != WordE183FieldHyperlink) {
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
				IGNORING_SB_ERRORS_BEGIN
				WordFont* font = [[link resultRange] fontObject];
				[font setUnderline:NO];
				[font setColorIndex:WordE110Auto];
				IGNORING_SB_ERRORS_END
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
	
	// Remove 3 paragraphs: export marker, import instructions and an empty paragraph
	WordTextRange *range = [doc->sbDoc textObject];
	range = [range collapseRangeDirection:WordE132CollapseStart];
	range = [range moveEndOfRangeBy:WordE129AParagraphItem count:3];
	[range setContent:@""];
	CHECK_STATUS_LOCKED(doc);
	
	// Don't attempt to restore cursor location. It doesn't work since the document size changes
	// and it doesn't make sense either.
	doc->cursorMoved = NO;
	
	RETURN_STATUS_LOCKED(doc, STATUS_OK);
	HANDLE_EXCEPTIONS_END
}

// Run on exit to clean up anything we played with
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
	
	deleteTemporaryFile();
	
	RETURN_STATUS_LOCKED(doc, restoreCursor(doc))
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
								getStoryRangeStoryType:WordE160FootnotesStory]
							   fields];
		CHECK_STATUS;
		fieldCollections[2] = [[doc->sbDoc
								getStoryRangeStoryType:WordE160EndnotesStory]
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
							moveEndOfRangeBy:WordE129ACharacterItem count:0];
		CHECK_STATUS
		WordTextRange* sbCreateRange = [[sbNote noteReference] 
										collapseRangeDirection:WordE132CollapseEnd];
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
	WordE160 storyType = [sbWhere storyType];
	CHECK_STATUS
	
	WordFootnote* sbNote = nil;
	NSInteger fieldStart;
	if(noteType && storyType == WordE160MainTextStory) {
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
	} else if(storyType == WordE160FootnotesStory) {
		sbNote = [[sbWhere footnotes] objectAtIndex:0];
		noteType = NOTE_FOOTNOTE;
	} else if(storyType == WordE160EndnotesStory) {
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
							 descriptorWithEnumCode:WordE183FieldPrintDate]
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
			NSInteger whereStart = [sbWhere startOfContent];
			WordTextRange* tmpRange = [doc->sbDoc createRangeStart:whereStart-1
															   end:([sbWhere endOfContent]+1)];
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
			doc->restoreFieldIdx = getEntryIndex(doc, (*returnValue)->sbField);
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
