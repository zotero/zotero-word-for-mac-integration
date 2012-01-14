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

// This object is used to asynchronously get fields
//@implementation FieldGetter
//- (void)getFields:(id)param{}
//@end

statusCode getFieldCollections(document_t *doc, NSArray** fieldCollections);
statusCode noteSwap(document_t *doc, WordFootnote* sbNote,
					unsigned short noteType,
					WordFootnote **returnValue);

// Frees a document struct. This 
void freeDocument(document_t* doc) {
	[doc->sbApp release];
	[doc->sbDoc release];
	[doc->sbView release];
	[doc->sbProperties release];
	free(doc->wordPath);
	free(doc);
}

// Free a field list. This only free the structure, and not the fields contained
// within it.
void freeFieldList(fieldListNode_t* fieldList) {
	fieldListNode_t* nextNode = fieldList;
	while(nextNode) {
		fieldListNode_t* currentNode = nextNode;
		nextNode = currentNode->next;
		free(currentNode);
	}
}

// Activates Word to deal with a document
statusCode activate(document_t *doc) {
	[doc->sbApp activate];
	CHECK_STATUS
	
	return STATUS_OK;
}

// Disables Track Changes settings so that we can read field codes
statusCode prepareReadFieldCode(document_t *doc) {
	if(doc->statusInsertionsAndDeletions) {
		[doc->sbView setShowInsertionsAndDeletions:NO];
		CHECK_STATUS
		doc->statusInsertionsAndDeletions = NO;
	}
	
	if(doc->statusFormatChanges) {
		[doc->sbView setShowFormatChanges:NO];
		CHECK_STATUS
		doc->statusFormatChanges = NO;
	}
	
	return STATUS_OK;
}

// Determines whether it is possible to insert a field at the current cursor
// position
statusCode canInsertField(document_t *doc, const char fieldType[],
						  bool* returnCode) {
	if([doc->sbView viewType] == WordE202WordNoteView) {
		displayAlert("Zotero cannot insert a citation here because Word does "
					 "not support inserting fields in Notebook Layout.",
					 DIALOG_ICON_STOP, DIALOG_BUTTONS_OK, NULL);
		return STATUS_EXCEPTION_ALREADY_DISPLAYED;
	}
	CHECK_STATUS
	
	WordE160 position = [[doc->sbApp selection] storyType];
	CHECK_STATUS
	*returnCode = (strcmp(fieldType, "Bookmark") != 0
				   && (position == WordE160FootnotesStory
					   || position == WordE160EndnotesStory))
	               || position == WordE160MainTextStory;
	return STATUS_OK;
}

// Determines whether the cursor is in a field. Returns the a field struct if
// it is, or NULL if it is not.
statusCode cursorInField(document_t *doc, const char fieldType[],
						 field_t **returnValue) {
	WordSelectionObject* sbSelection = [doc->sbApp selection];
	CHECK_STATUS
	
	if(strcmp(fieldType, "Field") == 0) {
		SBElementArray* sbFields = [sbSelection fields];
		CHECK_STATUS
		if([sbFields count]) {
			return initField(doc, (WordField*) [sbFields objectAtIndex:0], -1,
							 -1, NO, returnValue);
		}
	
		sbFields = [[[[sbSelection paragraphs] objectAtIndex:0] textObject]
				  fields];
		CHECK_STATUS
		if([sbFields count]) {
			// Check if fields are in the selection
			NSInteger selectionStart = [sbSelection selectionStart];
			CHECK_STATUS
			NSInteger selectionEnd = [sbSelection selectionEnd];
			CHECK_STATUS
			for(WordField* sbTestField in sbFields) {
				NSInteger fieldStart = [[sbTestField resultRange]
								   startOfContent];
				CHECK_STATUS
				NSInteger fieldEnd = [[sbTestField resultRange] endOfContent];
				CHECK_STATUS
				
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
					   if(status || *returnValue) return status;
				   } else if(fieldStart > selectionEnd) {
					   // Cursor already past end of selection
					   break;
				   }
			}
		}
	} else if(strcmp(fieldType, "Bookmark") == 0) {
		SBElementArray *sbBookmarks = [sbSelection bookmarks];
		CHECK_STATUS
		
		if(![doc->sbDoc isEqual:[NSNull null]]) {
			for(WordBookmark* sbBookmark in sbBookmarks) {
				statusCode status = initBookmark(doc, sbBookmark, -1, nil, NO,
												 returnValue);
				if(status || *returnValue) return status;
			}
			CHECK_STATUS
		}
	}
	
	*returnValue = NULL;
	return STATUS_OK;
}

// Gets document data
statusCode getDocumentData(document_t *doc, char **returnValue) {
	NSString* returnString;
	statusCode status = getProperty(doc, PREFS_PROPERTY, &returnString);
	if(status) return status;
	*returnValue = copyNSString(returnString);
	
	return STATUS_OK;
}

// Sets document data
statusCode setDocumentData(document_t *doc, const char documentData[]) {
	NSString* propertyValue = [NSString stringWithUTF8String:documentData];
	return setProperty(doc, PREFS_PROPERTY, propertyValue);
}

// Gets a property from the document. Remember to free the result!
statusCode getProperty(document_t *doc, NSString* propertyName,
					   NSString **returnValue) {
	NSInteger i = 1;
	NSMutableString* stringComponents = [NSMutableString
										 stringWithCapacity:512];
	NSString* propertyValue = nil;
	do {
		NSString* currentPropertyName = [NSString stringWithFormat:@"%@_%d",
										 propertyName, i];
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
	} while(propertyValue);
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
	for(NSUInteger i=0; i<numberOfProperties; i++) {
		NSString *currentPropertyName = [NSString stringWithFormat:@"%@_%d",
										 propertyName, i+1];
		
		NSString *currentPropertyValue;
		if(i == numberOfProperties-1) {
			currentPropertyValue = [propertyValue substringFromIndex:
									i*MAX_PROPERTY_LENGTH];
		} else {
			currentPropertyValue = [propertyValue  substringWithRange:
									NSMakeRange(i*MAX_PROPERTY_LENGTH,
												MAX_PROPERTY_LENGTH)];
		}
		
		WordCustomDocumentProperty* property = [doc->sbProperties
												objectWithName:
												currentPropertyName];
		[property setValue:currentPropertyValue];
		
		if(errorHasOccurred()) {
			clearError();
			
			// make new custom document property at active document with
			// properties {name:currentPropertyName, value:currentPropertyValue}
			NSAppleEventDescriptor *rd = [NSAppleEventDescriptor
										  recordDescriptor];
			[rd setDescriptor:[NSAppleEventDescriptor
							   descriptorWithString:currentPropertyName]
			 forKeyword:'pnam'];
			[rd setDescriptor:[NSAppleEventDescriptor
							   descriptorWithString:currentPropertyValue]
			 forKeyword:'DPVu'];
			[doc->sbApp sendEvent:'core' id:'crel' parameters:'kocl', @"mCDP",
			 'insh', doc->sbDoc, 'prdt', rd, nil];
			CHECK_STATUS
		}
	}
	
	// Delete extra fields
	NSUInteger i = numberOfProperties+1;
	while(true) {
		NSString *currentPropertyName = [NSString stringWithFormat:@"%s_%d",
										 propertyName, i];
		
		IGNORING_SB_ERRORS_BEGIN
		WordCustomDocumentProperty* property = [doc->sbProperties
												objectWithName:
												currentPropertyName];
		IGNORING_SB_ERRORS_END
		
		if(property && [property exists]) {
			[property delete];
			if(errorHasOccurred()) {
				clearError();
				break;
			}
		} else {
			break;
		}
	}
	
	return STATUS_OK;
}

// Makes a field at the selection.
statusCode insertField(document_t *doc, const char fieldType[],
					   unsigned short noteType, field_t **returnValue) {
	WordTextRange* sbWhere = [[doc->sbApp selection] textObject];
	CHECK_STATUS
	
	if(strcmp(fieldType, "Bookmark") == 0 && !noteType) {
		[[doc->sbApp selection] setContent:@ FIELD_PLACEHOLDER];
		CHECK_STATUS
	}
	
	statusCode status = insertFieldRaw(doc, fieldType, noteType, sbWhere,
									   nil, returnValue);
	if(!status) setText(*returnValue, FIELD_PLACEHOLDER, false);
	return status;
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
	if(noteType) {
		if(storyType == WordE160MainTextStory) {
			// Create new note
			
			NSString* noteTypeCode;
			if(noteType == NOTE_FOOTNOTE) {
				noteTypeCode = @"w156";
			} else if(noteType == NOTE_ENDNOTE) {
				noteTypeCode = @"w157";
			} else {
				DIE(@"Invalid field type");
			}
			
			// make new footnote/endnote at active document
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
			if([sbSelection selectionEnd] == [sbNoteReference startOfContent]) {
				[sbSelection setSelectionStart:[sbNoteReference endOfContent]];
				CHECK_STATUS
				[sbSelection setSelectionEnd:[sbNoteReference endOfContent]];
				CHECK_STATUS
			}
			CHECK_STATUS
			
			sbWhere = [sbNote textObject];
		} else if(storyType == WordE160FootnotesStory) {
			sbNote = [[sbWhere footnotes] objectAtIndex:0];
		} else if(storyType == WordE160EndnotesStory) {
			sbNote = [[sbWhere endnotes] objectAtIndex:0];
		}
		fieldStart = [sbWhere startOfContent];
	}
	
	if(strcmp(fieldType, "Field") == 0) {
		// field
		NSAppleEventDescriptor* w170 = [NSAppleEventDescriptor
										descriptorWithTypeCode:'w170'];
		// field type:field print date
		NSAppleEventDescriptor* wFtP = [NSAppleEventDescriptor
										recordDescriptor];
		[wFtP setDescriptor:[NSAppleEventDescriptor
							 descriptorWithEnumCode:WordE183FieldPrintDate]
				 forKeyword:'wFtP'];
		
		// Create as a print date field, since otherwise there is no way to
		// add any content to the result range.
		// make new field at %@ with properties {field type:field print date}
		WordField* sbField = [doc->sbApp sendEvent:'core' id:'crel'
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
			rangeNoLongerWorks = errorHasOccurred()
				|| [sbField isEqual:[NSNull null]];
			if(!rangeNoLongerWorks) {
				// It's possible the reference will fail if we try to use it
				[sbField entry_index];
				rangeNoLongerWorks = errorHasOccurred()
					|| [sbField isEqual:[NSNull null]];
			}
			
			if(rangeNoLongerWorks) {
				clearError();
				
				// Loop through fields in note text object until we find this
				// one
				SBElementArray* sbNoteFields = [[sbNote textObject] fields];
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
			
			entryIndex = [sbField entry_index];
			CHECK_STATUS
		} else {
			// Need to find the field within the document text. Luckily, we know
			// where we created it.
			WordTextRange* tmpRange = [doc->sbDoc createRangeStart:
									   ([sbWhere startOfContent]-1)
									   end:([sbWhere endOfContent]+1)];
			CHECK_STATUS
			entryIndex = [[[tmpRange fields] objectAtIndex:0] entry_index];
			CHECK_STATUS
			sbField = [[doc->sbDoc fields] objectAtIndex:(entryIndex-1)];
			CHECK_STATUS
		}
		
		if(returnValue) {
			return initField(doc, sbField, noteType, entryIndex, YES,
							 returnValue);
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

// Gets fields
statusCode getFields(document_t *doc, const char fieldType[],
					 fieldListNode_t** returnNode) {
	fieldListNode_t* fieldListStart = NULL;
	fieldListNode_t* fieldListEnd = NULL;
	
	if(strcmp(fieldType, "Field") == 0) {
		// Get fields in main body, footnotes, and endnotes
		NSArray* fieldCollections[3];
		statusCode status = getFieldCollections(doc, fieldCollections);
		if(status) return status;
		
		long currentFieldIndices[] = {0, 0, 0};
		field_t* currentFields[] = {NULL, NULL, NULL};
		
		// Get numbers of fields and first fields
		unsigned long fieldCollectionSizes[3];
		for(unsigned short noteType=0; noteType<3; noteType++) {
			fieldCollectionSizes[noteType] = [fieldCollections[noteType] count];
			
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
					if(status) return status;
					currentFieldIndices[noteType]++;
				}
			}
		}
		
		BOOL isNextField = YES;
		do {
			isNextField = NO;
			
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
					if(status) return status;
					
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
			
			if(isNextField) {
				// Add node to linked list
				fieldListNode_t* node = (fieldListNode_t*)
					malloc(sizeof(fieldListNode_t));
				node->field = currentFields[noteTypeA];
				node->next = NULL;
				
				if(fieldListEnd) {
					fieldListEnd->next = node;
					fieldListEnd = node;
				} else {
					fieldListStart = fieldListEnd = node;
				}
				
				// Get next node of type
				currentFields[noteTypeA] = NULL;
				while(currentFields[noteTypeA] == NULL &&
					  currentFieldIndices[noteTypeA] <
					  fieldCollectionSizes[noteTypeA]) {
					
					statusCode status = initField(doc, [fieldCollections[noteTypeA]
						objectAtIndex:(currentFieldIndices[noteTypeA])], noteTypeA,
						currentFieldIndices[noteTypeA]+1, NO,
						&currentFields[noteTypeA]);
					if(status) return status;
					currentFieldIndices[noteTypeA]++;
				}
			}
		} while(isNextField);
	} else if(strcmp(fieldType, "Bookmark") == 0) {
		[doc->sbView setShowBookmarks:YES];
		SBElementArray *sbBookmarks = [doc->sbDoc bookmarks];
		CHECK_STATUS
		// We are going to allocate enough memory for all of the bookmarks to be
		// fields. This may not end up being the case, but we need arrays and 
		// not linked lists to be able to use the 
		field_t** fields = (field_t**) malloc(sizeof(field_t *) * 
											[sbBookmarks count]);
		unsigned long nFields = 0;
		CHECK_STATUS
		
		// Generate field structures
		for(WordBookmark* sbBookmark in sbBookmarks) {
			statusCode status = initBookmark(doc, sbBookmark, -1, nil, NO,
											 &fields[nFields]);
			if(status) return status;
			if(fields[nFields]) nFields++;
		}
		
		// Sort
		statusCode status = STATUS_OK;
		qsort_r(fields, nFields, sizeof(field_t *), &status,
				&compareFieldsQsort);
		
		// Generate the linked list
		for(unsigned long i=0; i<nFields; i++) {
			fieldListNode_t* node = (fieldListNode_t*)
				malloc(sizeof(fieldListNode_t));
			node->field = fields[i];
			node->next = NULL;
			
			if(fieldListEnd) {
				fieldListEnd->next = node;
				fieldListEnd = node;
			} else {
				fieldListStart = fieldListEnd = node;
			}
		}
	} else {
		DIE(([NSString stringWithFormat:@"Unknown field type \"%s\"",
			fieldType]))
	}
	
	*returnNode = fieldListStart;
	return STATUS_OK;
}

statusCode convert(document_t *doc, field_t* fields[], unsigned long nFields,
				   const char toFieldType[], unsigned short toNoteTypes[]) {
	// Get field collections, keep track of offsets
	NSArray* fieldCollections[3];
	statusCode status = getFieldCollections(doc, fieldCollections);
	if(status) return status;
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
	unsigned long i = nFields;
	while(i != 0) {
		i--;
		field_t* field = fields[i];
		
		unsigned short toNoteType = toNoteTypes[i];
		
		if(field->sbField) {
			[field->sbField release];
			[field->sbContentRange release];
			
			field->sbField = [fieldCollections[field->noteType] objectAtIndex:
					   (field->entryIndex+offsets[field->noteType]-1)];
			CHECK_STATUS
			field->sbContentRange = [field->sbField resultRange];
			CHECK_STATUS
			
			[field->sbField retain];
			[field->sbContentRange retain];
		}
		
		// First, handle converting between footnote, endnote, and inline
		WordFootnote* sbNote;
		if(field->noteType == NOTE_FOOTNOTE) {
			sbNote = [[field->sbContentRange footnotes] objectAtIndex:0];
			CHECK_STATUS;
		} else if(field->noteType == NOTE_ENDNOTE) {
			sbNote = [[field->sbContentRange endnotes] objectAtIndex:0];
			CHECK_STATUS;
		} else {
			sbNote = nil;
		}
		
		// Word has special facilities for converting footnotes/endnotes
		if((toNoteType == NOTE_ENDNOTE && field->noteType == NOTE_FOOTNOTE) ||
		   (toNoteType == NOTE_FOOTNOTE && field->noteType == NOTE_ENDNOTE)) {
			// Check which fields are in this note
			unsigned long nFieldsInNote = 0;
			while(nFieldsInNote < nFields &&
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
				CHECK_STATUS
				bookmarkStarts = (NSInteger*) malloc(sizeof(NSInteger*)
													 * nFieldsInNote);
				bookmarkEnds = (NSInteger*) malloc(sizeof(NSInteger*)
													 * nFieldsInNote);
				for(unsigned long j=0; j<nFieldsInNote; j++) {
					bookmarkStarts[j] = [fields[i+j]->sbBookmark
										 startOfBookmark];
					CHECK_STATUS
					bookmarkEnds[j] = [fields[i+j]->sbBookmark endOfBookmark];
					CHECK_STATUS
				}
			}
			
			// Convert fields and get the mark
			WordFootnote *sbNewNote;
			status = noteSwap(doc, sbNote, field->noteType, &sbNewNote);
			if(status) return status;
			sbNote = sbNewNote;
			
			// Find the fields again
			if(field->sbField) {
				offsets[field->noteType] -= nFieldsInNote;
				offsets[toNoteType] += nFieldsInNote;
				field->noteType = toNoteType;
				
				if(toField) {
					// We don't need to convert the other entries in this note.
					i += nFieldsInNote-1;
				} else if(toBookmark) {
					// Need to update conversions so that notes can be converted
					// to bookmarks later
					[field->sbField release];
					[field->sbContentRange release];
					
					field->sbField = [[[sbNote textObject] fields]
									  objectAtIndex:0];
					CHECK_STATUS
					field->sbContentRange = [field->sbField resultRange];
					CHECK_STATUS
					NSInteger entryIndex = [field->sbField entry_index];
					CHECK_STATUS
					
					[field->sbField retain];
					[field->sbContentRange retain];
					
					for(unsigned long j=0; j<nFieldsInNote; j++) {
						fields[i+j]->entryIndex = entryIndex
							- (field->entryIndex) - offsets[toNoteType];
						CHECK_STATUS
						fields[i+j]->noteType = toNoteType;
						CHECK_STATUS
					}
				}
			} else if(field->sbBookmark) {
				WordTextRange* sbNoteTextObject = [sbNote textObject];
				CHECK_STATUS
				NSInteger startDifference = [sbNoteTextObject startOfContent]
				- oldNoteStart;
				CHECK_STATUS
				for(unsigned long j=0; j<nFieldsInNote; j++) {
					fields[i+j]->noteType = toNoteType;
					status = insertFieldRaw(doc, "Bookmark", 0,
											sbNoteTextObject,
											fields[i+j]->bookmarkNameNS, NULL);
					if(status) return status;
					[fields[i+j]->sbBookmark
					 setStartOfBookmark:bookmarkStarts[j]+startDifference];
					CHECK_STATUS
					[fields[i+j]->sbBookmark
					 setEndOfBookmark:bookmarkEnds[j]+startDifference];
					CHECK_STATUS
				}
			}
		}
		
		// Check whether the field is not the entire content of the note. If
		// this is the case, then we don't need to convert fields to notes
		BOOL inlineNoteField = sbNote && [[[sbNote textObject] content]
										   isNotEqualTo:[field->sbContentRange
														 content]];
		CHECK_STATUS
		
		// Skip further processing if it's unnecessary
		if(((field->sbField && toField) || (field->sbBookmark && toBookmark))
		   && (field->noteType == toNoteType || inlineNoteField)) {
			continue;
		}
		
		// Convert between field types and between inline and footnote/endnote
		// citations
		if(field->sbField) {
			prepareReadFieldCode(doc);
			
			field_t* newField;
			if(!(field->noteType)) {
				// Convert an in-text citation to note or bookmark
				NSInteger fieldStart = [[field->sbField fieldCode]
										startOfContent];
				CHECK_STATUS
				[field->sbField delete];
				CHECK_STATUS
				WordTextRange* sbWhere = [doc->sbDoc
										  createRangeStart:fieldStart-1
										  end:fieldStart-1];
				CHECK_STATUS
				status = insertFieldRaw(doc, toFieldType, toNoteType, sbWhere,
										nil, &newField);
				if(status) return status;
			} else if(field->noteType == toNoteType || inlineNoteField) {
				// Convert fields inside a note to bookmarks
				status = insertFieldRaw(doc, toFieldType, 0,
										field->sbContentRange, nil, &newField);
				if(status) return status;
				[newField->sbBookmark
				 setStartOfBookmark:([newField->sbBookmark endOfBookmark]+1)];
				CHECK_STATUS
				[field->sbField delete];
				CHECK_STATUS
			} else {
				// Convert note to inline
				status = insertFieldRaw(doc, toFieldType, toNoteType,
										[[sbNote noteReference]
										 collapseRangeDirection:
										 WordE132CollapseStart], nil,
										&newField);
				if(status) return status;
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
				CHECK_STATUS
				[[sbNote noteReference] setContent:@""];
				CHECK_STATUS
				sbWhere = [doc->sbDoc createRangeStart:referenceStart
												   end:referenceStart];
				CHECK_STATUS
			} else if(!(field->noteType) && toNoteType && !inlineNoteField) {
				// Convert in-text to note
				// Delete old in-text citation and put footnote reference
				// where it was
				NSInteger bookmarkStart = [field->sbBookmark startOfBookmark];
				CHECK_STATUS
				[field->sbContentRange setContent:@""];
				CHECK_STATUS
				sbWhere = [doc->sbDoc createRangeStart:bookmarkStart
												   end:bookmarkStart];
				CHECK_STATUS
			} else {
				// Convert in-text to in-text
				sbWhere = field->sbContentRange;
			}
			
			// Create bookmark or field
			if(toField) {
				field_t* newField;
				status = insertFieldRaw(doc, toFieldType, toNoteType, sbWhere,
										nil, &newField);
				if(status) return status;
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
	
	return STATUS_OK;
}

// Sets the "Bibliography" style
statusCode setBibliographyStyle(document_t* doc, long firstLineIndent, 
								long bodyIndent, unsigned long lineSpacing,
								unsigned long entrySpacing, long tabStops[],
								unsigned long tabStopCount) {
	WordWordStyle* sbBibliographyStyle;
	
	sbBibliographyStyle = [[doc->sbDoc WordStyles]
						   objectWithName:@"Bibliography"];
	if(!errorHasOccurred()) {
		SBElementArray* sbTabStops = [[sbBibliographyStyle paragraphFormat]
									  tabStops];
		[sbTabStops removeAllObjects];
	}
	
	if(errorHasOccurred()) {
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
		 [NSAppleEventDescriptor descriptorWithTypeCode:'w173'], 'insh', doc,
		 'prdt', reco, nil];
		CHECK_STATUS
	}
	
	WordParagraphFormat* sbParagraphFormat = [sbBibliographyStyle
											  paragraphFormat];
	CHECK_STATUS
	[sbParagraphFormat setFirstLineIndent:((double) firstLineIndent)/20];
	CHECK_STATUS
	[sbParagraphFormat setParagraphFormatLeftIndent:((double) bodyIndent)/20];
	CHECK_STATUS
	[sbParagraphFormat setLineSpacing:((double) lineSpacing)/20];
	CHECK_STATUS
	[sbParagraphFormat setSpaceAfter:((double) entrySpacing)/20];
	CHECK_STATUS
	
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
	
	return STATUS_OK;
}

// Run on exit to clean up anything we played with
statusCode cleanup(document_t *doc) {
	if(doc->restoreInsertionsAndDeletions
	   && !doc->statusInsertionsAndDeletions) {
		[doc->sbView setShowInsertionsAndDeletions:YES];
		CHECK_STATUS
		doc->statusInsertionsAndDeletions = YES;
	}
	if(doc->restoreFormatChanges && !doc->statusFormatChanges) {
		[doc->sbView setShowFormatChanges:YES];
		CHECK_STATUS
		doc->statusFormatChanges = YES;
	}
	if(doc->restoreFullScreenMode && !doc->statusFullScreenMode) {
		[doc->sbView setFullScreen:YES];
		CHECK_STATUS
		doc->statusFullScreenMode = YES;
	}
	//deleteTemporaryFile();
	return STATUS_OK;
}

// Get NSArrays representing fields in different parts of the document.
statusCode getFieldCollections(document_t *doc, NSArray** fieldCollections) {
	fieldCollections[0] = [doc->sbDoc fields];
	CHECK_STATUS;
	fieldCollections[1] = [[doc->sbDoc
							getStoryRangeStoryType:WordE160FootnotesStory] fields];
	CHECK_STATUS;
	fieldCollections[2] = [[doc->sbDoc
							getStoryRangeStoryType:WordE160EndnotesStory] fields];
	CHECK_STATUS;
	return STATUS_OK;
}

// Swap footnotes and endnotes
statusCode noteSwap(document_t *doc, WordFootnote* sbNote,
					unsigned short noteType,
					WordFootnote **returnValue) {
	if(doc->isWord2004) {
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
					   parameters:'kocl', @"w157", 'insh', sbCreateRange, nil];
			CHECK_STATUS
			*returnValue = [[doc->sbDoc endnotes]
							objectAtIndex:([[[sbReferenceRange endnotes]
											 objectAtIndex:0] entry_index]-1)];
			CHECK_STATUS
		} else if(noteType == NOTE_ENDNOTE) {
			// make new footnote at active document
			[doc->sbApp sendEvent:'core' id:'crel'
					   parameters:'kocl', @"w156", 'insh', sbCreateRange, nil];
			CHECK_STATUS
			*returnValue = [[doc->sbDoc footnotes]
							objectAtIndex:([[[sbReferenceRange footnotes]
											 objectAtIndex:0] entry_index]-1)];
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
							objectAtIndex:([[[sbReferenceRange endnotes]
											 objectAtIndex:0] entry_index]-1)];
			CHECK_STATUS
		} else if(noteType == NOTE_ENDNOTE) {
			[[[sbNote textObject] endnoteOptions] endnoteConvert];
			CHECK_STATUS
			*returnValue = [[doc->sbDoc footnotes]
							objectAtIndex:([[[sbReferenceRange footnotes]
											 objectAtIndex:0] entry_index]-1)];
			CHECK_STATUS
		} else {
			DIE(@"Invalid note type")
		}
	}
	
	return STATUS_OK;
}

