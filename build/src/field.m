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

NSString* FIELD_PREFIXES[] = {MAIN_FIELD_PREFIX, @" CSL_", NULL};
NSString* BOOKMARK_PREFIXES[] = {@"ZOTERO_", @"CSL_", NULL};

// Allocates a field structure based on a WordField, optionally checking to make
// sure that the field code actually matches a Zotero field.
statusCode initField(document_t *doc, WordField* sbField, short noteType,
					 NSInteger entryIndex, BOOL ignoreCode,
					 field_t **returnValue) {
	
	// sbFields initiated from cursor have their textObjects tied to the selection
	// so reusing this code tends to go haywire if you move the cursor somewhere
	// Thus we get the sbField from the document fields collection here
	if(noteType == -1) {
		WordWdStoryType storyType = [[sbField resultRange] storyType];
		CHECK_STATUS
		entryIndex = getEntryIndex(doc, sbField);
		CHECK_STATUS
		if(storyType == WordWdStoryTypeFootnotesStory) {
			noteType = NOTE_FOOTNOTE;
			sbField = [[[doc->sbDoc getStoryRangeStoryType:WordWdStoryTypeFootnotesStory] fields]
					   objectAtIndex:(entryIndex-1)];
		} else if(storyType == WordWdStoryTypeEndnotesStory) {
			noteType = NOTE_ENDNOTE;
			sbField = [[[doc->sbDoc getStoryRangeStoryType:WordWdStoryTypeEndnotesStory] fields]
					   objectAtIndex:(entryIndex-1)];
		} else {
			noteType = 0;
			sbField = [[doc->sbDoc fields] objectAtIndex:(entryIndex-1)];
		}
	}
	WordTextRange* sbCodeRange = [sbField fieldCode];
	CHECK_STATUS
	
	field_t *field = nil;
	if(ignoreCode) {
		field = (field_t*) malloc(sizeof(field_t));
		field->code = NULL;
		field->rawCode = nil;
	} else {
		// Read code
		statusCode readCodeStatus = prepareReadFieldCode(doc);
		if(readCodeStatus) return readCodeStatus;
		
		NSString* rawCode = [sbCodeRange content];
		CHECK_STATUS
		
		if(rawCode && [rawCode isNotEqualTo:[NSNull null]]) {
			// Check that this field is a valid Zotero field
			for(unsigned short i=0; FIELD_PREFIXES[i] != NULL; i++) {
				NSRange range = [rawCode rangeOfString:FIELD_PREFIXES[i]];
				if(range.location != NSNotFound) {
					field = (field_t*) malloc(sizeof(field_t));
					
					// If field code is all caps, make sure text object isn't in
					// all caps mode
					if([rawCode isEqualToString:[rawCode uppercaseString]]) {
						IGNORING_SB_ERRORS_BEGIN
						WordFont *sbFontObject = [sbCodeRange fontObject];
						if([sbFontObject allCaps]) {
							[sbFontObject setAllCaps:NO];
							rawCode = [sbCodeRange content];
						}
						IGNORING_SB_ERRORS_END
					}
					
					// Get code
					NSUInteger rawCodeLength = [rawCode length];
					NSUInteger codeStart = range.location+range.length;
					NSUInteger codeLength = rawCodeLength-codeStart;
					
					// Sometimes there is a space at the end of the code, which
					// we ignore here
					if([rawCode characterAtIndex:(rawCodeLength-1)] == ' ') {
						codeLength--;
					}
					field->rawCode = rawCode;
					[rawCode retain];
					field->code = copyNSString([rawCode substringWithRange:
												NSMakeRange(codeStart,
															codeLength)]);
					
					break;
				}
			}
		}
	}
	
	if(field) {
		field->text = NULL;
		field->bookmarkName = NULL;
		field->bookmarkNameNS = nil;
		field->sbBookmark = nil;
		field->textLocation = -1;
		field->noteLocation = -1;
		field->entryIndex = entryIndex;
		field->noteType = noteType;
		field->adjacent = false;
		
		field->doc = doc;
		field->sbField = sbField;
		field->sbCodeRange = sbCodeRange;
		field->sbContentRange = [sbField resultRange];
		CHECK_STATUS;
		
		[field->sbField retain];
		[field->sbCodeRange retain];
		[field->sbContentRange retain];
		
		// Add field to field list for document
		addValueToList(field, &(doc->allocatedFieldsStart),
					   &(doc->allocatedFieldsEnd));
		
		*returnValue = field;
		return STATUS_OK;
	}
	
	*returnValue = NULL;
	return STATUS_OK;
}

// Unfortunately after starting to update fields in reverse order in
// https://github.com/zotero/zotero/commit/419953f478eb6181717216dcb0dcfc27c303c953
// citations in footnotes started breaking when any in-text fields need updating
// (such as bibliography). We'll be re-initializing the sb objects here if such breakage
// occurs. Sigh.
statusCode checkFieldIntegrity(field_t *field) {
	HANDLE_EXCEPTIONS_BEGIN
	if (field->sbBookmark || field->noteType == 0) return STATUS_OK;
	// Field order upset. Sometimes SB returns 0 for entry index, which is invalid since
	// all those indices are 1-indexed, indicating.. something?.. that the field is no longer in the doc?
	NSInteger entryIndex = getEntryIndex(field->doc, field->sbField);
	if (errorHasOccurred()) {
		clearError();
	}
	else if (field->entryIndex == entryIndex) return STATUS_OK;
	
	[field->sbField release];
	[field->sbCodeRange release];
	[field->sbContentRange release];

	if (field->noteType == NOTE_FOOTNOTE) {
		field->sbField = [[[field->doc->sbDoc getStoryRangeStoryType:WordWdStoryTypeFootnotesStory] fields]
						  objectAtIndex:(field->entryIndex-1)];
	} else {
		field->sbField = [[[field->doc->sbDoc getStoryRangeStoryType:WordWdStoryTypeEndnotesStory] fields]
				   objectAtIndex:(field->entryIndex-1)];
	}
	field->sbCodeRange = [field->sbField fieldCode];
	field->sbContentRange = [field->sbField resultRange];
	[field->sbField retain];
	[field->sbCodeRange retain];
	[field->sbContentRange retain];
	CHECK_STATUS
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Allocates a field structure based on a WordBookmark
statusCode initBookmark(document_t *doc, WordBookmark* sbBookmark,
						short noteType,
						NSString* bookmarkName, BOOL ignoreCode,
						field_t **returnValue) {
	field_t *field = nil;
	
	// Read code
	NSString* resolvedBookmarkName;
	if(bookmarkName) {
		resolvedBookmarkName = bookmarkName;
	} else {
		resolvedBookmarkName = [sbBookmark name];
	}
	
	[doc->lock lock];
	sbBookmark = [[doc->sbDoc bookmarks] objectWithName:resolvedBookmarkName];
	CHECK_STATUS_LOCKED(doc)
	[doc->lock unlock];
	
	if(ignoreCode) {
		field = (field_t*) malloc(sizeof(field_t));
		field->code = NULL;
	} else {
		// Check that this field is a valid Zotero field
		for(unsigned short i=0; BOOKMARK_PREFIXES[i] != NULL; i++) {
			NSRange range = [resolvedBookmarkName
							 rangeOfString:BOOKMARK_PREFIXES[i]];
			if(range.location != NSNotFound) {
				// Get code from preferences
				NSString* propertyValue;
				statusCode status = getProperty(doc, resolvedBookmarkName,
												&propertyValue);
				if(status) return status;
				
				// Check that preferences code is valid
				for(unsigned short i=0; BOOKMARK_PREFIXES[i] != NULL; i++) {
					NSRange range = [propertyValue
									 rangeOfString:BOOKMARK_PREFIXES[i]];
					if(range.location != NSNotFound) {
						field = (field_t*) malloc(sizeof(field_t));
						NSUInteger codeStart = range.location+range.length;
						NSUInteger codeLength = [propertyValue length]-codeStart;
						
						field->code = copyNSString([propertyValue
													substringWithRange:
													NSMakeRange(codeStart,
																codeLength)]);
						break;
					}
				}
				
				if(field) break;
			}
		}
		
	}
	
	if(field) {
		field->text = NULL;
		field->sbField = nil;
		field->entryIndex = -1;
		field->textLocation = -1;
		field->noteLocation = -1;
		field->rawCode = nil;
		field->adjacent = false;
		
		field->bookmarkNameNS = resolvedBookmarkName;
		field->bookmarkName = copyNSString(resolvedBookmarkName);
		field->doc = doc;
		field->sbBookmark = sbBookmark;
		
		// Get the field contents
		field->sbCodeRange = [sbBookmark textObject];
		CHECK_STATUS;
		field->sbContentRange = field->sbCodeRange;
		
		if(noteType == -1) {
			WordWdStoryType storyType = [field->sbContentRange storyType];
			CHECK_STATUS
			if(storyType == WordWdStoryTypeFootnotesStory) {
				field->noteType = NOTE_FOOTNOTE;
			} else if(storyType == WordWdStoryTypeEndnotesStory) {
				field->noteType = NOTE_ENDNOTE;
			} else {
				field->noteType = 0;
			}
		} else {
			field->noteType = noteType;
		}
		
		[field->bookmarkNameNS retain];
		[field->sbBookmark retain];
		[field->sbCodeRange retain];
		[field->sbContentRange retain];
		
		// Add field to field list for document
		addValueToList(field, &(doc->allocatedFieldsStart),
					   &(doc->allocatedFieldsEnd));
		
		*returnValue = field;
		return STATUS_OK;
	}
	
	*returnValue = NULL;
	return STATUS_OK;
}

// Frees a field struct. This does not free the corresponding field list node.
void freeField(field_t* field) {
	if(field->text) free(field->text);
	if(field->code) free(field->code);
	if(field->bookmarkName) free(field->bookmarkName);
	[field->bookmarkNameNS release];
	[field->sbBookmark release];
	[field->sbField release];
	[field->sbCodeRange release];
	[field->sbContentRange release];
	[field->rawCode release];
	free(field);
}

// Deletes this field from the document
statusCode deleteField(field_t* field) {
	HANDLE_EXCEPTIONS_BEGIN
	ENSURE_OK(checkFieldIntegrity(field));
	short offset = field->sbField ? 1 : 0;
	if(field->noteType == NOTE_FOOTNOTE) {
		WordFootnote* note = [[field->sbContentRange footnotes]
							  objectAtIndex:0];
		if(([[note textObject] startOfContent]
				== [field->sbCodeRange startOfContent]-offset)
				&& ([[note textObject] endOfContent]
					== [field->sbContentRange endOfContent]+offset)) {
			WordWdStoryType storyType = [[field->doc->sbApp selection] storyType];
			CHECK_STATUS
			if (storyType == WordWdStoryTypeFootnotesStory) {
				[[field->doc->sbDoc createRangeStart:[[note noteReference] endOfContent] end:[[note noteReference] endOfContent]] sendEvent:'misc' id:'slct' parameters:'\00\00\00\00', nil];
				CHECK_STATUS
				storeCursorLocation(field->doc);
				field->doc->cursorMoved = NO;
			}
			[note delete];
			CHECK_STATUS
			return STATUS_OK;
		}
		CHECK_STATUS
	} else if(field->noteType == NOTE_ENDNOTE){ 
		WordEndnote* note = [[field->sbContentRange endnotes]
							 objectAtIndex:0];
		if(([[note textObject] startOfContent]
				== [field->sbCodeRange startOfContent]-offset)
			   && ([[note textObject] endOfContent]
				   == [field->sbContentRange endOfContent]+offset)) {
			WordWdStoryType storyType = [[field->doc->sbApp selection] storyType];
			CHECK_STATUS
			if (storyType == WordWdStoryTypeEndnotesStory) {
				[[field->doc->sbDoc createRangeStart:[[note noteReference] endOfContent] end:[[note noteReference] endOfContent]] sendEvent:'misc' id:'slct' parameters:'\00\00\00\00', nil];
				CHECK_STATUS
				storeCursorLocation(field->doc);
				field->doc->cursorMoved = NO;
			}
		   [note delete];
		   CHECK_STATUS
		   return STATUS_OK;
	   }
		CHECK_STATUS
	}
	
	if(field->sbBookmark) {
		[field->sbContentRange setContent:@""];
		CHECK_STATUS
	} else {
		[field->sbField delete];
		CHECK_STATUS
	}
	
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Removes a field code
statusCode removeCode(field_t* field) {
	HANDLE_EXCEPTIONS_BEGIN
	ENSURE_OK(checkFieldIntegrity(field));
	if(field->sbBookmark) {
		[field->sbBookmark delete];
	} else {
		[field->sbField sendEvent:'sWRD' id:'2489'
					   parameters:'\000\000\000\000', nil];
	}
	CHECK_STATUS
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Selects this field
statusCode selectField(field_t* field) {
	HANDLE_EXCEPTIONS_BEGIN
	ENSURE_OK(checkFieldIntegrity(field));
	[field->sbContentRange sendEvent:'misc' id:'slct' parameters:'\00\00\00\00', nil];
	CHECK_STATUS
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Sets text of this field
statusCode setText(field_t* field, const char string[], bool isRich) {
	HANDLE_EXCEPTIONS_BEGIN
	ENSURE_OK(checkFieldIntegrity(field));
	return setTextRaw(field, string, isRich);
	HANDLE_EXCEPTIONS_END
}

// Raw version of setText that allows the bookmark not to be deleted after
// citation insert
statusCode setTextRaw(field_t* field, const char string[], bool isRich) {
	BOOL locked = field->sbBookmark || isRich;
	BOOL restoreSelectionToEnd = NO;
	WordTextRange* insertRange;
	WordFont* font;
	double oldFontSize;
	NSString *oldFontName, *oldFontOtherName;
	WordWdColorIndex oldColorIndex;
	
	if (locked) {
		[(field->doc)->lock lock];
	}
	
	if(field->sbBookmark) {
		restoreSelectionToEnd = [[(field->doc)->sbApp selection]
								 selectionEnd] == [field->sbCodeRange endOfContent];
		CHECK_STATUS_LOCKED(field->doc);
	}
	
	// Store font info
	if (isRich) {
		IGNORING_SB_ERRORS_BEGIN
		// Make sure temp bookmark is gone
		[[[(field->doc)->sbDoc bookmarks] objectWithName:@ RTF_TEMP_BOOKMARK]
		 delete];
		
		// Save properties
		font = [field->sbContentRange fontObject];
		oldFontSize = [font fontSize];
		oldFontName = [font name];
		oldFontOtherName = [font otherName];
		oldColorIndex = [font colorIndex];
		IGNORING_SB_ERRORS_END
	} else {
		// Clear any superscripting. Unlike other font parameters,
		// superscripted fields are far more likely to come from a style change
		// than to be desired by the user.
		IGNORING_SB_ERRORS_BEGIN
		WordFont* font = [field->sbContentRange fontObject];
		[font setSuperscript:NO];
		IGNORING_SB_ERRORS_END
	}
	
	// Set text
	if (field->sbField) {
		if (isRich) {
			insertRange = field->sbContentRange;
			
			// Put RTF into the clipboard
			storePasteboardItems();
			replacePasteboardContentsWithRTF(string);
			
			// Paste RTF
			ENSURE_OK_LOCKED(field->doc, pasteStupid(field->doc, insertRange))
			
			// Restore clipboard contents and only then check for errors
			restorePasteboardContents();
			CHECK_STATUS_LOCKED(field->doc)
		}
		else {
			[field->sbContentRange
			 setContent:[NSString stringWithUTF8String:string]];
			CHECK_STATUS_LOCKED(field->doc)
		}
	}
	else {
		// Find a reference point in the appropriate story
		WordTextRange* referenceRange;
		if(field->noteType == NOTE_FOOTNOTE) {
			referenceRange = [[[(field->doc)->sbDoc footnotes]
							   objectAtIndex:
								   getEntryIndex(field->doc,
												 [[field->sbContentRange
												   footnotes]
												  objectAtIndex:0])-1]
							  textObject];
			CHECK_STATUS_LOCKED(field->doc)
		} else if(field->noteType == NOTE_ENDNOTE) {
			referenceRange = [[[(field->doc)->sbDoc endnotes]
							   objectAtIndex:
								   getEntryIndex(field->doc,
												 [[field->sbContentRange
												   endnotes]
												  objectAtIndex:0])-1]
							  textObject];
			CHECK_STATUS_LOCKED(field->doc)
		} else {
			referenceRange = [(field->doc)->sbDoc textObject];
			CHECK_STATUS_LOCKED(field->doc)
		}
		
		// Get positions
		NSInteger oldStart = [field->sbBookmark startOfBookmark];
		CHECK_STATUS_LOCKED(field->doc)
		NSInteger oldEnd = [field->sbBookmark endOfBookmark];
		CHECK_STATUS_LOCKED(field->doc)
		NSInteger oldStoryEnd = [referenceRange endOfContent];
		CHECK_STATUS_LOCKED(field->doc)
		
		// Set text (deletes bookmark)
		if (isRich) {
			// Put RTF into the clipboard
			storePasteboardItems();
			replacePasteboardContentsWithRTF(string);
			
			// Paste RTF
			ENSURE_OK_LOCKED(field->doc, pasteStupid(field->doc, field->sbContentRange))
			// Restore clipboard contents and only then check for errors
			restorePasteboardContents();
			CHECK_STATUS_LOCKED(field->doc)
		}
		else {
			[field->sbContentRange
			 setContent:[NSString stringWithUTF8String:string]];
			CHECK_STATUS_LOCKED(field->doc)
			
		}
			
		// Fix bookmark start and end
		NSInteger newEnd = oldEnd + [referenceRange endOfContent]
		- oldStoryEnd;
		CHECK_STATUS_LOCKED(field->doc)
		if (field->noteType) {
			// This approach to resetting the bookmark works everywhere,
			// but may have trouble with tables
			ENSURE_OK_LOCKED(field->doc,
							 insertFieldRaw(field->doc, "Bookmark", 0,
											referenceRange,
											field->bookmarkNameNS, NULL));
			[field->sbBookmark setStartOfBookmark:oldStart];
			CHECK_STATUS_LOCKED(field->doc)
			[field->sbBookmark setEndOfBookmark:newEnd];
			CHECK_STATUS_LOCKED(field->doc)
		} else {
			// This is a more appropriate way of setting range text, but
			// only works within the main story
			WordTextRange* newRange = [(field->doc)->sbDoc
									   createRangeStart:oldStart
									   end:newEnd];
			ENSURE_OK_LOCKED(field->doc,
							 insertFieldRaw(field->doc, "Bookmark", 0,
											newRange,
											field->bookmarkNameNS, NULL));
		}
	}
	
	// Restore font info
	if (isRich) {
		// Set style
		IGNORING_SB_ERRORS_BEGIN
		if(strncmp(field->code, "BIBL", 4) == 0) {
			[field->sbContentRange setStyle:WordWdBuiltinStyleStyleBibliography];
		}
		
		// Set properties back to saved
		font = [field->sbContentRange fontObject];
		[font setFontSize:oldFontSize];
		[font setName:oldFontName];
		[font setOtherName:oldFontOtherName];
		[font setColorIndex:oldColorIndex];
		IGNORING_SB_ERRORS_END
	}
	
	// If selection was at end of mark, put it there again
	if (restoreSelectionToEnd) {
		[[(field->doc)->sbApp selection] setSelectionStart:
		 [field->sbCodeRange endOfContent]];
		CHECK_STATUS_LOCKED(field->doc)
		[[[(field->doc)->sbApp selection] fontObject] reset];
		CHECK_STATUS_LOCKED(field->doc)
		field->doc->cursorMoved = NO;
	}
	
	if(locked) {
		[(field->doc)->lock unlock];
	}
	
	if(field->text) free(field->text);
	field->text = NULL;
	return STATUS_OK;
}

// Gets text inside this field
statusCode getText(field_t* field, char** returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	if (!field->text) {
		[(field->doc)->lock lock];
		statusCode status = prepareReadFieldCode(field->doc);
		[(field->doc)->lock unlock];
		ENSURE_OK(status)
		
		NSString* content = [field->sbContentRange content];
		// If a field is empty, Word for Mac will return missing value instead
		// of an empty string
		if(content == nil) {
			content = [NSString string];
		}
		
		field->text = copyNSString(content);
		CHECK_STATUS
	}
	
	*returnValue = field->text;
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

statusCode getCode(field_t* field, char** returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	*returnValue = field->code;
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Sets the field code
statusCode setCode(field_t *field, const char code[]) {
	HANDLE_EXCEPTIONS_BEGIN
	ENSURE_OK(checkFieldIntegrity(field));
	if(field->sbBookmark) {
		[(field->doc)->lock lock];
		NSString* rawCode = [NSString stringWithFormat:@"%@%@ ",
							 BOOKMARK_PREFIXES[0],
							 [NSString stringWithUTF8String:code]];
		statusCode status = setProperty(field->doc, field->bookmarkNameNS,
										rawCode);
		ENSURE_OK_LOCKED(field->doc, status)
		[(field->doc)->lock unlock];
	} else {
		if(field->rawCode != nil) {
			NSString* currentCode = [field->sbCodeRange content];
			if([currentCode isNotEqualTo:field->rawCode]) {
				DIE(@"Document modified during update")
			}
		}
		NSString* rawCode = [NSString stringWithFormat:@"%@%@ ",
							 FIELD_PREFIXES[0],
							 [NSString stringWithUTF8String:code]];
		[field->sbCodeRange setContent:rawCode];
		CHECK_STATUS

		[field->rawCode release];
		[rawCode retain];
		field->rawCode = rawCode;
	}
	
	// Store code in struct
	if(field->code) {
		free(field->code);
	}
	size_t codeLength = strlen(code)+1;
	field->code = (char*) malloc(codeLength);
	memcpy(field->code, code, codeLength);
	
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Returns the index of the note in which this field resides
statusCode getNoteIndex(field_t* field, unsigned long *returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	ENSURE_OK(checkFieldIntegrity(field));
	if(field->noteType == NOTE_FOOTNOTE) {
		WordFootnote* note = [[field->sbContentRange footnotes]
							  objectAtIndex:0];
		*returnValue = getEntryIndex(field->doc, note);
	} else if(field->noteType == NOTE_ENDNOTE){ 
		WordEndnote* note = [[field->sbContentRange endnotes]
							 objectAtIndex:0];
		*returnValue = getEntryIndex(field->doc, note);
	} else {
		*returnValue = 0;
	}
	
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Returns whether the next field is adjacent to this field
statusCode isAdjacentToNextField(field_t* field, bool *returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	*returnValue = field->adjacent;
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

statusCode equals(field_t *a, field_t *b, bool *returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	ENSURE_OK(checkFieldIntegrity(a));
	ENSURE_OK(checkFieldIntegrity(b));
	*returnValue = false;
	if (a->bookmarkName != NULL && b->bookmarkName != NULL) {
		*returnValue = (strcmp(a->bookmarkName, b->bookmarkName) == 0);
	}
	else {
		*returnValue = a->noteType == b->noteType && a->entryIndex == b->entryIndex;
	}
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Compares two fields to determine which comes before which
statusCode compareFields(field_t* a, field_t* b, short *returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	ENSURE_OK(checkFieldIntegrity(a));
	ENSURE_OK(checkFieldIntegrity(b));
	statusCode status;
	status = ensureTextLocationSet(a);
	if(status) return status;
	status = ensureTextLocationSet(b);
	if(status) return status;
	
	if(a->textLocation < b->textLocation) {
		*returnValue = -1;
		return STATUS_OK;
	} else if(b->textLocation < a->textLocation) {
		*returnValue = 1;
		return STATUS_OK;
	}
	
	// Compare positions inside a footnote
	if(a->noteType && b->noteType) {
		status = ensureNoteLocationSet(a);
		if(status) return status;
		status = ensureNoteLocationSet(b);
		if(status) return status;
		
		if(a->noteLocation < b->noteLocation) {
			*returnValue = -1;
			return STATUS_OK;
		} else if(b->noteLocation < a->noteLocation) {
			*returnValue = 1;
			return STATUS_OK;
		}
	}
	
	*returnValue = 0;
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

// Adapts the method signature for compareFields to work with the libc qsort_r
// function
int compareFieldsQsort(void* status, const void* a, const void* b) {
	if(*((statusCode*) status)) return 0;
	
	short returnValue;
	*((statusCode*) status) = compareFields(*((field_t **) a),
											*((field_t **) b), &returnValue);
	
	return returnValue;
}

// Make sure textLocation is set in the field struct
statusCode ensureTextLocationSet(field_t* field) {
	if(field->textLocation == -1) {
		if(field->noteType == NOTE_FOOTNOTE) {
			WordFootnote* note = [[field->sbContentRange footnotes]
								  objectAtIndex:0];
			CHECK_STATUS
			field->textLocation = [[note noteReference] startOfContent];
			CHECK_STATUS
		} else if(field->noteType == NOTE_ENDNOTE){ 
			WordEndnote* note = [[field->sbContentRange endnotes]
								 objectAtIndex:0];
			CHECK_STATUS
			field->textLocation = [[note noteReference] startOfContent];
			CHECK_STATUS
		} else {
			field->textLocation = [field->sbContentRange startOfContent];
			CHECK_STATUS
		}
	}
	
	return STATUS_OK;
}

// Make sure noteLocation is set in the field struct
statusCode ensureNoteLocationSet(field_t* field) {
	if(field->noteLocation == -1) {
		field->noteLocation = [field->sbContentRange startOfContent];
		CHECK_STATUS
	}
	return STATUS_OK;
}
