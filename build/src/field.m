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

NSString* FIELD_PREFIXES[] = {@" ADDIN ZOTERO_", @" CSL_", NULL};
NSString* BOOKMARK_PREFIXES[] = {@"ZOTERO_", @"CSL_", NULL};

// Allocates a field structure based on a WordField, optionall checking to make
// sure that the field code actually matches a Zotero field.
statusCode initField(document_t *doc, WordField* sbField, short noteType,
					 NSInteger entryIndex, BOOL ignoreCode,
					 field_t **returnValue) {
	WordTextRange* sbCodeRange = [sbField fieldCode];
	CHECK_STATUS
	
	field_t *field = nil;
	if(ignoreCode) {
		field = (field_t*) malloc(sizeof(field_t));
		field->code = NULL;
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
					
					// Get code
					NSUInteger rawCodeLength = [rawCode length];
					NSUInteger codeStart = range.location+range.length;
					NSUInteger codeLength = rawCodeLength-codeStart;
					
					// Sometimes there is a space at the end of the code, which
					// we ignore here
					if([rawCode characterAtIndex:(rawCodeLength-1)] == ' ') {
						codeLength--;
					}
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
		
		field->doc = doc;
		field->sbField = sbField;
		field->sbCodeRange = sbCodeRange;
		field->sbContentRange = [sbField resultRange];
		CHECK_STATUS;
		
		if(noteType == -1) {
			WordE160 storyType = [field->sbContentRange storyType];
			CHECK_STATUS
			if(storyType == WordE160FootnotesStory) {
				field->noteType = NOTE_FOOTNOTE;
			} else if(storyType == WordE160EndnotesStory) {
				field->noteType = NOTE_ENDNOTE;
			} else {
				field->noteType = 0;
			}
		} else {
			field->noteType = noteType;
		}
		
		if(entryIndex == -1) {
			field->entryIndex = [field->sbField entry_index];
			CHECK_STATUS
		} else {
			field->entryIndex = entryIndex;
		}
		
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
		
		field->bookmarkNameNS = resolvedBookmarkName;
		field->bookmarkName = copyNSString(resolvedBookmarkName);
		field->doc = doc;
		field->sbBookmark = sbBookmark;
		
		// Get the field contents
		field->sbCodeRange = [sbBookmark textObject];
		CHECK_STATUS;
		field->sbContentRange = field->sbCodeRange;
		
		if(noteType == -1) {
			WordE160 storyType = [field->sbContentRange storyType];
			CHECK_STATUS
			if(storyType == WordE160FootnotesStory) {
				field->noteType = NOTE_FOOTNOTE;
			} else if(storyType == WordE160EndnotesStory) {
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
	free(field);
}

// Deletes this field
statusCode deleteField(field_t* field) {
	short offset = field->sbField ? 1 : 0;
	if(field->noteType == NOTE_FOOTNOTE) {
		WordFootnote* note = [[field->sbContentRange footnotes]
							  objectAtIndex:0];
		if(([[note textObject] startOfContent]
			== [field->sbCodeRange startOfContent]-offset)
			&& ([[note textObject] endOfContent]
				== [field->sbContentRange endOfContent]+offset)) {
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
}

// Removes a field code
statusCode removeCode(field_t* field) {
	if(field->sbBookmark) {
		[field->sbBookmark delete];
	} else {
		[(field->doc)->lock lock];
		[(field->doc)->sbApp unlink:field->sbField];
		CHECK_STATUS_LOCKED(field->doc)
		[(field->doc)->lock unlock];
	}
	CHECK_STATUS
	return STATUS_OK;
}

// Selects this field
statusCode selectField(field_t* field) {
	[field->sbContentRange select];
	CHECK_STATUS
	return STATUS_OK;
}

// Sets text of this field
statusCode setText(field_t* field, const char string[], bool isRich) {
	return setTextRaw(field, string, isRich, YES);
}

// Raw version of setText that allows the bookmark not to be deleted after
// citation insert
statusCode setTextRaw(field_t* field, const char string[], bool isRich,
					  BOOL deleteBM) {
	if(isRich) {
		[(field->doc)->lock lock];
		
		IGNORING_SB_ERRORS_BEGIN
		// Make sure temp bookmark is gone
		[[[(field->doc)->sbDoc bookmarks] objectWithName:@ RTF_TEMP_BOOKMARK]
		 delete];
		
		// Save properties
		WordFont* font = [field->sbContentRange fontObject];
		double oldFontSize = [font fontSize];
		NSString* oldFontName = [font name];
		NSString* oldFontOtherName = [font otherName];
		WordE110 oldColorIndex = [font colorIndex];
		IGNORING_SB_ERRORS_END
		
		WordBookmark* tempBookmark;
		BOOL selectionAtEnd = NO;
		const char* bookmarkName;
		if(field->sbBookmark) {
			// Rename bookmark to a temporary name
			statusCode status = insertFieldRaw(field->doc, "Bookmark", 0,
											   field->sbCodeRange,
											   @ RTF_TEMP_BOOKMARK, NULL);
			if(status) return status;
			[field->sbBookmark delete];
			CHECK_STATUS_LOCKED(field->doc)
			tempBookmark = [[(field->doc)->sbDoc bookmarks]
							objectWithName:@ RTF_TEMP_BOOKMARK];
			CHECK_STATUS_LOCKED(field->doc)
			
			// Check if cursor is at end of citation
			selectionAtEnd = [[(field->doc)->sbApp selection] selectionEnd]
			== [field->sbCodeRange endOfContent];
			CHECK_STATUS_LOCKED(field->doc)
			
			// We are going to insert bookmark with the proper name
			bookmarkName = field->bookmarkName;
		} else {
			// We are going to insert a temporary bookmark
			bookmarkName = RTF_TEMP_BOOKMARK;
			
			// Clear content range
			[field->sbContentRange setContent:@""];
			CHECK_STATUS_LOCKED(field->doc)
		}
		
		// Write RTF to a file
		size_t newStringSize = strlen(string)-6;
		char* newString = (char*) malloc(newStringSize);
		strlcpy(newString, string+6, newStringSize);
		FILE* temporaryFile = getTemporaryFile();
		fprintf(temporaryFile, "{\\rtf {\\bkmkstart %s}%s{\\bkmkend %s}}",
				bookmarkName, newString, bookmarkName);
		fflush(temporaryFile);
		free(newString);
		
		// Insert file
		NSString* temporaryFilePath = getTemporaryFilePath();
		[(field->doc)->sbApp insertFileAt:field->sbContentRange
								 fileName:posixPathToHFSPath(temporaryFilePath)
								fileRange:[NSString
										   stringWithUTF8String:bookmarkName]
					   confirmConversions:NO
									 link:NO];
		CHECK_STATUS_LOCKED(field->doc)
		
		if(field->sbBookmark) {
			// Delete temporary bookmark text
			[[tempBookmark textObject] setContent:@""];
		}
		
		if(deleteBM) {
			IGNORING_SB_ERRORS_BEGIN
			[[[field->sbContentRange bookmarks]
			  objectWithName:@ RTF_TEMP_BOOKMARK] delete];
			IGNORING_SB_ERRORS_END
		}
		
		// Set style
		IGNORING_SB_ERRORS_BEGIN
		if(strncmp(field->code, "BIBL", 4) == 0) {
			[[field->sbContentRange propertyWithCode:'1695']
			 setTo:@"Bibliography"];
		}
		
		// Set properties back to saved
		font = [field->sbContentRange fontObject];
		[font setFontSize:oldFontSize];
		[font setName:oldFontName];
		[font setOtherName:oldFontOtherName];
		[font setColorIndex:oldColorIndex];
		IGNORING_SB_ERRORS_END
		
		// If selection was at end of mark, put it there again
		if(selectionAtEnd) {
			[[(field->doc)->sbApp selection] setSelectionStart:
			 [field->sbCodeRange endOfContent]];
			CHECK_STATUS_LOCKED(field->doc)
			[[[(field->doc)->sbApp selection] fontObject] reset];
			CHECK_STATUS_LOCKED(field->doc)
		}
		
		[(field->doc)->lock unlock];
	} else {
		if(field->sbBookmark) {
			[(field->doc)->lock lock];
			// Find a reference point in the appropriate story
			WordTextRange* referenceRange;
			if(field->noteType == NOTE_FOOTNOTE) {
				referenceRange = [[[(field->doc)->sbDoc footnotes] objectAtIndex:
								   [[[field->sbContentRange footnotes]
									 objectAtIndex:0] entry_index]-1] textObject];
				CHECK_STATUS_LOCKED(field->doc)
			} else if(field->noteType == NOTE_ENDNOTE) {
				referenceRange = [[[(field->doc)->sbDoc endnotes] objectAtIndex:
								   [[[field->sbContentRange endnotes]
									 objectAtIndex:0] entry_index]-1] textObject];
				CHECK_STATUS_LOCKED(field->doc)
			} else {
				referenceRange = [(field->doc)->sbDoc textObject];
				CHECK_STATUS_LOCKED(field->doc)
			}
			
			// Get some data about this bookmark
			NSInteger oldStart = [field->sbBookmark startOfBookmark];
			CHECK_STATUS_LOCKED(field->doc)
			NSInteger oldEnd = [field->sbBookmark endOfBookmark];
			CHECK_STATUS_LOCKED(field->doc)
			NSInteger oldStoryEnd = [referenceRange endOfContent];
			CHECK_STATUS_LOCKED(field->doc)
			
			// Set text (deletes bookmark)
			[field->sbContentRange
			 setContent:[NSString stringWithUTF8String:string]];
			CHECK_STATUS_LOCKED(field->doc)
			
			// Regenerate bookmark (an empty bookmark wouldn't have gotten
			// erased)
			if(oldStart != oldEnd) {
				insertFieldRaw(field->doc, "Bookmark", 0, referenceRange,
							   field->bookmarkNameNS, NULL);
			}
			
			// Move around text
			[field->sbBookmark setStartOfBookmark:oldStart];
			CHECK_STATUS_LOCKED(field->doc)
			[field->sbBookmark setEndOfBookmark:
			 (oldEnd+[referenceRange endOfContent]-oldStoryEnd)];
			CHECK_STATUS_LOCKED(field->doc)
			
			[(field->doc)->lock unlock];
		} else {
			[field->sbContentRange
			 setContent:[NSString stringWithUTF8String:string]];
			CHECK_STATUS
		}
	}
	if(field->text) free(field->text);
	field->text = NULL;
	return STATUS_OK;
}

// Gets text inside this field
statusCode getText(field_t* field, char** returnValue) {
	[(field->doc)->lock lock];
	
	statusCode status = prepareReadFieldCode(field->doc);
	ENSURE_OK_LOCKED(field->doc, status)
	if(!field->text) {
		field->text = copyNSString([field->sbContentRange content]);
		CHECK_STATUS_LOCKED(field->doc);
	}
	
	[(field->doc)->lock unlock];
	*returnValue = field->text;
	return STATUS_OK;
}

// Sets the field code
statusCode setCode(field_t *field, const char code[]) {
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
		NSString* rawCode = [NSString stringWithFormat:@"%@%@ ",
							 FIELD_PREFIXES[0],
							 [NSString stringWithUTF8String:code]];
		[field->sbCodeRange setContent:rawCode];
		CHECK_STATUS
	}
	
	// Store code in struct
	if(field->code) {
		free(field->code);
	}
	size_t codeLength = strlen(code)+1;
	field->code = (char*) malloc(codeLength);
	memcpy(field->code, code, codeLength);
	
	return STATUS_OK;
}

// Returns the index of the note in which this field resides
statusCode getNoteIndex(field_t* field, unsigned long *returnValue) {
	if(field->noteType == NOTE_FOOTNOTE) {
		WordFootnote* note = [[field->sbContentRange footnotes]
							  objectAtIndex:0];
		*returnValue = [note entry_index];
	} else if(field->noteType == NOTE_ENDNOTE){ 
		WordEndnote* note = [[field->sbContentRange endnotes]
							 objectAtIndex:0];
		*returnValue = [note entry_index];
	} else {
		*returnValue = 0;
	}
	
	return STATUS_OK;
}

// Compares two fields to determine which comes before which
statusCode compareFields(field_t* a, field_t* b, short *returnValue) {
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
}

// Adapts the method signature for compareFields to work with the libc qsort_r
// function
int compareFieldsQsort(void* voidStatus, const void* a, const void* b) {
	statusCode* status = ((statusCode*) voidStatus);
	if(status) return 0;
	
	short returnValue;
	*status = compareFields(*((field_t **) a), *((field_t **) b), &returnValue);
	
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