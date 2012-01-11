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
		
		// Check that this field is a valid Zotero field
		for(unsigned short i=0; FIELD_PREFIXES[i] != NULL; i++) {
			NSRange range = [rawCode rangeOfString:FIELD_PREFIXES[i]];
			if(range.location != NSNotFound) {
				field = (field_t*) malloc(sizeof(field_t));
				
				// Get code
				NSUInteger rawCodeLength = [rawCode length];
				NSUInteger codeStart = range.location+range.length;
				NSUInteger codeLength = rawCodeLength-codeStart;
				
				// Sometimes there is a space at the end of the code, which we
				// ignore here
				if([rawCode characterAtIndex:(rawCodeLength-1)] == ' ') {
					codeLength--;
				}
				field->code = copyNSString([rawCode substringWithRange:
											NSMakeRange(codeStart, codeLength)]);
				
				break;
			}
		}
	}
	
	if(field) {
		field->text = NULL;
		
		field->sbApp = doc->sbApp;
		field->sbDoc = doc->sbDoc;
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
		
		field->textLocation = -1;
		field->noteLocation = -1;
		
		[field->sbDoc retain];
		[field->sbApp retain];
		[field->sbField retain];
		[field->sbCodeRange retain];
		[field->sbContentRange retain];
		
		*returnValue = field;
		return STATUS_OK;
	}
	
	*returnValue = NULL;
	return STATUS_OK;
}

// Allocates a field structure based on a WordBookmark
statusCode initBookmark(document_t *doc, WordBookmark* sbBookmark, BOOL ignoreCode,
						field_t **field) {
	// TODO
	DIE(@"Bookmarks not implemented");
}

// Frees a field struct. Does not alter the underlying field.
void freeFields(field_t* fields[], unsigned long nFields) {
	for(unsigned long i=0; i<nFields; i++) {
		field_t* field = fields[i];
		
		if(field->text) free(field->text);
		if(field->code) free(field->code);
		[field->sbApp release];
		[field->sbDoc release];
		[field->sbField release];
		[field->sbCodeRange release];
		[field->sbContentRange release];
		free(field);
	}
}

// Deletes this field
statusCode deleteField(field_t* field) {
	if(field->noteType == NOTE_FOOTNOTE) {
		WordFootnote* note = [[field->sbContentRange footnotes]
							  objectAtIndex:0];
		if(([[note textObject] startOfContent]
			== [field->sbCodeRange startOfContent]-1)
			&& ([[note textObject] endOfContent]
				== [field->sbContentRange endOfContent]+1)) {
				[note delete];
				CHECK_STATUS
				return STATUS_OK;
			}
		CHECK_STATUS
	} else if(field->noteType == NOTE_ENDNOTE){ 
		WordEndnote* note = [[field->sbContentRange endnotes]
							 objectAtIndex:0];
		if(([[note textObject] startOfContent]
			== [field->sbCodeRange startOfContent]-1)
		   && ([[note textObject] endOfContent]
			   == [field->sbContentRange endOfContent]+1)) {
			   [note delete];
			   CHECK_STATUS
			   return STATUS_OK;
		   }
		CHECK_STATUS
	}
	
	[field->sbField delete];
	
	return STATUS_OK;
}

// Removes a field code
statusCode removeCode(field_t* field) {
	[field->sbApp unlink:field->sbField];
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
		IGNORING_SB_ERRORS_BEGIN
		// Make sure temp bookmark is gone
		[[[field->sbDoc bookmarks] objectWithName:@ RTF_TEMP_BOOKMARK] delete];
		
		// Save properties
		WordFont* font = [field->sbContentRange fontObject];
		double oldFontSize = [font fontSize];
		NSString* oldFontName = [font name];
		NSString* oldFontOtherName = [font otherName];
		WordE110 oldColorIndex = [font colorIndex];
		IGNORING_SB_ERRORS_END
		
		// Write RTF to a file
		size_t newStringSize = strlen(string)-6;
		char* newString = (char*) malloc(newStringSize);
		strlcpy(newString, string+6, newStringSize);
		FILE* temporaryFile = getTemporaryFile();
		fprintf(temporaryFile,
				"{\\rtf {\\bkmkstart "RTF_TEMP_BOOKMARK"}%s{\\bkmkend "
				RTF_TEMP_BOOKMARK"}}", newString);
		fflush(temporaryFile);
		free(newString);
		
		// Insert file
		[field->sbContentRange setContent:@""];
		CHECK_STATUS
		[field->sbApp insertFileAt:field->sbContentRange
						  fileName:posixPathToHFSPath(getTemporaryFilePath())
						 fileRange:@ RTF_TEMP_BOOKMARK
				confirmConversions:NO
							  link:NO];
		CHECK_STATUS
		
		if(deleteBM) {
			[[[field->sbContentRange bookmarks]
			  objectWithName:@ RTF_TEMP_BOOKMARK] delete];
			CHECK_STATUS
		}
		
		// Set style
		IGNORING_SB_ERRORS_BEGIN
		if(strncmp(field->code, "BIBL", 4) == 0) {
			// XXX this probably doesn't work
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
	} else {
		[field->sbContentRange
		 setContent:[NSString stringWithUTF8String:string]];
		CHECK_STATUS
	}
	if(field->text) free(field->text);
	field->text = NULL;
	return STATUS_OK;
}

// Gets text inside this field
statusCode getText(field_t* field, char** returnValue) {
	if(!field->text) {
		field->text = copyNSString([field->sbContentRange content]);
		CHECK_STATUS
	}
	*returnValue = field->text;
	return STATUS_OK;
}

// Sets the field code
statusCode setCode(field_t *field, const char code[]) {
	// Set code
	NSString* rawCode = [NSString stringWithFormat:@"%@%@ ",
						 FIELD_PREFIXES[0],
						 [NSString stringWithUTF8String:code]];
	[field->sbCodeRange setContent:rawCode];
	CHECK_STATUS
	
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


// Make sure textLocation is set in the field struct
statusCode ensureTextLocationSet(field_t* field) {
	if(field->textLocation == -1) {
		if(field->noteType == NOTE_FOOTNOTE) {
			WordFootnote* note = [[field->sbContentRange footnotes]
								  objectAtIndex:0];
			CHECK_STATUS
			field->textLocation = [[note textObject] startOfContent];
			CHECK_STATUS
		} else if(field->noteType == NOTE_ENDNOTE){ 
			WordEndnote* note = [[field->sbContentRange endnotes]
								 objectAtIndex:0];
			CHECK_STATUS
			field->textLocation = [[note textObject] startOfContent];
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