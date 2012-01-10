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
				statusCode status = initBookmark(doc, sbBookmark, NO,
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
	return getProperty(doc, PREFS_PROPERTY, returnValue);
}

// Sets document data
statusCode setDocumentData(document_t *doc, const char documentData[]) {
	return setProperty(doc, PREFS_PROPERTY, documentData);
}

// Gets a property from the document. Remember to free the result!
statusCode getProperty(document_t *doc, const char propertyName[],
					   char **returnValue) {
	NSInteger i = 1;
	NSMutableString* stringComponents = [NSMutableString
										 stringWithCapacity:512];
	NSString* propertyValue = nil;
	do {
		NSString* currentPropertyName = [NSString stringWithFormat:@"%s_%d",
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
	
	*returnValue = copyNSString(stringComponents);
	return STATUS_OK;
}

// Stores a property in the document.
statusCode setProperty(document_t *doc, const char propertyName[],
					   const char propertyValue[]) {
	NSString *propertyValueNS = [NSString stringWithUTF8String:propertyValue];
	NSUInteger propertyValueLength = [propertyValueNS length];
	NSUInteger numberOfProperties = ceil(((float) propertyValueLength)
										 /MAX_PROPERTY_LENGTH);
	
	// Set fields with value
	for(NSUInteger i=0; i<numberOfProperties; i++) {
		NSString *currentPropertyName = [NSString stringWithFormat:@"%s_%d",
										 propertyName, i+1];
		
		NSString *currentPropertyValue;
		if(i == numberOfProperties-1) {
			currentPropertyValue = [propertyValueNS substringFromIndex:
									i*MAX_PROPERTY_LENGTH];
		} else {
			currentPropertyValue = [propertyValueNS substringWithRange:
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
	NSString* bookmarkName;
	
	if(strcmp(fieldType, "Bookmark") == 0) {
		if(!noteType) {
			[[doc->sbApp selection] setContent:FIELD_PLACEHOLDER];
			CHECK_STATUS
		}
		bookmarkName = [NSString stringWithFormat:@"%@_%@",
						BOOKMARK_REFERENCE_PROPERTY,
						generateRandomString(21)];
	} else {
		bookmarkName = nil;
	}
	
	return insertFieldRaw(doc, fieldType, noteType, bookmarkName, returnValue);
}

// Makes a field at the specified range. Uses the specified bookmark name for
// a bookmark.
statusCode insertFieldRaw(document_t *doc, const char fieldType[],
						  unsigned short noteType, NSString* bookmarkName,
						  field_t** returnValue) {
	WordField* sbField;
	
	// field
	NSAppleEventDescriptor* w170 = [NSAppleEventDescriptor
									descriptorWithTypeCode:'w170'];
	
	// field type:field print date
	NSAppleEventDescriptor* wFtP = [NSAppleEventDescriptor recordDescriptor];
	[wFtP setDescriptor:[NSAppleEventDescriptor
						 descriptorWithEnumCode:WordE183FieldPrintDate]
			 forKeyword:'wFtP'];
	
	WordTextRange* sbWhere = [[doc->sbApp selection] textObject];
	CHECK_STATUS
	WordE160 storyType = [sbWhere storyType];
	CHECK_STATUS
	
	// This is ridiculous. Scripting Bridge can't create the AppleEvents
	// we need, so we have to use AppleScript.
	if(storyType == WordE160MainTextStory && noteType) {
		NSString* noteTypeCode;
		if(noteType == NOTE_FOOTNOTE) {
			noteTypeCode = @"w156";
		} else if(noteType == NOTE_ENDNOTE) {
			noteTypeCode = @"w157";
		} else {
			DIE(@"Invalid field type");
		}
		
		// make new footnote/endnote at active document
		WordFootnote* sbNote = [doc->sbApp sendEvent:'core' id:'crel'
										  parameters:'kocl', noteTypeCode,
								'insh', doc->sbDoc, nil];
		/*script = [NSString stringWithFormat:
				  @"tell application \"%u\" to make new %@ at active document "
				  "with properties {text object:text object of selection}",
				  escapedWordPath, noteTypeString];
		scriptObject = [[NSAppleScript alloc] initWithSource:script];
		result = [scriptObject executeAndReturnError:nil];
		[scriptObject release];*/
		
		// Clear range content
		WordTextRange *sbNoteTextObject = [sbNote textObject];
		CHECK_STATUS
		[sbNoteTextObject setContent:@""];
		CHECK_STATUS
		
		// Move selection end past new footnote
		//WordTextRange *sbNoteReference = [sbNote noteReference];
		CHECK_STATUS
		//[sbWhere setSelectionStart:[sbNoteReference endOfContent]];
		CHECK_STATUS
		//[sbWhere setSelectionEnd:[sbNoteReference endOfContent]];
		CHECK_STATUS
		
		// Insert field
		
		// make new field at %@ with properties {field type:field print date}
		sbField = [doc->sbApp sendEvent:'core' id:'crel' parameters:'kocl',
				  w170, 'insh', sbNoteTextObject, 'prdt', wFtP, nil];
		CHECK_STATUS
		/*script = [NSString stringWithFormat:
				  @"tell application \"%u\" to make new field at %@ %d of "
				  "active document with properties "
				  "{field type:field print date}",
				  escapedWordPath, noteTypeString, noteIndex];
		scriptObject = [[NSAppleScript alloc] initWithSource:script];
		result = [scriptObject executeAndReturnError:nil];
		[scriptObject release];*/
	} else {
		// make new field at selection with properties
		// {field type:field print date}
		sbField = [doc->sbApp sendEvent:'core' id:'crel' parameters:'kocl',
				  w170, 'insh', sbWhere, 'prdt', wFtP, nil];
		CHECK_STATUS
	}
	
	NSInteger entryIndex;
	if(storyType == WordE160MainTextStory) {
		WordTextRange* tmpRange = [doc->sbDoc
								   createRangeStart:([sbWhere startOfContent]-1)
								   end:([sbWhere endOfContent]+1)];
		CHECK_STATUS
		entryIndex = [[[tmpRange fields] objectAtIndex:0] entry_index];
		CHECK_STATUS
		sbField = [[doc->sbDoc fields] objectAtIndex:(entryIndex-1)];
		CHECK_STATUS
	} else {
		sbField = [[sbWhere fields] objectAtIndex:1];
		CHECK_STATUS
		entryIndex = [sbField entry_index];
	}
	
	// TODO bookmarks
	return initField(doc, sbField, noteType, entryIndex, YES, returnValue);
}

// Gets fields
statusCode getFields(document_t *doc, const char fieldType,
					 fieldListNode_t** returnNode) {
	// Get fields in main body, footnotes, and endnotes
	SBElementArray* fieldCollections[] = {[doc->sbDoc fields],
		[[doc->sbDoc getStoryRangeStoryType:WordE160FootnotesStory] fields],
		[[doc->sbDoc getStoryRangeStoryType:WordE160EndnotesStory] fields]};
	CHECK_STATUS
	
	long currentFieldIndices[] = {0, 0, 0};
	field_t* currentFields[] = {NULL, NULL, NULL};
	fieldListNode_t* fieldListStart = NULL;
	fieldListNode_t* fieldListEnd = NULL;
	
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
				
				if(returnValue) {
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
	
	*returnNode = fieldListStart;
	return STATUS_OK;
}

// Sets the "Bibliography" style
statusCode setBibliographyStyle(document_t *doc, long firstLineIndent, 
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

