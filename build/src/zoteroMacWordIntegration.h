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

#ifndef zoteroMacWordIntegration_h
#define zoteroMacWordIntegration_h

#include "Word.h"
#define STATUS_OK 0
#define STATUS_EXCEPTION 1
#define STATUS_EXCEPTION_ALREADY_DISPLAYED 2

#define DIALOG_ICON_STOP 0
#define DIALOG_ICON_NOTICE 1
#define DIALOG_ICON_CAUTION 2
#define DIALOG_BUTTONS_OK 0
#define DIALOG_BUTTONS_OK_CANCEL 1
#define DIALOG_BUTTONS_YES_NO 2
#define DIALOG_BUTTONS_YES_NO_CANCEL 3
#define NOTE_FOOTNOTE 1
#define NOTE_ENDNOTE 2

#define MAX_PROPERTY_LENGTH 255
#define FIELD_PLACEHOLDER @"{Citation}"
#define BOOKMARK_REFERENCE_PROPERTY @"ZOTERO_BREF"
#define RTF_TEMP_BOOKMARK "ZOTERO_TEMP_BOOKMARK"
#define PREFS_PROPERTY "ZOTERO_PREF"

#define CHECK_STATUS \
if(errorHasOccurred()) {\
	flagError(__FUNCTION__, __FILE__, __LINE__-1);\
	return STATUS_EXCEPTION;\
}

#define DIE(x) \
{ throwError(x, __FUNCTION__, __FILE__, __LINE__-1);\
return STATUS_EXCEPTION; }

#define IGNORING_SB_ERRORS_BEGIN setErrorMonitor(false);
#define IGNORING_SB_ERRORS_END setErrorMonitor(true);

typedef struct Document {
	char* wordPath;
	bool isWord2004;
	WordApplication *sbApp;
	WordDocument *sbDoc;
	WordView *sbView;
	SBElementArray *sbProperties;
	
	BOOL restoreFullScreenMode;
	BOOL statusFullScreenMode;
	
	BOOL restoreInsertionsAndDeletions;
	BOOL statusInsertionsAndDeletions;
	
	BOOL restoreFormatChanges;
	BOOL statusFormatChanges;
} document_t;

typedef struct Field {
	// The field code
	char* code;
	
	// The field text
	char *text;
	
	// Self-evident
	WordApplication* sbApp;
	WordField* sbField;
	WordDocument* sbDoc;
	
	// The range corresponding to the field code, for a field
	WordTextRange* sbCodeRange;
	
	// The range corresponding to the content of a field
	WordTextRange *sbContentRange;
	
	// The note type (0, NOTE_FOOTNOTE, or NOTE_ENDNOTE)
	unsigned short noteType;
	
	// The index of this footnote in its parent collection (main doc, footnotes,
	// or endnotes)
	long entryIndex;
	
	// The location of this field relative to the start of the main body text.
	// For a footnote, this would be the position of the superscripted note
	// reference.
	long textLocation;
	
	// The location of this field relative to the start of the footnote/endnote
	// story.
	long noteLocation;
} field_t;

typedef struct FieldListNode {
	field_t* field;
	struct FieldListNode* next;
} fieldListNode_t;

typedef unsigned short statusCode;

// utilities.m
@class ZoteroSBApplicationDelegate;
@interface ZoteroSBApplicationDelegate : NSObject <SBApplicationDelegate>
@end
BOOL errorHasOccurred(void);
void setErrorMonitor(BOOL status);
void flagError(const char file[], const char function[], unsigned int line);
void throwError(NSString *errorString, const char file[], const char function[],
				unsigned int line);
void clearError(void);
char* getError(void);

FILE* getTemporaryFile(void);
void deleteTemporaryFile(void);
NSString* getTemporaryFilePath(void);
NSString* posixPathToHFSPath(NSString *posixPath);

char* copyNSString(NSString* string);
NSMutableString* escapeString(const char string[]);
NSString* generateRandomString(NSUInteger length);

// application.m
statusCode getDocument(bool isWord2004, const char* wordPath,
					   const char* documentName, document_t** returnValue);
statusCode displayAlert(char const dialogText[], unsigned short icon,
						unsigned short buttons, unsigned short* returnValue);

// document.m
@interface FieldGetter : NSObject
- (void)initWithDocument:(document_t)doc;
- (void)getFields:(id)param;
@end

void freeDocument(document_t *doc);
statusCode activate(document_t *doc);
statusCode canInsertField(document_t *doc, const char fieldType[],
						  bool* returnValue);
statusCode cursorInField(document_t *doc, const char fieldType[],
						 field_t** returnValue);
statusCode getDocumentData(document_t *doc, char **returnValue);
statusCode setDocumentData(document_t *doc, const char documentData[]);
statusCode insertField(document_t *doc, const char fieldType[],
					   unsigned short noteType, field_t **returnValue);
statusCode getFields(document_t *doc, const char fieldType[],
					 fieldListNode_t** returnNode);
void freeFieldList(fieldListNode_t *fieldList);
statusCode setBibliographyStyle(document_t *doc, long firstLineIndent, 
								long bodyIndent, unsigned long lineSpacing,
								unsigned long entrySpacing, long tabStops[],
								unsigned long tabStopCount);
statusCode cleanup(document_t *doc);

statusCode insertFieldRaw(document_t *doc, const char fieldType[],
						  unsigned short noteType, NSString* bookmarkName,
						  field_t **returnValue);
statusCode getProperty(document_t *doc, const char propertyName[],
					   char** returnValue);
statusCode setProperty(document_t *doc, const char propertyName[],
					   const char propertyValue[]);
statusCode prepareReadFieldCode(document_t *doc);

// field.m
void freeFields(field_t* fields[], unsigned long nFields);
statusCode deleteField(field_t* field);
statusCode removeCode(field_t* field);
statusCode selectField(field_t* field);
statusCode setText(field_t* field, const char string[], bool isRich);
statusCode getText(field_t* field, char** returnValue);
statusCode setCode(field_t *field, const char code[]);
statusCode getNoteIndex(field_t* field, unsigned long *returnValue);

statusCode initField(document_t *doc, WordField* sbField, short noteType,
					 NSInteger entryIndex, BOOL ignoreCode,
					 field_t **returnValue);
statusCode initBookmark(document_t *doc, WordBookmark* sbBookmark, BOOL ignoreCode,
						field_t **returnValue);
statusCode compareFields(field_t* a, field_t* b, short *returnValue);
statusCode setTextRaw(field_t* field, const char string[], bool isRich,
					  BOOL deleteBM);
statusCode ensureTextLocationSet(field_t* field);
statusCode ensureNoteLocationSet(field_t* field);

// install.m
statusCode install(const char zoteroDotPath[]);

#endif
