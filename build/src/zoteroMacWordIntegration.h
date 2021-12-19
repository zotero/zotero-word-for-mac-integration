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
#include "XPCZoteroWordIntegration.h"


enum DIALOG_ICON {
	DIALOG_ICON_STOP = 0,
	DIALOG_ICON_NOTICE = 1,
	DIALOG_ICON_CAUTION = 2
};

enum DIALOG_BUTTONS {
	DIALOG_BUTTONS_OK = 0,
	DIALOG_BUTTONS_OK_CANCEL = 1,
	DIALOG_BUTTONS_YES_NO = 2,
	DIALOG_BUTTONS_YES_NO_CANCEL = 3
};

#define MAX_PROPERTY_LENGTH 255
#define FIELD_PLACEHOLDER "{Citation}"
#define BOOKMARK_REFERENCE_PROPERTY @"ZOTERO_BREF"
#define RTF_TEMP_BOOKMARK "ZOTERO_TEMP_BOOKMARK"
#define PREFS_PROPERTY @"ZOTERO_PREF"
#define BOOKMARK_PREFIX = "ZOTERO_"
#define IMPORT_LINK_URL @"https://www.zotero.org/"
#define IMPORT_ITEM_PREFIX @"ITEM CSL_CITATION "
#define IMPORT_BIBL_PREFIX @"BIBL "
#define IMPORT_DOC_PREFS_PREFIX @"DOCUMENT_PREFERENCES "
#define MAIN_FIELD_PREFIX @" ADDIN ZOTERO_"


// Checks to see that the last Scripting Bridge call succeeded. If not, returns
// STATUS_EXCEPTION.
#define CHECK_STATUS \
if(errorHasOccurred()) {\
	flagError(__FUNCTION__, [@__FILE__ lastPathComponent], __LINE__-1);\
	if (	getErrorCode() == -1743) {\
		return STATUS_EXCEPTION_SB_DENIED;\
	}\
	return STATUS_EXCEPTION;\
}

// Checks an OSStatus. If the OSStatus is false, returns.
#define CHECK_OSSTATUS(expr) \
{ OSStatus statusToEnsure = expr; \
if(statusToEnsure) {\
return flagOSError(statusToEnsure, __FUNCTION__, [@__FILE__ lastPathComponent], __LINE__-1);\
} }

// Same as CHECK_STATUS, but also unlocks an NSLock on (document_t*)doc before
// returning.
#define CHECK_STATUS_LOCKED(doc) \
if(errorHasOccurred()) {\
	[(doc)->lock unlock];\
	flagError(__FUNCTION__, [@__FILE__ lastPathComponent], __LINE__-1);\
	if (	getErrorCode() == -1743) {\
		return STATUS_EXCEPTION_SB_DENIED;\
	}\
	return STATUS_EXCEPTION;\
}

// Returns expr if expr is non-zero.
#define ENSURE_OK(expr) \
{ \
statusCode statusToEnsure = expr; \
if(statusToEnsure != STATUS_OK) return statusToEnsure; \
}

// If y is non-zero, unlocks the lock on (document_t*)doc and then returns y.
#define ENSURE_OK_LOCKED(doc, expr) \
{ \
statusCode statusToEnsure = expr; \
if(statusToEnsure) {\
	[(doc)->lock unlock];\
	return statusToEnsure;\
} \
}

// Unlocks the lock on (document_t*)doc and then returns.
#define RETURN_STATUS_LOCKED(doc, expr) \
{ [(doc)->lock unlock];\
return expr; }

// Sets an error code doc and then returns STATUS_EXCEPTION.
#define DIE(err) \
{ throwError(err, __FUNCTION__, [@__FILE__ lastPathComponent], __LINE__-1);\
return STATUS_EXCEPTION; }

#define IGNORING_SB_ERRORS_BEGIN setErrorMonitor(false);
#define IGNORING_SB_ERRORS_END setErrorMonitor(true);

#define HANDLE_EXCEPTIONS_BEGIN \
@try {

#define HANDLE_EXCEPTIONS_END \
} @catch(NSException* e) { \
throwError([NSString stringWithFormat:@"%@", e], __FUNCTION__, [@__FILE__ lastPathComponent], __LINE__); \
return STATUS_EXCEPTION; \
}

typedef struct Document {
	char* wordPath;
	int wordVersion;
	WordApplication* sbApp;
	WordDocument* sbDoc;
	WordView* sbView;
	SBElementArray* sbProperties;
	
	BOOL restoreFullScreenMode;
	
	BOOL restoreInsertionsAndDeletions;
	BOOL statusInsertionsAndDeletions;
	
	BOOL restoreFormatChanges;
	BOOL statusFormatChanges;
	
	BOOL restoreTrackRevisions;
	BOOL statusTrackRevisions;
	
	WordE160 restoreNoteType;
	NSInteger restoreNote, restoreCursorEnd, restoreFieldIdx;
	BOOL cursorMoved, shouldRestoreCursor;
	NSInteger insertTextIntoNote;
	
	listNode_t* allocatedFieldsStart;
	listNode_t* allocatedFieldsEnd;
	listNode_t* allocatedFieldListsStart;
	listNode_t* allocatedFieldListsEnd;
	
	NSRecursiveLock* lock;
} document_t;

typedef struct Field {
	// The field code
	char* code;
	
	// The field text
	char* text;
	
	// The note type (0, NOTE_FOOTNOTE, or NOTE_ENDNOTE)
	unsigned short noteType;
	
	// The index of this footnote in its parent collection (main doc, footnotes,
	// or endnotes)
	long entryIndex;
	
	// The bookmark name
	char* bookmarkName;
	
	// Only one of these will be set
	WordField* sbField;
	WordBookmark* sbBookmark;
	
	// The bookmark name, as an NSString
	NSString* bookmarkNameNS;
	
	// The corresponding document
	document_t* doc;
	
	// The range corresponding to the field code, for a field
	WordTextRange* sbCodeRange;
	
	// The range corresponding to the content of a field
	WordTextRange *sbContentRange;
	
	// The location of this field relative to the start of the main body text.
	// For a footnote, this would be the position of the superscripted note
	// reference.
	NSInteger textLocation;
	
	// The location of this field relative to the start of the footnote/endnote
	// story.
	NSInteger noteLocation;
	
	// The raw field code for fields only
	NSString* rawCode;
} field_t;

// utilities.m
void preventAppNap(void);
void allowAppNap(void);

void storeCursorLocation(document_t* doc);
statusCode moveCursorOutOfNote(document_t* doc);
statusCode restoreCursor(document_t* doc);

FILE* getTemporaryFile(document_t *doc);
void deleteTemporaryFile(void);
NSString* getTemporaryFilePath(void);

NSString* generateRandomString(NSUInteger length);
NSInteger getEntryIndex(document_t* x, SBObject* y);

// application.m
statusCode getDocument(int wordVersion, const char* wordPath,
					   const char* documentName, document_t** returnValue);

// document.m
void freeDocument(document_t *doc);
statusCode activate(document_t *doc);
statusCode displayAlert(document_t *doc, char const dialogText[],
						unsigned short icon, unsigned short buttons,
						unsigned short* returnValue);
statusCode canInsertField(document_t *doc, const char fieldType[],
						  bool* returnValue);
statusCode cursorInField(document_t *doc, const char fieldType[],
						 field_t** returnValue);
statusCode getDocumentData(document_t *doc, char **returnValue);
statusCode setDocumentData(document_t *doc, const char documentData[]);
statusCode insertField(document_t *doc, const char fieldType[],
					   unsigned short noteType, field_t **returnValue);
statusCode getFields(document_t *doc, const char fieldType[],
					 listNode_t** returnNode);
statusCode getFieldsAsync(document_t *doc, const char fieldType[],
						  listNode_t** returnNode,
						  void (*onProgress)(int progress));
statusCode convert(document_t *doc, field_t* fields[], unsigned long nFields,
				   const char toFieldType[], unsigned short noteType[]);
statusCode setBibliographyStyle(document_t *doc, long firstLineIndent, 
								long bodyIndent, unsigned long lineSpacing,
								unsigned long entrySpacing, long tabStops[],
								unsigned long tabStopCount);
statusCode exportDocument(document_t *doc, const char fieldType[],
						const char importInstructions[]);
statusCode importDocument(document_t *doc, const char fieldType[], bool *returnValue);
statusCode insertText(document_t *doc, const char htmlString[]);
statusCode convertPlaceholdersToFields(document_t *doc, const char* placeholders[],
									   const unsigned long nPlaceholders ,const unsigned short noteType,
									   const char fieldType[], listNode_t** returnNode);
statusCode cleanup(document_t *doc);
statusCode complete(document_t *doc);

statusCode getProperty(document_t *doc, NSString* propertyName,
					   NSString** returnValue);
statusCode setProperty(document_t *doc, NSString* propertyName,
					   NSString* propertyValue);
statusCode prepareReadFieldCode(document_t *doc);
statusCode insertFieldRaw(document_t *doc, const char fieldType[],
						  unsigned short noteType, WordTextRange *sbWhere,
						  NSString* bookmarkName, field_t** returnValue);
void addValueToList(void* value, listNode_t** listStart, listNode_t** listEnd);

// field.m
void freeField(field_t* field);
statusCode deleteField(field_t* field);
statusCode removeCode(field_t* field);
statusCode selectField(field_t* field);
statusCode setText(field_t* field, const char string[], bool isRich);
statusCode getText(field_t* field, char** returnValue);
statusCode setCode(field_t *field, const char code[]);
statusCode getCode(field_t* field, char** returnValue);
statusCode getNoteIndex(field_t* field, unsigned long *returnValue);

statusCode initField(document_t *doc, WordField* sbField, short noteType,
					 NSInteger entryIndex, BOOL ignoreCode,
					 field_t **returnValue);
statusCode checkFieldIntegrity(field_t *field);
statusCode initBookmark(document_t *doc, WordBookmark* sbBookmark, short noteType,
						NSString* bookmarkName, BOOL ignoreCode,
						field_t **returnValue);
statusCode compareFields(field_t* a, field_t* b, short *returnValue);
statusCode equals(field_t *a, field_t *b, bool *returnValue);
int compareFieldsQsort(void* statusCode, const void* a, const void* b);
statusCode setTextRaw(field_t* field, const char string[], bool isRich,
					  BOOL deleteBM);
statusCode ensureTextLocationSet(field_t* field);
statusCode ensureNoteLocationSet(field_t* field);

// install.m
statusCode install(const char zoteroDotPath[], const char zoteroDotmPath[], const char zoteroScptPath[]);
statusCode getScriptItemsDirectory(char** scriptFolder);
statusCode writeScript(char* scriptPath, char* scriptContent);

#endif
