//
//  ZoteroWordIntegrationServiceClient.h
//  libZoteroMacWordIntegration
//
//  Created by Adomas Venckauskas on 2021-01-06.
//

#ifndef ZoteroWordIntegrationServiceClient_h
#define ZoteroWordIntegrationServiceClient_h

#import <AppKit/AppKit.h>
#import <sys/sysctl.h>
#include "XPCZoteroWordIntegration.h"

#define HANDLE_REPLY(status, doc) \
replyStatus = status; \
if (replyStatus != STATUS_OK) { \
	getRemoteErrorString(doc); \
	return; \
}

#define REPLY_BLOCK(doc) \
^(statusCode status) { \
HANDLE_REPLY(status, doc); \
}

#define HANDLE_EXCEPTIONS_BEGIN \
@try {

#define HANDLE_EXCEPTIONS_END \
} @catch(NSException* e) { \
	throwError([NSString stringWithFormat:@"%@", e], __FUNCTION__, [@__FILE__ lastPathComponent], __LINE__); \
	return STATUS_EXCEPTION; \
} @finally {\
	if(errorHasOccurred()) {\
		flagError(__FUNCTION__, [@__FILE__ lastPathComponent], __LINE__-1);\
		return STATUS_EXCEPTION;\
	} \
}

typedef struct RemoteDocument {
	NSXPCConnection *connection;
	NSObject <ZoteroWordIntegrationServiceProtocol> *remoteObject;
	
	listNode_t* allocatedFieldListsStart;
	listNode_t* allocatedFieldListsEnd;
} RemoteDocument;

// utility
int isRosetta(void);
bool isWordArm(void);
void addValueToList(void* value, listNode_t** listStart, listNode_t** listEnd);
void freeFieldList(listNode_t* fieldList);

// remoteUtility
statusCode getRemoteErrorString(RemoteDocument *doc);

// application
statusCode getDocument(int wordVersion, const char* wordPath,
						const char* documentName, bool ignoreArmIsSupported, RemoteDocument** returnValue);

// document
statusCode activate(RemoteDocument *doc);
statusCode displayAlert(RemoteDocument *doc, char const dialogText[],
						unsigned short icon, unsigned short buttons,
						unsigned short* returnValue);
statusCode canInsertField(RemoteDocument *doc, const char fieldType[],
						bool* returnValue);
statusCode cursorInField(RemoteDocument *doc, const char fieldType[],
						XPCField* returnValue);
statusCode getDocumentData(RemoteDocument *doc, char **returnValue);
statusCode setDocumentData(RemoteDocument *doc, const char documentData[]);
statusCode insertField(RemoteDocument *doc, const char fieldType[],
						unsigned short noteType, XPCField *returnValue);
statusCode getFields(RemoteDocument *doc, const char fieldType[],
						listNode_t** returnNode);
statusCode getFieldsAsync(RemoteDocument *doc, const char fieldType[],
						listNode_t** returnNode,
						void (*onProgress)(int progress));
statusCode convert(RemoteDocument *doc, XPCField fields[], unsigned long nFields,
						const char toFieldType[], unsigned short noteType[]);
statusCode setBibliographyStyle(RemoteDocument *doc, long firstLineIndent,
						long bodyIndent, unsigned long lineSpacing,
						unsigned long entrySpacing, long tabStops[],
						unsigned long tabStopCount);
statusCode exportDocument(RemoteDocument *doc, const char fieldType[],
						const char importInstructions[]);
statusCode importDocument(RemoteDocument *doc, const char fieldType[], bool *returnValue);
statusCode cleanup(RemoteDocument *doc);
statusCode complete(RemoteDocument *doc);

// field
statusCode deleteField(RemoteDocument *doc, XPCField field);
statusCode removeCode(RemoteDocument *doc, XPCField field);
statusCode selectField(RemoteDocument *doc, XPCField field);
statusCode setText(RemoteDocument *doc, XPCField field, const char string[], bool isRich);
statusCode getText(RemoteDocument *doc, XPCField field, char** returnValue);
statusCode setCode(RemoteDocument *doc, XPCField field, const char code[]);
statusCode getCode(RemoteDocument *doc, XPCField field, char** returnValue);
statusCode getNoteIndex(RemoteDocument *doc, XPCField field, unsigned long *returnValue);
statusCode equals(RemoteDocument *doc, XPCField a, XPCField b, bool *returnValue);

// install.m
statusCode install(const char zoteroDotPath[], const char zoteroDotmPath[]);
statusCode getScriptItemsDirectory(char** scriptFolder);
statusCode writeScript(char* scriptPath, char* scriptContent);


#endif /* ZoteroWordIntegrationServiceClient_h */
