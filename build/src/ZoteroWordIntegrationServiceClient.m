//
//  ZoteroWordIntegrationServiceClient.m
//  libZoteroMacWordIntegration
//
//  Created by Adomas Venckauskas on 2021-01-07.
//

#include "ZoteroWordIntegrationServiceClient.h"

statusCode replyStatus;

// Definitions of internal utility functions
void addValueToList(void* value, listNode_t** listStart, listNode_t** listEnd);
void freeFieldList(listNode_t* fieldList);

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


// Free a field list.
void freeFieldList(listNode_t* fieldList) {
	listNode_t* nextNode = fieldList;
	while(nextNode) {
		listNode_t* currentNode = nextNode;
		nextNode = currentNode->next;
		free(currentNode);
	}
}


// ----- XPC Calls begin
// remoteUtility

statusCode getRemoteErrorString(RemoteDocument *doc) {
	HANDLE_EXCEPTIONS_BEGIN
	__block statusCode thisReplyStatus;
	[doc->remoteObject getErrorWithReply:^(statusCode status, char *reply) {
		thisReplyStatus = thisReplyStatus;
		if (status != STATUS_OK) {
			throwError(@"Failed to retrieve the remote error string", __FUNCTION__, [@__FILE__ lastPathComponent], __LINE__);
			return;
		}
		setError(reply);
	}];
	return thisReplyStatus;
	HANDLE_EXCEPTIONS_END
}

// application
statusCode getDocument(int wordVersion, const char* wordPath,
					   const char* documentName, RemoteDocument** returnValue)
{
	HANDLE_EXCEPTIONS_BEGIN
	NSXPCConnection *xpcConnection;
	xpcConnection = [[NSXPCConnection alloc] initWithServiceName:@"zotero.ZoteroWordIntegrationService"];
	xpcConnection.remoteObjectInterface = [NSXPCInterface interfaceWithProtocol:@protocol(ZoteroWordIntegrationServiceProtocol)];
	[xpcConnection resume];
	*returnValue = malloc(sizeof(RemoteDocument));
	(*returnValue)->allocatedFieldListsStart = (*returnValue)->allocatedFieldListsEnd = NULL;
	(*returnValue)->connection = xpcConnection;
	
	NSObject <ZoteroWordIntegrationServiceProtocol> *remoteObject = [xpcConnection synchronousRemoteObjectProxyWithErrorHandler:handleXPCError];
	(*returnValue)->remoteObject = remoteObject;
	
	[remoteObject getDocumentWithWordVersion:wordVersion wordPath:wordPath documentName: documentName andReply: ^(statusCode status) {
		HANDLE_REPLY(status, *returnValue);
	}];

	[xpcConnection retain];
	[remoteObject retain];
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}

// document
statusCode activate(RemoteDocument *doc) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->remoteObject activateWithReply:REPLY_BLOCK(doc)];
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}
statusCode displayAlert(RemoteDocument *doc, char const dialogText[],
						unsigned short icon, unsigned short buttons,
						unsigned short* returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->remoteObject displayAlertWithText:dialogText icon:icon buttons:buttons
								   andReply:^(statusCode status, unsigned short reply) {
		HANDLE_REPLY(status, doc)
		*returnValue = reply;
	}];
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}
statusCode canInsertField(RemoteDocument *doc, const char fieldType[],
						  bool* returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->remoteObject canInsertFieldWithFieldType:fieldType andReply:^(statusCode status, bool reply) {
		HANDLE_REPLY(status, doc)
		*returnValue = reply;
	}];
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}
statusCode cursorInField(RemoteDocument *doc, const char fieldType[],
						 XPCField* returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->remoteObject cursorInFieldWithFieldType:fieldType andReply:^(statusCode status, XPCField field) {
		HANDLE_REPLY(status, doc)
		*returnValue = field;
	}];
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}
statusCode getDocumentData(RemoteDocument *doc, char **returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->remoteObject getDocumentDataWithReply:^(statusCode status, char *reply) {
		HANDLE_REPLY(status, doc)
		*returnValue = strdup(reply);
	}];
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}
statusCode setDocumentData(RemoteDocument *doc, const char documentData[]) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->remoteObject setDocumentDataWithDocumentData:documentData andReply:REPLY_BLOCK(doc)];
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}
statusCode insertField(RemoteDocument *doc, const char fieldType[],
					   unsigned short noteType, XPCField *returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->remoteObject insertFieldWithFieldType:fieldType noteType:noteType
									   andReply:^(statusCode status, XPCField field) {
		HANDLE_REPLY(status, doc)
		*returnValue = field;
	}];
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}
statusCode getFields(RemoteDocument *doc, const char fieldType[],
					 listNode_t** returnNode) {
	HANDLE_EXCEPTIONS_BEGIN
	__block listNode_t* fieldListStart = NULL;
	__block listNode_t* fieldListEnd = NULL;
	[doc->remoteObject getFieldsWithFieldType:fieldType andReply:^(statusCode status, NSArray *fields) {
		HANDLE_REPLY(status, doc)
		for (NSValue *value in fields) {
			XPCField field = [value nonretainedObjectValue];
			addValueToList(field, &fieldListStart, &fieldListEnd);
		}
	}];
	addValueToList(fieldListStart, &(doc->allocatedFieldListsStart), &(doc->allocatedFieldListsEnd));
	*returnNode = fieldListStart;
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}

statusCode getFieldsAsync(RemoteDocument *doc, const char fieldType[],
						  listNode_t** returnNode,
						  void (*onProgress)(int)) {
	HANDLE_EXCEPTIONS_BEGIN
	dispatch_queue_t currentQueue = dispatch_get_current_queue();
	id <ZoteroWordIntegrationServiceProtocol> remoteObject = [doc->connection remoteObjectProxyWithErrorHandler:^(NSError * _Nonnull error) {
		throwError([error localizedDescription], __FUNCTION__, [@__FILE__ lastPathComponent], __LINE__);
		dispatch_async(currentQueue, ^{
			getRemoteErrorString(doc);
			onProgress(-1);
		});
	}];
	__block listNode_t* fieldListStart = NULL;
	__block listNode_t* fieldListEnd = NULL;
	[remoteObject getFieldsWithFieldType:fieldType andReply:^(statusCode status, NSArray *fields) {
		if (status != STATUS_OK) {
			dispatch_async(currentQueue, ^{
				getRemoteErrorString(doc);
				onProgress(-1);
			});
			return;
		}
		for (NSData *data in fields) {
			XPCField field = *(XPCField *)[data bytes];
			addValueToList(field, &fieldListStart, &fieldListEnd);
		}
		addValueToList(fieldListStart, &(doc->allocatedFieldListsStart), &(doc->allocatedFieldListsEnd));
		*returnNode = fieldListStart;
		dispatch_async(currentQueue, ^{
			onProgress(100);
		});
	}];
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}

statusCode convert(RemoteDocument *doc, XPCField fields[], unsigned long nFields,
				   const char toFieldType[], unsigned short noteType[]) {
	HANDLE_EXCEPTIONS_BEGIN
	NSMutableArray *fieldArray = [NSMutableArray new];
	for (unsigned int i = 0; i < nFields; i++) {
		[fieldArray addObject:[NSData dataWithBytes:&(fields[i]) length:sizeof(XPCField)]];
	}
	[doc->remoteObject convertFields:fieldArray toFieldType:toFieldType withNoteTypes:[NSData dataWithBytes:noteType length:sizeof(unsigned short) * nFields] andReply:REPLY_BLOCK(doc)];
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}
statusCode setBibliographyStyle(RemoteDocument *doc, long firstLineIndent,
								long bodyIndent, unsigned long lineSpacing,
								unsigned long entrySpacing, long tabStops[],
								unsigned long tabStopCount) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->remoteObject setBibliographyStyleWithFirstLineIndent:firstLineIndent bodyIndent:bodyIndent lineSpacing:lineSpacing entrySpacing:entrySpacing tabStops:[NSData dataWithBytes:tabStops length:sizeof(long) * tabStopCount] tabStopCount:tabStopCount andReply:REPLY_BLOCK(doc)];
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}
statusCode exportDocument(RemoteDocument *doc, const char fieldType[],
						  const char importInstructions[]) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->remoteObject exportDocumentWithFieldType:fieldType importInstructions:importInstructions andReply:REPLY_BLOCK(doc)];
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}
statusCode importDocument(RemoteDocument *doc, const char fieldType[], bool *returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->remoteObject importDocumentWithFieldType:fieldType andReply:^(statusCode status, bool reply) {
		HANDLE_REPLY(status, doc);
		*returnValue = reply;
	}];
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}
statusCode insertText(RemoteDocument *doc, const char htmlString[]) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->remoteObject insertTextWithHtmlString:htmlString andReply:REPLY_BLOCK(doc)];
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}
statusCode convertPlaceholdersToFields(RemoteDocument *doc, const char* placeholders[],
									   const unsigned long nPlaceholders ,const unsigned short noteType,
									   const char fieldType[], listNode_t** returnNode) {
	HANDLE_EXCEPTIONS_BEGIN
	__block listNode_t* fieldListStart = NULL;
	__block listNode_t* fieldListEnd = NULL;
	[doc->remoteObject convertPlaceholdersToFieldsWithPlaceholders:placeholders placeholderCount:nPlaceholders noteType:noteType fieldType:fieldType andReply:^(statusCode status, NSArray *fields) {
		HANDLE_REPLY(status, doc)
		for (NSValue *value in fields) {
			XPCField field = [value nonretainedObjectValue];
			addValueToList(field, &fieldListStart, &fieldListEnd);
		}
	}];
	addValueToList(fieldListStart, &(doc->allocatedFieldListsStart), &(doc->allocatedFieldListsEnd));
	*returnNode = fieldListStart;
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}
statusCode cleanup(RemoteDocument *doc) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->remoteObject cleanupWithReply:REPLY_BLOCK(doc)];
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}
statusCode complete(RemoteDocument *doc) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->remoteObject completeWithReply:REPLY_BLOCK(doc)];
	// Free allocated field lists
	listNode_t* nextNode = doc->allocatedFieldListsStart;
	while(nextNode) {
		listNode_t* currentNode = nextNode;
		freeFieldList(currentNode->value);
		nextNode = currentNode->next;
		free(currentNode);
	}
	[doc->connection invalidate];
	[doc->remoteObject release];
	[doc->connection release];
	free(doc);
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}

// field
statusCode deleteField(RemoteDocument *doc, XPCField field) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->remoteObject deleteField:field withReply:REPLY_BLOCK(doc)];
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}
statusCode removeCode(RemoteDocument *doc, XPCField field) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->remoteObject removeCode:field withReply:REPLY_BLOCK(doc)];
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}
statusCode selectField(RemoteDocument *doc, XPCField field) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->remoteObject selectField:field withReply:REPLY_BLOCK(doc)];
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}
statusCode setText(RemoteDocument *doc, XPCField field, const char string[], bool isRich) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->remoteObject setText:field to:string isRich:isRich withReply:REPLY_BLOCK(doc)];
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}
statusCode getText(RemoteDocument *doc, XPCField field, char** returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->remoteObject getText:field withReply:^(statusCode status, char *reply) {
		HANDLE_REPLY(status, doc);
		*returnValue = strdup(reply);
	}];
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}
statusCode setCode(RemoteDocument *doc, XPCField field, const char code[]) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->remoteObject setCode:field to:code withReply:REPLY_BLOCK(doc)];
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}
statusCode getCode(RemoteDocument *doc, XPCField field, char** returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->remoteObject getCode:field withReply:^(statusCode status, char *reply) {
		HANDLE_REPLY(status, doc);
		*returnValue = strdup(reply);
	}];
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}
statusCode getNoteIndex(RemoteDocument *doc, XPCField field, unsigned long *returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->remoteObject getNoteIndex:field withReply:^(statusCode status, unsigned long reply) {
		HANDLE_REPLY(status, doc);
		*returnValue = reply;
	}];
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}
statusCode equals(RemoteDocument *doc, XPCField a, XPCField b, bool *returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->remoteObject equals:a to:b withReply:^(statusCode status, bool reply) {
		HANDLE_REPLY(status, doc);
		*returnValue = reply;
	}];
	return replyStatus;
	HANDLE_EXCEPTIONS_END
}
