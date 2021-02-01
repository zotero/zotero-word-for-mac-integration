//
//  ZoteroWordIntegrationService.m
//  ZoteroWordIntegrationService
//
//  Created by Adomas Venckauskas on 2021-01-06.
//

#import "ZoteroWordIntegrationService.h"

@implementation ZoteroWordIntegrationService

// application
- (void)getDocumentWithWordVersion:(int)wordVersion wordPath:(const char*)wordPath
					  documentName:(const char*)documentName andReply:(void (^)(statusCode))reply {
	document_t *doc;
	statusCode status = getDocument(wordVersion, wordPath, documentName, &(doc));
	self.doc = doc;
	reply(status);
}


// document
- (void)activateWithReply:(void (^)(statusCode))reply {
	reply(activate(self.doc));
}

- (void)displayAlertWithText:(char const *)dialogText
						icon:(unsigned short)icon buttons:(unsigned short)buttons
					andReply:(void (^)(statusCode, unsigned short))reply {
	unsigned short result;
	statusCode status = displayAlert(self.doc, dialogText, icon, buttons, &result);
	reply(status, result);
}

- (void)canInsertFieldWithFieldType:(const char *)fieldType
						   andReply:(void (^)(statusCode, bool))reply {
	bool result;
	statusCode status = canInsertField(self.doc, fieldType, &result);
	reply(status, result);
}

- (void)cursorInFieldWithFieldType:(const char *)fieldType
						  andReply:(void (^)(statusCode, XPCField))reply {
	XPCField result;
	statusCode status = cursorInField(self.doc, fieldType, (field_t **)&result);
	reply(status, result);
}

- (void)getDocumentDataWithReply:(void (^)(statusCode, char *))reply {
	char *result = NULL;
	statusCode status = getDocumentData(self.doc, &result);
	reply(status, result);
	free(result);
}

- (void)setDocumentDataWithDocumentData:(const char *)documentData
							   andReply:(void (^)(statusCode))reply {
	reply(setDocumentData(self.doc, documentData));
}

- (void)insertFieldWithFieldType:(const char *)fieldType
						noteType:(unsigned short)noteType
						andReply:(void (^)(statusCode, XPCField))reply {
	XPCField result;
	statusCode status = insertField(self.doc, fieldType, noteType, (field_t **)&result);
	reply(status, result);
}

- (void)getFieldsWithFieldType:(const char *)fieldType
					  andReply:(void (^)(statusCode, NSArray *))reply {
	NSMutableArray *fields = [NSMutableArray new];
	listNode_t *fieldList;
	statusCode status = getFields(self.doc, fieldType, &fieldList);
	if (status == STATUS_OK) {
		listNode_t* nextNode = fieldList;
		while (nextNode) {
			listNode_t* currentNode = nextNode;
			[fields addObject:[NSData dataWithBytes:&(currentNode->value) length:sizeof(field_t *)]];
			nextNode = currentNode->next;
		}
	}
	reply(status, fields);
}

- (void)convertFields:(NSArray *)fields toFieldType:(const char *)toFieldType
		withNoteTypes:(NSData *)noteTypeData
			 andReply:(void (^)(statusCode))reply {
	unsigned short *noteType = [noteTypeData bytes];
	field_t **fieldArray;
	unsigned long count = [fields count];
	fieldArray = malloc(sizeof(field_t *) * count);
	for (int i = 0; i < count; i++) {
		NSData *data = [fields objectAtIndex:i];
		fieldArray[i] = *(field_t **)[data bytes];
	}
	reply(convert(self.doc, fieldArray, count, toFieldType, noteType));
	free(fieldArray);
}

- (void)setBibliographyStyleWithFirstLineIndent:(long)firstLineIndent
									 bodyIndent:(long)bodyIndent lineSpacing:(unsigned long)lineSpacing
								   entrySpacing:(unsigned long)entrySpacing tabStops:(NSData *)tabStopsData
								   tabStopCount:(unsigned long)tabStopCount
									   andReply:(void (^)(statusCode))reply {
	long *tabStops = [tabStopsData bytes];
	reply(setBibliographyStyle(self.doc, firstLineIndent, bodyIndent, lineSpacing, entrySpacing, tabStops, tabStopCount));
}

- (void)exportDocumentWithFieldType:(const char *)fieldType
				 importInstructions:(const char *)importInstructions
						   andReply:(void (^)(statusCode))reply {
	reply(exportDocument(self.doc, fieldType, importInstructions));
}

- (void)importDocumentWithFieldType:(const char *)fieldType
						   andReply:(void (^)(statusCode, bool))reply {
	bool result;
	statusCode status = importDocument(self.doc, fieldType, &result);
	reply(status, result);
}

- (void)cleanupWithReply:(void (^)(statusCode))reply {
	reply(cleanup(self.doc));
}

- (void)completeWithReply:(void (^)(statusCode))reply {
	reply(complete(self.doc));
	freeDocument(self.doc);
}

// field
- (void)deleteField:(XPCField)field withReply:(void (^)(statusCode))reply {
	reply(deleteField((field_t *) field));
}

- (void)removeCode:(XPCField)field withReply:(void (^)(statusCode))reply {
	reply(removeCode((field_t *) field));
}

- (void)selectField:(XPCField)field withReply:(void (^)(statusCode))reply {
	reply(selectField((field_t *) field));
}

- (void)setText:(XPCField)field to:(const char *)string isRich:(bool)isRich
	  withReply:(void (^)(statusCode))reply {
	reply(setText((field_t *) field, string, isRich));
}

- (void)getText:(XPCField)field withReply:(void (^)(statusCode, char *))reply {
	char *result;
	statusCode status = getText((field_t *) field, &result);
	reply(status, result);
}

- (void)setCode:(XPCField)field to:(const char *)code withReply:(void (^)(statusCode))reply {
	reply(setCode((field_t *) field, code));
}

- (void)getCode:(XPCField)field withReply:(void (^)(statusCode, char *))reply {
	char *result;
	statusCode status = getCode((field_t *) field, &result);
	reply(status, result);
}

- (void)getNoteIndex:(XPCField)field withReply:(void (^)(statusCode, unsigned long))reply {
	unsigned long result;
	statusCode status = getNoteIndex((field_t *) field, &result);
	reply(status, result);
}

- (void)equals:(XPCField)a to:(XPCField)b withReply:(void (^)(statusCode, bool))reply {
	bool result;
	statusCode status = equals((field_t *) a, (field_t *) b, &result);
	reply(status, result);
}

// utilities
- (void)getErrorWithReply:(void (^)(statusCode, char *))reply {
	reply(STATUS_OK, getError());
}

@end
