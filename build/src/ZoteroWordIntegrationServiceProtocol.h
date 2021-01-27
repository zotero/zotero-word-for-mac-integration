//
//  ZoteroWordIntegrationServiceProtocol.h
//  ZoteroWordIntegrationService
//
//  Created by Adomas Venckauskas on 2021-01-06.
//

#import <Foundation/Foundation.h>

typedef unsigned short statusCode;

enum STATUS {
	STATUS_OK = 0,
	STATUS_EXCEPTION = 1,
	STATUS_EXCEPTION_ALREADY_DISPLAYED = 2,
	STATUS_EXCEPTION_SB_DENIED = 3,
	STATUS_EXCEPTION_ARM_NOT_SUPPORTED = 4,
	STATUS_EXCEPTION_ARM_SUPPORTED = 5
};

// This is actually a pointer, but if we try to pass a void pointer (or even NSValue with a pointer) via
// XPC the framework seems to try to do some verification of the memory or something
// and we get communication interruptions with no useful error messages (4097)
// so we pass it as a long. Ha! Hacks!
typedef unsigned long XPCField;

// The protocol that this service will vend as its API. This header file will also need to be visible to the process hosting the service.
@protocol ZoteroWordIntegrationServiceProtocol

// application
- (void)getDocumentWithWordVersion:(int)wordVersion wordPath:(const char*)wordPath
	documentName:(const char*)documentName andReply:(void (^)(statusCode))reply;

// document
- (void)activateWithReply:(void (^)(statusCode))reply;
- (void)displayAlertWithText:(char const *)dialogText
					icon:(unsigned short)icon buttons:(unsigned short)buttons
					andReply:(void (^)(statusCode, unsigned short))reply;
- (void)canInsertFieldWithFieldType:(const char *)fieldType
					andReply:(void (^)(statusCode, bool))reply;
- (void)cursorInFieldWithFieldType:(const char *)fieldType
					andReply:(void (^)(statusCode, XPCField))reply;
- (void)getDocumentDataWithReply:(void (^)(statusCode, char *))reply;
- (void)setDocumentDataWithDocumentData:(const char *)documentData
					andReply:(void (^)(statusCode))reply;
- (void)insertFieldWithFieldType:(const char *)fieldType
					noteType:(unsigned short)noteType
					andReply:(void (^)(statusCode, XPCField))reply;
- (void)getFieldsWithFieldType:(const char *)fieldType
					andReply:(void (^)(statusCode, NSArray *))reply;
- (void)convertFields:(NSArray *)fields toFieldType:(const char *)toFieldType
					withNoteTypes:(NSData *)noteTypeData
					andReply:(void (^)(statusCode))reply;
- (void)setBibliographyStyleWithFirstLineIndent:(long)firstLineIndent
					bodyIndent:(long)bodyIndent lineSpacing:(unsigned long)lineSpacing
					entrySpacing:(unsigned long)entrySpacing tabStops:(NSData *)tabStopsData
					tabStopCount:(unsigned long)tabStopCount
					andReply:(void (^)(statusCode))reply;
- (void)exportDocumentWithFieldType:(const char *)fieldType
					importInstructions:(const char *)importInstructions
					andReply:(void (^)(statusCode))reply;
- (void)importDocumentWithFieldType:(const char *)fieldType
					andReply:(void (^)(statusCode, bool))reply;
- (void)cleanupWithReply:(void (^)(statusCode))reply;
- (void)completeWithReply:(void (^)(statusCode))reply;

// field
- (void)deleteField:(XPCField)field withReply:(void (^)(statusCode))reply;
- (void)removeCode:(XPCField)field withReply:(void (^)(statusCode))reply;
- (void)selectField:(XPCField)field withReply:(void (^)(statusCode))reply;
- (void)setText:(XPCField)field to:(const char *)string isRich:(bool)isRich
					withReply:(void (^)(statusCode))reply;
- (void)getText:(XPCField)field withReply:(void (^)(statusCode, char *))reply;
- (void)setCode:(XPCField)field to:(const char *)code withReply:(void (^)(statusCode))reply;
- (void)getCode:(XPCField)field withReply:(void (^)(statusCode, char *))reply;
- (void)getNoteIndex:(XPCField)field withReply:(void (^)(statusCode, unsigned long))reply;
- (void)equals:(XPCField)a to:(XPCField)b withReply:(void (^)(statusCode, bool))reply;

// utilities
- (void)getErrorWithReply:(void (^)(statusCode, char *))reply;
    
@end

/*
 To use the service from an application or other process, use NSXPCConnection to establish a connection to the service by doing something like this:

     _connectionToService = [[NSXPCConnection alloc] initWithServiceName:@"zotero.ZoteroWordIntegrationService"];
     _connectionToService.remoteObjectInterface = [NSXPCInterface interfaceWithProtocol:@protocol(ZoteroWordIntegrationServiceProtocol)];
     [_connectionToService resume];

Once you have a connection to the service, you can use it like this:

     [[_connectionToService remoteObjectProxy] upperCaseString:@"hello" withReply:^(NSString *aString) {
         // We have received a response. Update our text field, but do it on the main thread.
         NSLog(@"Result string was: %@", aString);
     }];

 And, when you are finished with the service, clean up the connection like this:

     [_connectionToService invalidate];
*/
