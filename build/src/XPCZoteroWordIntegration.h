//
//  XPCZoteroWordIntegration.h
//  libZoteroMacWordIntegration
//
//  Created by Adomas Venckauskas on 2021-01-06.
//

#ifndef XPCZoteroWordIntegration_h
#define XPCZoteroWordIntegration_h

#import <Foundation/Foundation.h>
#import <AppKit/AppKit.h>
#import <ScriptingBridge/ScriptingBridge.h>
#include "ZoteroWordIntegrationServiceProtocol.h"

enum NOTE_TYPE {
	NOTE_FOOTNOTE = 1,
	NOTE_ENDNOTE = 2
};

typedef struct ListNode {
	void* value;
	struct ListNode* next;
} listNode_t;

// sharedUtilites.m
@class ZoteroSBApplicationDelegate;
@interface ZoteroSBApplicationDelegate : NSObject <SBApplicationDelegate>
@end
void (^handleXPCError)(NSError *);

NSString* posixPathToHFSPath(NSString *posixPath);
char* copyNSString(NSString* string);

BOOL errorHasOccurred(void);
void setErrorMonitor(BOOL status);
void flagError(const char function[], NSString* file, unsigned int line);
void clearError(void);
char *getError(void);
void setError(char *);
NSInteger getErrorCode(void);
void throwError(NSString *errorString, const char function[], NSString* file,
				unsigned int line);
statusCode flagOSError(OSStatus status, const char function[], NSString* file,
					   unsigned int line);

void freeData(void* ptr);

#endif /* XPCZoteroWordIntegration_h */
