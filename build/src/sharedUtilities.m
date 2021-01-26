//
//  sharedUtilities.m
//  libZoteroMacWordIntegration
//
//  Created by Adomas Venckauskas on 2021-01-07.
//

#import <sys/sysctl.h>
#include "XPCZoteroWordIntegration.h"


BOOL monitorErrors = true;
NSError* lastError = nil;
char* lastErrorString = NULL;


// Converts a Cocoa path to an HFS (colon-delimited) path
NSString* posixPathToHFSPath(NSString *posixPath) {
	CFURLRef url = CFURLCreateWithFileSystemPath(NULL, (CFStringRef)
												 posixPath,
												 kCFURLPOSIXPathStyle, NO);
	return (NSString *)CFURLCopyFileSystemPath(url, kCFURLHFSPathStyle);
}

// A delegate to keep track of the last error encountered
@implementation ZoteroSBApplicationDelegate
- (id)eventDidFail:(const AppleEvent *)event withError:(NSError *)error;
{
	if(monitorErrors) {
		clearError();
		lastError = error;
		[lastError retain];
	}
	return nil;
}
@end

void (^handleXPCError)(NSError *) = ^(NSError * _Nonnull error) {
	if(monitorErrors) {
		clearError();
		lastError = error;
		[lastError retain];
	}
};

// Copies an NSString to a malloc'd char*
char* copyNSString(NSString* string) {
	const char* utf8String = [string UTF8String];
	
	// Copy the string so it doesn't get autoreleased
	size_t stringSize = strlen(utf8String)+1;
	char* newString = (char*) malloc(stringSize);
	strlcpy(newString, utf8String, stringSize);
	return newString;
}

// Checks if an error has occurred
BOOL errorHasOccurred(void) {
	return lastError != nil;
}

// Sets status of the error monitor
void setErrorMonitor(BOOL status) {
	monitorErrors = status;
}

// Generates an error string
void flagError(const char function[], NSString* file, unsigned int line) {
	if(lastErrorString) free(lastErrorString);
	lastErrorString = copyNSString([NSString stringWithFormat:@"%@ @[%s:%@:%d]",
									[lastError localizedDescription],
									function, file, line]);
}

// Clears the last error encountered
void clearError(void) {
	if(lastError) {
		[lastError release];
		lastError = nil;
	}
	
	if(lastErrorString) {
		free(lastErrorString);
		lastErrorString = NULL;
	}
}

// Gets the last error encountered
char* getError(void) {
	return lastErrorString;
}

void setError(char *errorString) {
	clearError();
	lastErrorString = strdup(errorString);
}

NSInteger getErrorCode(void) {
	return [lastError code];
}

// Manually throws an exception up to JS
void throwError(NSString *errorString, const char function[], NSString* file,
				unsigned int line) {
	if(lastErrorString) free(lastErrorString);
	lastErrorString = copyNSString([NSString stringWithFormat:@"%@ @[%s:%@:%d]",
									errorString, function, file, line]);
}

// Gets a string for an authorization failure
statusCode flagOSError(OSStatus status, const char function[], NSString* file,
					   unsigned int line) {
	if(status == errAuthorizationCanceled) {
		return STATUS_EXCEPTION_ALREADY_DISPLAYED;
	}
	
	NSString* err = [[NSError errorWithDomain:NSOSStatusErrorDomain code:status
									 userInfo:nil] localizedDescription];
	throwError(err, function, file, line);
	return STATUS_EXCEPTION;
}

// Frees a C string
void freeData(void* ptr) {
	free(ptr);
}

// From https://developer.apple.com/forums/thread/653009 (https://archive.is/etgKB)
// Checks if Zotero is running as translated Rosetta 2 app (on an M1 Apple)
int isZoteroRosetta() {
	int ret = 0;
	size_t size = sizeof(ret);
	// Call the sysctl and if successful return the result
	if (sysctlbyname("sysctl.proc_translated", &ret, &size, NULL, 0) != -1)
		return ret;
	// If "sysctl.proc_translated" is not present then must be native
	if (errno == ENOENT)
		return 0;
	return -1;
}

char *getWordVersion(const char wordPath[]) {
	NSBundle *wordBundle = [NSBundle bundleWithPath:[NSString stringWithUTF8String:wordPath]];
	NSString *wordVersion = wordBundle.infoDictionary[@"CFBundleShortVersionString"];
	return copyNSString(wordVersion);
}

bool isWordArm() {
	NSArray *runningApplications = [[NSWorkspace sharedWorkspace] runningApplications];
	for (NSRunningApplication* app in runningApplications) {
		if ([[app bundleIdentifier] isEqual:(@"com.microsoft.Word")]) {
			return [app executableArchitecture] != NSBundleExecutableArchitectureX86_64;
		}
	}
	return false;
}
