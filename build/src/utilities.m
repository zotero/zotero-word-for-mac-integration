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

BOOL monitorErrors = true;
NSError* lastError = nil;
char* lastErrorString = NULL;

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

// Checks if an error has occurred
BOOL errorHasOccurred(void) {
	return lastError != nil;
}

// Sets status of the error monitor
void setErrorMonitor(BOOL status) {
	monitorErrors = status;
}

// Generates an error string
void flagError(const char file[], const char function[], unsigned int line) {
	if(lastErrorString) free(lastErrorString);
	lastErrorString = copyNSString([NSString stringWithFormat:@"%@ @[%s:%s:%d]",
									[lastError localizedDescription],
									file, function, line]);
}

// Gets a string for an authorization failure
statusCode flagOSError(OSStatus status, const char file[],
					   const char function[], unsigned int line) {
	if(status == errAuthorizationCanceled) {
		return STATUS_EXCEPTION_ALREADY_DISPLAYED;
	}
	
	NSString* err = [NSError errorWithDomain:NSOSStatusErrorDomain code:status
									userInfo:nil];
	throwError(err, file, function, line);
	return STATUS_EXCEPTION;
}

// Manually throws an exception up to JS
void throwError(NSString *errorString, const char file[], const char function[],
				unsigned int line) {
	if(lastErrorString) free(lastErrorString);
	lastErrorString = copyNSString([NSString stringWithFormat:@"%@ @[%s:%s:%d]",
									errorString, file, function, line]);
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

FILE* tempFile = NULL;
char* tempFileString = NULL;
NSString* tempFileStringNS = nil;

// Gets a FILE for the temporary file, truncating it to zero length
FILE* getTemporaryFile(void) {
	if(tempFile == NULL) {
		const char *tempFileTemplate = [[NSTemporaryDirectory()
										 stringByAppendingPathComponent:
										 @"zotero.XXXXXX.rtf"]
										fileSystemRepresentation];
		size_t tempFileLength = strlen(tempFileTemplate)+1;
		tempFileString = (char *)malloc(tempFileLength);
		strlcpy(tempFileString, tempFileTemplate, tempFileLength);
		int tempFileDescriptor = mkstemps(tempFileString, 4);
		tempFileStringNS = [[NSFileManager defaultManager]
					  stringWithFileSystemRepresentation:tempFileString
					  length:strlen(tempFileString)];
		[tempFileStringNS retain];
		tempFile = fdopen(tempFileDescriptor, "w");
	}
	rewind(tempFile);
	ftruncate(fileno(tempFile), 0);
	return tempFile;
}

// Deletes the temp file
void deleteTemporaryFile(void) {
	if(tempFile == NULL) return;
	unlink(tempFileString);
	fclose(tempFile);
	
	tempFile = NULL;
	tempFileString = NULL;
	[tempFileStringNS release];
	tempFileStringNS = nil;
}

// Gets an NSString representing the temp file
NSString* getTemporaryFilePath(void) {
	return tempFileStringNS;
}

// Converts a Cocoa path to an HFS (colon-delimited) path
NSString* posixPathToHFSPath(NSString *posixPath) {
	CFURLRef url = CFURLCreateWithFileSystemPath(NULL, (CFStringRef)
												 posixPath,
												 kCFURLPOSIXPathStyle, NO);
	return (NSString *)CFURLCopyFileSystemPath(url, kCFURLHFSPathStyle);
}

// Copies an NSString to a malloc'd char*
char* copyNSString(NSString* string) {
	const char* utf8String = [string UTF8String];
	
	// Copy the string so it doesn't get autoreleased
	size_t stringSize = strlen(utf8String)+1;
	char* newString = (char*) malloc(stringSize);
	strlcpy(newString, utf8String, stringSize);
	return newString;
}

// Escapes a C string for use with AppleScript
void freeString(char* string) {
	free(string);
}

// Generates a random string.
NSString* generateRandomString(NSUInteger length) {
	char *alphabet = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXZY01234"
		"56789";
	size_t n = strlen(alphabet);
	char *randomString = (char*) malloc(sizeof(char) * (length+1));
	for(NSUInteger i=0; i<length; i++) {
		randomString[i] = alphabet[arc4random() % n]; 
	}
	randomString[length] = 0;
	
	return [NSString stringWithUTF8String:randomString];
}

// If (document_t*)x is a Word 2004 document, returns [y entryIndex]. Otherwise,
// returns [y entry_index].
NSInteger getEntryIndex(document_t* x, SBObject* y) {
	if(x->isWord2004) {
		return (NSInteger) [y performSelector:@selector(entryIndex)];
	} else {
		return (NSInteger) [y performSelector:@selector(entry_index)];
	}
}