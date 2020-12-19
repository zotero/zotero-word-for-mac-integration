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
void flagError(const char function[], NSString* file, unsigned int line) {
	if(lastErrorString) free(lastErrorString);
	lastErrorString = copyNSString([NSString stringWithFormat:@"%@ @[%s:%@:%d]",
									[lastError localizedDescription],
									function, file, line]);
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

// Manually throws an exception up to JS
void throwError(NSString *errorString, const char function[], NSString* file,
				unsigned int line) {
	if(lastErrorString) free(lastErrorString);
	lastErrorString = copyNSString([NSString stringWithFormat:@"%@ @[%s:%@:%d]",
									errorString, function, file, line]);
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

NSInteger getErrorCode(void) {
	return [lastError code];
}

FILE* tempFile = NULL;
char* tempFileString = NULL;
NSString* tempFileStringNS = nil;

// Gets a FILE for the temporary file, truncating it to zero length
FILE* getTemporaryFile(document_t *doc) {
	if(tempFile == NULL) {
        const char *tempFileTemplate;
        if(doc->wordVersion == 2016 || doc->wordVersion >= 16 && doc->wordVersion < 2000) {
            tempFileTemplate = [[NSTemporaryDirectory()
                                             stringByAppendingPathComponent:
                                             @"com.microsoft.Word/zotero.XXXXXX.rtf"]
                                            fileSystemRepresentation];
        } else {
            tempFileTemplate = [[NSTemporaryDirectory()
                                             stringByAppendingPathComponent:
                                             @"zotero.XXXXXX.rtf"]
                                            fileSystemRepresentation];
        }
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

// From https://developer.apple.com/forums/thread/653009 (https://archive.is/etgKB)
// Checks if Zotero is running as translated Rosetta 2 app (on an M1 Apple)
int isRosetta(void) {
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

// Checks if Word is running with the ARM architecture
bool isWordArm(void) {
	NSArray *runningApplications = [[NSWorkspace sharedWorkspace] runningApplications];
	for (NSRunningApplication* app in runningApplications) {
		NSRange range = [[app bundleIdentifier] rangeOfString:@"com.microsoft.Word"];
		if (range.location != NSNotFound) {
			return [app executableArchitecture] != NSBundleExecutableArchitectureX86_64;
		}
	}
	return false;
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

// Frees a C string
void freeData(void* ptr) {
	free(ptr);
}

void storeCursorLocation(document_t* doc) {
	if (doc->wordVersion >= 16 && doc->wordVersion < 2000) {
		doc->restoreCursorEnd = doc->restoreNote = -1;
		WordE160 storyType = [[doc->sbApp selection] storyType];
		if (storyType == WordE160EndnotesStory || storyType == WordE160FootnotesStory) {
			doc->restoreNoteType = storyType;
			if (storyType == WordE160FootnotesStory) {
				doc->restoreNote = getEntryIndex(doc, [[[doc->sbApp selection] footnotes] objectAtIndex:0])-1;
			} else {
				doc->restoreNote = getEntryIndex(doc, [[[doc->sbApp selection] endnotes] objectAtIndex:0])-1;
			}
		} else {
			if ([[[doc->sbApp selection] fields] count]) {
				doc->restoreFieldIdx = getEntryIndex(doc, [[[doc->sbApp selection] fields] objectAtIndex:0]);
			} else {
				doc->restoreCursorEnd = [[doc->sbApp selection] selectionEnd];
			}
		}
	}
}

statusCode moveCursorOutOfNote(document_t* doc) {
	if (doc->wordVersion >= 16 && doc->wordVersion < 2000) {
		WordE160 storyType = [[doc->sbApp selection] storyType];
		if (storyType == WordE160EndnotesStory || storyType == WordE160FootnotesStory) {
			WordTextRange* noteSelection;
			if (storyType == WordE160FootnotesStory) {
				noteSelection = [[[[doc->sbApp selection] footnotes] objectAtIndex:0] noteReference];
			} else {
				noteSelection = [[[[doc->sbApp selection] endnotes] objectAtIndex:0] noteReference];
			}
			// Absolutely clueless how Simon figured these out. Taken from selectField()
			[noteSelection sendEvent:'misc' id:'slct' parameters:'\00\00\00\00', nil];
			CHECK_STATUS;
			doc->cursorMoved = YES;
		}
	}
	return STATUS_OK;
}

statusCode restoreCursor(document_t* doc) {
	if (!doc->cursorMoved || !doc->shouldRestoreCursor) {
		return STATUS_OK;
	}
	if (doc->restoreNote != -1) {
		if (doc->restoreNoteType == WordE160FootnotesStory) {
			[[[[doc->sbDoc footnotes] objectAtIndex:doc->restoreNote] textObject]
			 	sendEvent:'misc' id:'slct' parameters:'\00\00\00\00', nil];
			CHECK_STATUS;
		} else {
			[[[[doc->sbDoc endnotes] objectAtIndex:doc->restoreNote] textObject]
			 	sendEvent:'misc' id:'slct' parameters:'\00\00\00\00', nil];
			CHECK_STATUS;
		}
	} else if (doc->restoreFieldIdx != -1) {
		long position = [[[[doc->sbDoc fields] objectAtIndex:doc->restoreFieldIdx-1] resultRange] endOfContent]+1;
		[[doc->sbDoc createRangeStart:position end:position]
		 sendEvent:'misc' id:'slct' parameters:'\00\00\00\00', nil];
		CHECK_STATUS;
	} else if (doc->restoreCursorEnd != -1) {
		[[doc->sbDoc createRangeStart:doc->restoreCursorEnd end:doc->restoreCursorEnd]
		 sendEvent:'misc' id:'slct' parameters:'\00\00\00\00', nil];
		CHECK_STATUS;
	}
	doc->cursorMoved = NO;
	return STATUS_OK;
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
	
	NSString *returnValue = [NSString stringWithUTF8String:randomString];
	free(randomString);
	return returnValue;
}

// If (document_t*)x is a Word 2004 document, returns [y entryIndex]. Otherwise,
// returns [y entry_index].
NSInteger getEntryIndex(document_t* x, SBObject* y) {
	if(x->wordVersion == 2004) {
		return (NSInteger) [y performSelector:@selector(entryIndex)];
	} else {
		return (NSInteger) [y performSelector:@selector(entry_index)];
	}
}
