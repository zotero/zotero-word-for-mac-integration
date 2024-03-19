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

// Cache for WordApplication objects, since initialization can take a while.
NSMutableDictionary* wordApps = nil;

statusCode getDocument(int wordVersion, const char* wordPath,
					   const char* documentName, document_t **returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	clearError();
	
	document_t *doc = (document_t*) malloc(sizeof(document_t));
	doc->allocatedFieldsStart = NULL;
	doc->allocatedFieldsEnd = NULL;
	doc->allocatedFieldListsStart = NULL;
	doc->allocatedFieldListsEnd = NULL;
	doc->restoreNote = doc->restoreCursorEnd = doc->restoreFieldIdx = -1;
	doc->cursorMoved = NO;
	doc->shouldRestoreCursor = YES;
	doc->insertTextIntoNote = 0;
	
	// Get application by path if a path has been specified
	NSString* wordPathNS;
	if(!wordPath) {
		NSURL* url;
		CHECK_OSSTATUS(LSGetApplicationForInfo(kLSUnknownType, 'MSWD', NULL,
												  kLSRolesAll, NULL,
												  (CFURLRef*) &url));
		wordPathNS = [url path];
	} else {
		wordPathNS = [NSString stringWithUTF8String:wordPath];
	}
	
	// Get WordApplication object for path only if not already cached
	WordApplication *wordApp = nil;
	if(!wordApps || !(wordApp = [wordApps objectForKey:wordPathNS])) {
		NSURL* wordURL = [[NSURL alloc] initFileURLWithPath:wordPathNS];
		wordApp = [SBApplication applicationWithURL:wordURL];
		[wordURL release];
		if(!wordApps) {
			wordApps = [[NSMutableDictionary alloc] initWithCapacity:3];
		}
		[wordApps setObject:wordApp forKey:wordPathNS];
	}
	
	// Tell Word to execute one of the scripting additions functions
	// so that it loads the necessary handling code for Scripting Additions stuff
	// and sending an event to display a dialog works
	// Scripting additions are loaded automatically when running MacScript
	// from Word, but it triggers a Run-Time Error 5 for some users.
	//
	// But not with ARM Word because that causes Zotero to freeze/crash (see commit message)
	if (!isWordArm()) {
		NSString* scriptNS = [[NSString alloc] initWithFormat:
							  @"tell application \"%@\" to time to GMT",
							  wordPathNS];
		NSAppleScript* scriptObject = [[NSAppleScript alloc]
									   initWithSource:scriptNS];
		NSDictionary *errorDict = nil;
		[scriptObject executeAndReturnError:&errorDict];
	}
	
	[wordApp setTimeout:kNoTimeOut];
	CHECK_STATUS
	
	doc->sbApp = wordApp;
	[doc->sbApp retain];
	
	// Initialize lock
	doc->lock = [[NSRecursiveLock alloc] init];
	
	ZoteroSBApplicationDelegate *myDelegate = [[ZoteroSBApplicationDelegate
												alloc] init];
	[doc->sbApp setDelegate:myDelegate];
	
	// Set other parameters
	doc->wordVersion = wordVersion;
	
	WordDocument* sbActiveDocument = [doc->sbApp activeDocument];
	CHECK_STATUS
	
	// See if we can change the reference from "active document" to one by name
	NSString* activeDocumentName = [sbActiveDocument name];
	CHECK_STATUS
	if(!activeDocumentName) {
		// Show appropriate error if there is no document to initialize, or if VBA
		// is not installed
		displayAlert(doc,
					 "Zotero could not perform this action.\n\n"
					 "In your Applications folder, select Microsoft Word, "
					 "go to File → Get Info, and make sure “Open in Rosetta” "
					 "is unchecked.\n\n"
					 "If you continue to have trouble, it may help to delete "
					 "Microsoft Word and reinstall it.", DIALOG_ICON_STOP,
					 DIALOG_BUTTONS_OK, NULL);
		[doc->sbApp release];
		free(doc);
		return STATUS_EXCEPTION_ALREADY_DISPLAYED;
	}
	CHECK_STATUS
	SBElementArray* sbDocuments = [doc->sbApp documents];
	CHECK_STATUS
	
	NSArray* documentNames = [sbDocuments valueForKey:@"name"];
	
	NSUInteger foundDocuments = 0;
	for(NSString* documentName in documentNames) {
		if([activeDocumentName isEqualTo:documentName]) {
			foundDocuments++;
		}
	}
	
	// It's possible that there are multiple documents by the same name. If
	// that's the case, we should just retain our reference to "active document"
	// Otherwise, we should use a reference by name
	if(foundDocuments == 1) {
		doc->sbDoc = [sbDocuments objectWithName:activeDocumentName];
	} else {
		doc->sbDoc = sbActiveDocument;
	}
	
	[doc->sbDoc retain];
	
	WordWindow* sbWindow = [doc->sbDoc activeWindow];
	CHECK_STATUS
	doc->sbView = [sbWindow view];
	CHECK_STATUS
	[doc->sbView retain];
	doc->sbProperties = [doc->sbDoc customDocumentProperties];
	CHECK_STATUS
	[doc->sbProperties retain];
	
	// Capture settings
	IGNORING_SB_ERRORS_BEGIN
	if([doc->sbView respondsToSelector:@selector(fullScreen)]) {
		doc->restoreFullScreenMode = [doc->sbView fullScreen];
	} else {
		doc->restoreFullScreenMode = false;
	}
	
	if([doc->sbView respondsToSelector:@selector(showInsertionsAndDeletions)]) {
		doc->restoreInsertionsAndDeletions = [doc->sbView
											  showInsertionsAndDeletions];
	} else {
		doc->restoreInsertionsAndDeletions = false;
	}
	
	if([doc->sbView respondsToSelector:@selector(showFormatChanges)]) {
		doc->restoreFormatChanges = [doc->sbView showFormatChanges];
	} else {
		doc->restoreFormatChanges = false;
	}
	
	if ([doc->sbDoc respondsToSelector:@selector(trackRevisions)]) {
		doc->restoreTrackRevisions = [doc->sbDoc trackRevisions];
		if (doc->restoreTrackRevisions) {
			[doc->sbDoc setTrackRevisions:NO];
		}
	} else {
		doc->restoreTrackRevisions = false;
	}
	
	if ([doc->sbView respondsToSelector:@selector(showFieldCodes)]) {
		[doc->sbView setShowFieldCodes:NO];
	}
	
	IGNORING_SB_ERRORS_END
	
	// Get out of full screen mode
	if(doc->restoreFullScreenMode) {
		[doc->sbView setFullScreen:NO];
		CHECK_STATUS
	}
	
	// Set statuses
	doc->statusInsertionsAndDeletions = doc->restoreInsertionsAndDeletions;
	doc->statusFormatChanges = doc->restoreFormatChanges;
	doc->statusTrackRevisions = doc->restoreTrackRevisions;
	
	// Put path into structure
	doc->wordPath = copyNSString(wordPathNS);
	
	storeCursorLocation(doc);
	CHECK_STATUS
	
	*returnValue = doc;
	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}
