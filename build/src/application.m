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

statusCode getDocument(bool isWord2004, const char* wordPath,
					   const char* documentName, document_t **returnValue) {
	clearError();
	document_t *doc = (document_t*) malloc(sizeof(document_t));
	doc->allocatedFieldsStart = NULL;
	doc->allocatedFieldsEnd = NULL;
	doc->allocatedFieldListsStart = NULL;
	doc->allocatedFieldListsEnd = NULL;
	
	// Get application by path
	NSString* wordPathNS = [NSString stringWithUTF8String:wordPath];
	WordApplication *wordApp = nil;
	
	// Get WordApplication object for path only if not already cached
	if(!wordApps || !(wordApp = [wordApps objectForKey:wordPathNS])) {
		NSURL* wordURL = [[NSURL alloc] initFileURLWithPath:wordPathNS];
		wordApp = [SBApplication applicationWithURL:wordURL];
		[wordURL release];
		if(!wordApps) {
			wordApps = [[NSMutableDictionary alloc] initWithCapacity:3];
		}
		[wordApps setObject:wordApp forKey:wordPathNS];
	}
	
	doc->sbApp = wordApp;
	[doc->sbApp retain];
	
	ZoteroSBApplicationDelegate *myDelegate = [[ZoteroSBApplicationDelegate
												alloc] init];
	[doc->sbApp setDelegate:myDelegate];
	
	// Set other parameters
	doc->isWord2004 = isWord2004;
	doc->sbDoc = [doc->sbApp activeDocument];
	CHECK_STATUS
	
	[doc->sbDoc retain];
	
	WordWindow* sbWindow = [doc->sbDoc activeWindow];
	CHECK_STATUS
	doc->sbView = [sbWindow view];
	CHECK_STATUS
	[doc->sbView retain];
	doc->sbProperties = [doc->sbDoc customDocumentProperties];
	CHECK_STATUS
	[doc->sbProperties retain];
	
	// Show appropriate error if there is no document to initialize, or if VBA
	// is not installed
	if(!doc->sbDoc || [doc->sbDoc isEqual:[NSNull null]]) {
		freeDocument(doc);
		if(!sbWindow || [sbWindow isEqual:[NSNull null]]) {
			displayAlert("Zotero could not find an open document. Please open \
						 a document and try again.", DIALOG_ICON_STOP, 
						 DIALOG_BUTTONS_OK, NULL);
		} else {
			displayAlert("Zotero could not perform this action. Please ensure "
						 "that a document is open. If you have performed a "
						 "custom installation of Office, you may need to run "
						 "the installer again, ensuring that \"Visual Basic "
						 "for Applications\" is selected.", DIALOG_ICON_STOP, 
						 DIALOG_BUTTONS_OK, NULL);
		}
		// TODO raise ExceptionAlreadyDisplayed
	}
	
	// Capture settings
	IGNORING_SB_ERRORS_BEGIN
	doc->restoreFullScreenMode = [doc->sbView fullScreen];
	doc->restoreInsertionsAndDeletions = [doc->sbView
										  showInsertionsAndDeletions];
	doc->restoreFormatChanges = [doc->sbView showFormatChanges];
	IGNORING_SB_ERRORS_END
	
	// Get out of full screen mode
	if(doc->restoreFullScreenMode) {
		[doc->sbView setFullScreen:NO];
		CHECK_STATUS
	}
	
	// Set statuses
	doc->statusFullScreenMode = false;
	doc->statusInsertionsAndDeletions = doc->restoreInsertionsAndDeletions;
	doc->statusFormatChanges = doc->restoreFormatChanges;
	
	// Put path into structure
	size_t pathSize = strlen(wordPath) + 1;
	doc->wordPath = (char*) malloc(sizeof(char)*pathSize);
	memcpy(doc->wordPath, wordPath, pathSize);
	
	// Initialize lock
	doc->lock = [[NSRecursiveLock alloc] init];
	
	*returnValue = doc;
	return STATUS_OK;
}

// Displays an alert within Word
statusCode displayAlert(char const dialogText[], unsigned short icon,
						unsigned short buttons, unsigned short *returnValue) {
	// Escape quotation marks and backslashes
	NSMutableString* dialogTextNS = escapeString(dialogText);
	
	// Get array of buttons
	NSArray *buttonArray;
	if(buttons == DIALOG_BUTTONS_OK_CANCEL) {
		buttonArray = [NSArray arrayWithObjects:@"Cancel", @"OK", nil];
	} else if(buttons == DIALOG_BUTTONS_YES_NO) {
		buttonArray = [NSArray arrayWithObjects:@"No", @"Yes", nil];
	} else if(buttons == DIALOG_BUTTONS_YES_NO_CANCEL) {
		buttonArray = [NSArray arrayWithObjects:@"Cancel", @"No", @"Yes",
					   nil];
	} else {
		buttonArray = [NSArray arrayWithObjects:@"OK", nil];
	}
	NSString *buttonCode = [buttonArray componentsJoinedByString: @"\", \""];
	
	// Get icon
	NSString *dialogType;
	if(icon == DIALOG_ICON_CAUTION) {
		dialogType = @"warning";
	} else if(icon == DIALOG_ICON_NOTICE) {
		dialogType = @"informational";
	} else {
		dialogType = @"critical";
	}
	
	// Generate dialog code
	NSString* dialogCode = [[NSString alloc] initWithFormat:
							@"tell application \"Microsoft Word\" "
							"to display alert \"%@\" as %@ buttons {\"%@\"}",
							dialogTextNS, dialogType, buttonCode];
	
	// Display dialog
	NSDictionary *errorDict = nil;
	NSAppleScript* scriptObject = [[NSAppleScript alloc]
								   initWithSource:dialogCode];
	NSAppleEventDescriptor *resultDescriptor = [scriptObject
												executeAndReturnError:
												&errorDict];
	[scriptObject release];
	
	// Check button clicked
	if(returnValue != NULL) {
		NSAppleEventDescriptor *buttonDesc = [resultDescriptor
											  descriptorForKeyword:'bhit'];
		NSString *buttonPressed = [buttonDesc stringValue];
		*returnValue = [buttonArray indexOfObject:buttonPressed];
	}
	
	return STATUS_OK;
}