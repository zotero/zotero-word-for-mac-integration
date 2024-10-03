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

FILE* tempFile = NULL;
char* tempFileString = NULL;
NSString* tempFileStringNS = nil;
NSArray *savedPasteboardItems = nil;
id <NSObject> _appNapPreventingActivity;
dispatch_queue_t _mainQueue;

void storePasteboardItems(void) {
	NSPasteboard *pasteboard = [NSPasteboard generalPasteboard];
	NSArray *items = [pasteboard pasteboardItems];
	
	NSMutableArray *copiedItems = [NSMutableArray arrayWithCapacity:[items count]];
	
	for (NSPasteboardItem *item in items) {
		NSPasteboardItem *itemCopy = [[NSPasteboardItem alloc] init];
		for (NSString *type in [item types]) {
			NSData *data = [item dataForType:type];
			[itemCopy setData:[data copy] forType:type];
		}
		[copiedItems addObject:itemCopy];
	}
	
	savedPasteboardItems = [copiedItems copy];
}

void restorePasteboardContents(void) {
	NSPasteboard *pasteboard = [NSPasteboard generalPasteboard];
	if (savedPasteboardItems) {
		[pasteboard clearContents];
		[pasteboard writeObjects:savedPasteboardItems];
	}
}

void replacePasteboardContentsWithRTF(const char *rtfString) {
	NSPasteboard *pasteboard = [NSPasteboard generalPasteboard];
	[pasteboard clearContents];
	
	
	NSString *nsRtfString = [NSString stringWithUTF8String:rtfString];
	NSData *rtfData = [nsRtfString dataUsingEncoding:NSUTF8StringEncoding];
	NSAttributedString *attributedString = [[NSAttributedString alloc] initWithRTF:rtfData documentAttributes:nil];
	
	[pasteboard writeObjects:@[attributedString]];
}

void replacePasteboardContentsWithHTML(const char *htmlString) {
	NSPasteboard *pasteboard = [NSPasteboard generalPasteboard];
	[pasteboard clearContents];
	
	NSString *nsHtmlString = [NSString stringWithUTF8String:htmlString];
	
	// Create an NSAttributedString from the HTML
	NSData *htmlData = [nsHtmlString dataUsingEncoding:NSUTF8StringEncoding];
	NSDictionary *options = @{NSDocumentTypeDocumentAttribute: NSHTMLTextDocumentType,
							  NSCharacterEncodingDocumentAttribute: @(NSUTF8StringEncoding)};
	NSAttributedString *attributedString = [[NSAttributedString alloc] initWithData:htmlData
																			options:options
																 documentAttributes:nil
																			  error:nil];
	
	// Write both HTML and RTF representations to the pasteboard
	[pasteboard writeObjects:@[nsHtmlString, attributedString]];
	
	// Explicitly set the types for better compatibility
	[pasteboard setString:attributedString.string forType:NSPasteboardTypeString];
	[pasteboard setPropertyList:nsHtmlString forType:NSPasteboardTypeHTML];
	[pasteboard setData:[attributedString RTFFromRange:NSMakeRange(0, attributedString.length) documentAttributes:@{}]
				forType:NSPasteboardTypeRTF];
}

void preventAppNap(void) {
	if (_appNapPreventingActivity != nil) return;
	NSActivityOptions options = NSActivityUserInitiatedAllowingIdleSystemSleep;
	NSString *reason = @"Zotero App Nap is disabled during a Word Integration operation";
	_appNapPreventingActivity = [[NSProcessInfo processInfo] beginActivityWithOptions:options reason:reason];
	[_appNapPreventingActivity retain];
}

void allowAppNap(void) {
	if (_appNapPreventingActivity == nil) return;
	[[NSProcessInfo processInfo] endActivity:_appNapPreventingActivity];
	[_appNapPreventingActivity release];
	_appNapPreventingActivity = nil;
}

void storeCursorLocation(document_t* doc) {
	IGNORING_SB_ERRORS_BEGIN
	if (doc->wordVersion >= 16 && doc->wordVersion < 2000) {
		doc->restoreCursorEnd = doc->restoreNote = -1;
		WordWdStoryType storyType = [[doc->sbApp selection] storyType];
		if (storyType == WordWdStoryTypeEndnotesStory || storyType == WordWdStoryTypeFootnotesStory) {
			doc->restoreNoteType = storyType;
			if (storyType == WordWdStoryTypeFootnotesStory) {
				doc->restoreNote = getEntryIndex(doc, [[[doc->sbApp selection] footnotes] objectAtIndex:0])-1;
			} else {
				doc->restoreNote = getEntryIndex(doc, [[[doc->sbApp selection] endnotes] objectAtIndex:0])-1;
			}
			doc->restoreCursorEnd = [[doc->sbApp selection] selectionEnd];
		} else {
			if ([[[doc->sbApp selection] fields] count]) {
				doc->restoreFieldIdx = getEntryIndex(doc, [[[doc->sbApp selection] fields] objectAtIndex:0]);
			} else {
				doc->restoreCursorEnd = [[doc->sbApp selection] selectionEnd];
			}
		}
	}
	IGNORING_SB_ERRORS_END
}

statusCode moveCursorOutOfNote(document_t* doc) {
	if (doc->wordVersion >= 16 && doc->wordVersion < 2000) {
		WordWdStoryType storyType = [[doc->sbApp selection] storyType];
		if (storyType == WordWdStoryTypeEndnotesStory || storyType == WordWdStoryTypeFootnotesStory) {
			WordTextRange* noteSelection;
			if (storyType == WordWdStoryTypeFootnotesStory) {
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
	IGNORING_SB_ERRORS_BEGIN
	if (!doc->cursorMoved || !doc->shouldRestoreCursor) {
		IGNORING_SB_ERRORS_END
		return STATUS_OK;
	}
	if (doc->restoreNote != -1) {
		if (doc->restoreNoteType == WordWdStoryTypeFootnotesStory) {
			[[[[doc->sbDoc footnotes] objectAtIndex:doc->restoreNote] textObject]
			 	sendEvent:'misc' id:'slct' parameters:'\00\00\00\00', nil];
		} else {
			[[[[doc->sbDoc endnotes] objectAtIndex:doc->restoreNote] textObject]
			 	sendEvent:'misc' id:'slct' parameters:'\00\00\00\00', nil];
		}
		[[doc->sbApp selection] setSelectionStart: doc->restoreCursorEnd];
		[[doc->sbApp selection] setSelectionEnd: [[doc->sbApp selection] selectionStart]];
	} else if (doc->restoreFieldIdx != -1) {
		long position = [[[[doc->sbDoc fields] objectAtIndex:doc->restoreFieldIdx-1] resultRange] endOfContent]+1;
		[[doc->sbDoc createRangeStart:position end:position]
		 sendEvent:'misc' id:'slct' parameters:'\00\00\00\00', nil];
	} else if (doc->restoreCursorEnd != -1) {
		[[doc->sbDoc createRangeStart:doc->restoreCursorEnd end:doc->restoreCursorEnd]
		 sendEvent:'misc' id:'slct' parameters:'\00\00\00\00', nil];
	}
	doc->cursorMoved = NO;
	IGNORING_SB_ERRORS_END
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
