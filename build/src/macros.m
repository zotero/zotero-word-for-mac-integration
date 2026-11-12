/*
	***** BEGIN LICENSE BLOCK *****

	Copyright (c) 2026  Corporation for Digital Scholarship
						Vienna, Virginia, USA
						https://www.zotero.org

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

// Runs macros of the Zotero template for resource/macroDocument.mjs, which implements documents
// and fields with them. This file only passes the input and output of a macro; the protocol is
// described in ZoteroMacros.bas.
//
// Input and output go through document variables, since macros run with Apple Events can't take
// arguments or return values. Word handles one Apple Event per display frame (17 ms on a 60 Hz
// display), so the events that write the input and run the macro are sent without waiting for
// replies, and Word handles them together with the event that reads the output.

#include "zoteroMacWordIntegration.h"

#define MACRO_INPUT_PREFIX @"ZOTERO_MACRO_IN_"
#define MACRO_OUTPUT_PREFIX @"ZOTERO_MACRO_OUT_"
// Document variables hold at most 65,280 characters
#define MACRO_CHUNK_LENGTH 60000

// Splits data into variable values, the first of which starts with the number of values
static NSArray* splitIntoChunks(NSString* data) {
	NSMutableArray* chunks = [NSMutableArray array];
	NSUInteger length = [data length];
	NSUInteger position = 0;
	do {
		NSUInteger chunkLength = MIN(MACRO_CHUNK_LENGTH - 16, length - position);
		// Don't split a surrogate pair across variables
		if(position + chunkLength < length
				&& CFStringIsSurrogateHighCharacter([data characterAtIndex:
													 position + chunkLength - 1])) {
			chunkLength--;
		}
		[chunks addObject:[data substringWithRange:NSMakeRange(position, chunkLength)]];
		position += chunkLength;
	} while(position < length);
	chunks[0] = [NSString stringWithFormat:@"%lu;%@", (unsigned long) [chunks count], chunks[0]];
	return chunks;
}

// Sends "make new variable at <document> with properties {name:..., variable value:...}".
// Word doesn't support Scripting Bridge's -[SBElementArray addObject:].
static void sendMakeVariable(document_t* doc, NSString* name, NSString* value) {
	NSAppleEventDescriptor* properties = [NSAppleEventDescriptor recordDescriptor];
	[properties setDescriptor:[NSAppleEventDescriptor descriptorWithString:name]
				   forKeyword:'pnam'];
	[properties setDescriptor:[NSAppleEventDescriptor descriptorWithString:value]
				   forKeyword:'1550'];
	[doc->sbApp sendEvent:'core' id:'crel' parameters:'kocl',
	 [NSAppleEventDescriptor descriptorWithTypeCode:'w112'], 'insh', doc->sbDoc,
	 'prdt', properties, nil];
}

// Deletes variables without waiting for replies
static void deleteVariables(document_t* doc, NSString* prefix, NSUInteger count) {
	AESendMode sendMode = [doc->sbApp sendMode];
	[doc->sbApp setSendMode:kAENoReply];
	IGNORING_SB_ERRORS_BEGIN
	@try {
		for(NSUInteger i=1; i<=count; i++) {
			[[[doc->sbDoc variables] objectWithName:
			  [NSString stringWithFormat:@"%@%lu", prefix, (unsigned long) i]] delete];
		}
	}
	@finally {
		IGNORING_SB_ERRORS_END
		[doc->sbApp setSendMode:sendMode];
	}
}

// Reads the output of a macro and sets *chunkCount to the number of variables it was in. Returns
// nil, with the error flagged, if there is no output.
static NSString* readOutput(document_t* doc, NSUInteger* chunkCount) {
	*chunkCount = 0;
	SBElementArray* sbVariables = [doc->sbDoc variables];
	NSString* first = [[sbVariables objectWithName:
						[MACRO_OUTPUT_PREFIX stringByAppendingString:@"1"]] variableValue];
	if(errorHasOccurred() || !first) return nil;
	NSRange separator = [first rangeOfString:@";" options:NSLiteralSearch];
	if(separator.location == NSNotFound) return nil;
	NSInteger count = [[first substringToIndex:separator.location] integerValue];
	if(count < 1) return nil;
	NSMutableString* output = [NSMutableString stringWithString:
							   [first substringFromIndex:separator.location + 1]];
	for(NSInteger i=2; i<=count; i++) {
		NSString* chunk = [[sbVariables objectWithName:
							[NSString stringWithFormat:@"%@%ld", MACRO_OUTPUT_PREFIX, (long) i]]
						   variableValue];
		if(errorHasOccurred() || !chunk) return nil;
		[output appendString:chunk];
	}
	*chunkCount = count;
	return output;
}

static statusCode callMacroLocked(document_t *doc, const char macroName[], const char input[],
								  const char clipboardData[], unsigned short clipboardFormat,
								  bool probe, char **returnValue);

// Runs a macro of the Zotero template with the given input and returns its output. If
// clipboardFormat is CLIPBOARD_RTF or CLIPBOARD_HTML, clipboardData is put on the clipboard while
// the macro runs, and the clipboard is restored afterwards.
//
// If probe is true, waits for Word to start the macro, and sets *returnValue to NULL instead of
// failing if the macro doesn't exist or doesn't respond. This tells whether the template has the
// macros.
statusCode callMacro(document_t *doc, const char macroName[], const char input[],
					 const char clipboardData[], unsigned short clipboardFormat, bool probe,
					 char **returnValue) {
	HANDLE_EXCEPTIONS_BEGIN
	[doc->lock lock];
	@try {
		return callMacroLocked(doc, macroName, input, clipboardData, clipboardFormat, probe,
							   returnValue);
	}
	@finally {
		[doc->lock unlock];
	}
	HANDLE_EXCEPTIONS_END
}

static statusCode callMacroLocked(document_t *doc, const char macroName[], const char input[],
								  const char clipboardData[], unsigned short clipboardFormat,
								  bool probe, char **returnValue) {
	clearError();
	*returnValue = NULL;

	NSString* name = [NSString stringWithUTF8String:macroName];
	NSArray* chunks = splitIntoChunks([NSString stringWithUTF8String:input]);
	NSString* output = nil;
	NSUInteger outputChunkCount = 0;

	AESendMode sendMode = [doc->sbApp sendMode];
	BOOL clipboardSaved = NO;
	@try {
		if(clipboardFormat != CLIPBOARD_NONE) {
			storePasteboardItems();
			clipboardSaved = YES;
			if(clipboardFormat == CLIPBOARD_RTF) {
				replacePasteboardContentsWithRTF(clipboardData);
			} else {
				replacePasteboardContentsWithHTML(clipboardData);
			}
		}

		[doc->sbApp setSendMode:kAENoReply];
		for(NSUInteger i=0; i<[chunks count]; i++) {
			sendMakeVariable(doc, [NSString stringWithFormat:@"%@%lu", MACRO_INPUT_PREFIX,
								   (unsigned long) i+1], chunks[i]);
		}
		if(probe) [doc->sbApp setSendMode:sendMode];
		[doc->sbApp runVBMacroMacroName:name];
		[doc->sbApp setSendMode:sendMode];
		if(!errorHasOccurred()) {
			// Waits until the macro has finished
			output = readOutput(doc, &outputChunkCount);
		}
	}
	@finally {
		[doc->sbApp setSendMode:sendMode];
		// The output has been read, so the macro is done with the clipboard
		if(clipboardSaved) restorePasteboardContents();
	}

	if(!output) {
		// The input is left behind if the macro didn't run
		deleteVariables(doc, MACRO_INPUT_PREFIX, [chunks count]);
		if(probe) {
			clearError();
			return STATUS_OK;
		}
		CHECK_STATUS
		DIE(@"The macro wrote no output")
	}

	// Don't leave the output in the document. Word handles this with the events of the next call.
	deleteVariables(doc, MACRO_OUTPUT_PREFIX, outputChunkCount);

	*returnValue = copyNSString(output);
	return STATUS_OK;
}
