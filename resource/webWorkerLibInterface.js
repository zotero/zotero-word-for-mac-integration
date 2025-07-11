/*
	***** BEGIN LICENSE BLOCK *****
	
	Copyright © 2023 Corporation for Digital Scholarship
					 Vienna, Virginia, USA
					 https://digitalscholar.org
	
	This file is part of Zotero.
	
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

// A web worker didn't used to be necessary until fx102 where we seem to be unable to pass
// a callback to ctypes and call it asynchronously. The callback can be called immediately
// but an attempt to call it asynchronously results in a segfault/memory access error,
// pointing to the JS engine unallocating, moving or restricting access to the callback function

// We need an async call to fetch fields, which takes a while and we cannot afford to block the main
// thread during the call

/**
 * Loads libZoteroMacWordIntegration.dylib and initializes js-ctypes functions
 */

import { MessagingGeneric } from "./messagingGeneric.mjs";

var field_t, document_t, fieldListNode_t, progressFunction_t, lib, fn, fieldPtr;
var docs = {};
var Zotero = {};

const STATUS_EXCEPTION = 1;
const STATUS_EXCEPTION_ALREADY_DISPLAYED = 2;
const STATUS_EXCEPTION_SB_DENIED = 3;

var Zotero = {};

var Messaging = new MessagingGeneric({
	sendMessage: (...args) => postMessage(args),
	addMessageListener: (listener) => addEventListener('message', (event) => {
		listener(event.data);
	}),
	functionOverrides: {
		'debug': true,
		'getString': true,
		'launchURL': true,
	},
	overrideTarget: Zotero
});
// Add listener for init
Messaging.addMessageListener('init', init);

function init(libPath) {
	if(lib) return;
	
	lib = ctypes.open(libPath);
	
	document_t = new ctypes.StructType("Document");
	
	field_t = ctypes.unsigned_long;
	
	fieldListNode_t = new ctypes.StructType("fieldListNode_t");
	fieldListNode_t.define([
		{ field: field_t },
		{ next: fieldListNode_t.ptr }
	]);

	fieldPtr = new ctypes.PointerType(field_t);
	
	progressFunction_t = new ctypes.FunctionType(ctypes.default_abi, ctypes.void_t,
		[ctypes.int]).ptr;
	
	var statusCode = ctypes.unsigned_short;
	fn = {
		// char* getError(void);
		getError: lib.declare("getError", ctypes.default_abi, ctypes.char.ptr),
		
		// void clearError(void);
		clearError: lib.declare("clearError", ctypes.default_abi, ctypes.void_t),

		// void preventAppNap(void);
		preventAppNap: lib.declare("preventAppNap", ctypes.default_abi, ctypes.void_t),

		// void allowAppNap(void);
		allowAppNap: lib.declare("allowAppNap", ctypes.default_abi, ctypes.void_t),
		
		// statusCode getDocument(int wordVersion, const char* wordPath,
		//					   const char* documentName, Document** returnValue);
		getDocument: lib.declare("getDocument", ctypes.default_abi, statusCode, ctypes.int,
			ctypes.char.ptr, ctypes.char.ptr, document_t.ptr.ptr),
		
		// statusCode activate(Document *doc);
		activate: lib.declare("activate", ctypes.default_abi, statusCode, document_t.ptr),
		
		// statusCode displayAlert(char const dialogText[], unsigned short icon,
		//						   unsigned short buttons, unsigned short* returnValue);
		displayAlert: lib.declare("displayAlert", ctypes.default_abi, ctypes.unsigned_short,
			document_t.ptr, ctypes.char.ptr, ctypes.unsigned_short, ctypes.unsigned_short,
			ctypes.unsigned_short.ptr),
		
		// statusCode canInsertField(Document *doc, const char fieldType[], bool* returnValue);
		canInsertField: lib.declare("canInsertField", ctypes.default_abi, statusCode,
			document_t.ptr, ctypes.char.ptr, ctypes.bool.ptr),
		
		// statusCode cursorInField(Document *doc, const char fieldType[], Field** returnValue);
		cursorInField: lib.declare("cursorInField", ctypes.default_abi, statusCode, document_t.ptr,
			ctypes.char.ptr, field_t.ptr),
		
		// statusCode getDocumentData(Document *doc, char **returnValue);
		getDocumentData: lib.declare("getDocumentData", ctypes.default_abi, statusCode,
			document_t.ptr, ctypes.char.ptr.ptr),
		
		// statusCode setDocumentData(Document *doc, const char documentData[]);
		setDocumentData: lib.declare("setDocumentData", ctypes.default_abi, statusCode,
			document_t.ptr, ctypes.char.ptr),
		
		// statusCode insertField(Document *doc, const char fieldType[],
		//					      unsigned short noteType, Field **returnValue)
		insertField: lib.declare("insertField", ctypes.default_abi, statusCode, document_t.ptr,
			ctypes.char.ptr, ctypes.unsigned_short, field_t.ptr),
		
		// statusCode getFields(document_t *doc, const char fieldType[],
		//					    fieldListNode_t** returnNode);
		getFields: lib.declare("getFields", ctypes.default_abi, statusCode, document_t.ptr,
			ctypes.char.ptr, fieldListNode_t.ptr.ptr),
		
		// statusCode setBibliographyStyle(Document *doc, long firstLineIndent, 
		//								   long bodyIndent, unsigned long lineSpacing,
		//								   unsigned long entrySpacing, long tabStops[],
		//								   unsigned long tabStopCount);
		setBibliographyStyle: lib.declare("setBibliographyStyle", ctypes.default_abi,
			statusCode, document_t.ptr, ctypes.long, ctypes.long, ctypes.unsigned_long,
			ctypes.unsigned_long, ctypes.long.array(), ctypes.unsigned_long),
		
		// statusCode exportDocument(Document *doc, const char fieldType[],
		// 							const char importInstructions[]);
		exportDocument: lib.declare("exportDocument", ctypes.default_abi, statusCode, document_t.ptr,
			ctypes.char.ptr, ctypes.char.ptr),

		// statusCode importDocument(Document *doc, const char fieldType[],
		// 							bool *returnValue);
		importDocument: lib.declare("importDocument", ctypes.default_abi, statusCode, document_t.ptr,
			ctypes.char.ptr, ctypes.bool.ptr),
		
		// statusCode insertText(Document *doc, const char htmlString[]);
		insertText: lib.declare("insertText", ctypes.default_abi, statusCode, document_t.ptr,
			ctypes.char.ptr),

		// statusCode convertPlaceholdersToFields(Document *doc, char* placeholders[],
		//		unsigned long nPlaceholders, unsigned short noteType, char fieldType[], listNode_t** returnNode);
		convertPlaceholdersToFields: lib.declare("convertPlaceholdersToFields", ctypes.default_abi, statusCode, document_t.ptr,
			ctypes.char.ptr.ptr, ctypes.unsigned_long, ctypes.unsigned_short,
			ctypes.char.ptr, fieldListNode_t.ptr.ptr),
		
		// statusCode convert(Document *doc, field_t* fields[], unsigned long nFields,
		//				      const char toFieldType[], unsigned short noteType[]);
		convert: lib.declare("convert", ctypes.default_abi, statusCode, document_t.ptr,
			field_t.ptr, ctypes.unsigned_long, ctypes.char.ptr, ctypes.unsigned_short.ptr),
		
		// statusCode cleanup(Document *doc);
		cleanup: lib.declare("cleanup", ctypes.default_abi, statusCode, document_t.ptr),
		
		// statusCode cleanup(Document *doc);
		complete: lib.declare("complete", ctypes.default_abi, statusCode, document_t.ptr),

		// statusCode deleteField(XPCField field);
		deleteField: lib.declare("deleteField", ctypes.default_abi, statusCode, field_t),

		// statusCode removeCode(XPCField field);
		removeCode: lib.declare("removeCode", ctypes.default_abi, statusCode, field_t),
			
		// statusCode selectField(XPCField field);
		selectField: lib.declare("selectField", ctypes.default_abi, statusCode, field_t),
			
		// statusCode setText(XPCField field, const char string[], bool isRich);
		setText: lib.declare("setText", ctypes.default_abi, statusCode, field_t,
			ctypes.char.ptr, ctypes.bool),
			
		// statusCode getText(XPCField field, char** returnValue);
		getText: lib.declare("getText", ctypes.default_abi, statusCode, field_t,
			ctypes.char.ptr.ptr),
			
		// statusCode setCode(XPCField field, const char code[]);
		setCode: lib.declare("setCode", ctypes.default_abi, statusCode, field_t,
			ctypes.char.ptr),
		
		// statusCode getCode(XPCField field, char** returnValue);
		getCode: lib.declare("getCode", ctypes.default_abi, statusCode, field_t,
			ctypes.char.ptr.ptr),
		
		// statusCode getNoteIndex(XPCField field, unsigned long *returnValue);
		getNoteIndex: lib.declare("getNoteIndex", ctypes.default_abi, statusCode,
			field_t, ctypes.unsigned_long.ptr),

		// statusCode isAdjacentToNextField(XPCField field, unsigned long *returnValue);
		isAdjacentToNextField: lib.declare("isAdjacentToNextField", ctypes.default_abi, statusCode,
			field_t, ctypes.bool.ptr),
		
		// statusCode equals(XPCField a, XPCField b, bool *returnValue);
		equals: lib.declare("equals", ctypes.default_abi, statusCode,
			field_t, field_t, ctypes.bool.ptr),
		
		// statusCode install(const char zoteroDotPath[], const char zoteroDotmPath[],
		// 					  const char zoteroScptPath[]);
		install: lib.declare("install", ctypes.default_abi, statusCode, ctypes.char.ptr,
			ctypes.char.ptr, ctypes.char.ptr),

		// statusCode canInsertField(bool* returnValue);
		isWordInstalled: lib.declare("isWordInstalled", ctypes.default_abi, statusCode, ctypes.bool.ptr),
		
		// statusCode freeData(void* ptr);
		freeData: lib.declare("freeData", ctypes.default_abi, statusCode, ctypes.void_t.ptr),

		// bool isWordArm();
		isWordArm: lib.declare("isWordArm", ctypes.default_abi, ctypes.bool),
		
		// char *getMacOSVersion();
		getMacOSVersion: lib.declare("getMacOSVersion", ctypes.default_abi, ctypes.char.ptr),
		
		// int isZoteroRosetta();
		isZoteroRosetta: lib.declare("isZoteroRosetta", ctypes.default_abi, ctypes.int),
	};
	// Application handlers
	Messaging.addMessageListener('getDocument', getDocument)
	Messaging.addMessageListener('getActiveDocument', getDocument)
	// Document handlers
	for (let func of ['activate', 'displayAlert', 'canInsertField',
			'getDocumentData', 'setDocumentData', 'getFields',
			'setBibliographyStyle', 'exportDocument', 'importDocument', 'insertText',
			'convertPlaceholdersToFields', 'cleanup', 'complete']) {
		Messaging.addMessageListener(func, (...args) => {
			let doc = docs[args[0]];
			if (!doc) {
				throw new Error(`Document ${args[0]} not found`);
			}
			return doc[func](...args.slice(1));
		});
	}
	
	// Document handlers where a field may be returned
	for (let func of ['cursorInField', 'insertField']) {
		Messaging.addMessageListener(func, async (...args) => {
			let doc = docs[args[0]];
			if (!doc) {
				throw new Error(`Document ${args[0]} not found`);
			}
			let field = await doc[func](...args.slice(1));
			return field ? field.id : null;
		});
	}
	// Document handlers where field array is returned
	for (let func of ['getFields', 'convertPlaceholdersToFields']) {
		Messaging.addMessageListener(func, async (...args) => {
			let doc = docs[args[0]];
			if (!doc) {
				throw new Error(`Document ${args[0]} not found`);
			}
			let fields = await doc[func](...args.slice(1));
			return fields.map(field => field.id);
		});
	}
	// Document handler where a field array is the first argument
	Messaging.addMessageListener('convert', async (...args) => {
		let doc = docs[args[0]];
		if (!doc) {
			throw new Error(`Document ${args[0]} not found`);
		}
		let fields = args[1].map(id => doc.fields[id]);
		return doc.convert(fields, ...args.slice(2));
	});
	// Field handlers
	for (let func of ['delete', 'removeCode', 'select', 'setText', 'getText', 'setCode', 'getCode',
			'getNoteIndex', 'isAdjacentToNextField']) {
		Messaging.addMessageListener(func, (...args) => {
			let doc = docs[args[0]];
			if (!doc) {
				throw new Error(`Document ${args[0]} not found`);
			}
			let field = doc.fields[args[1]];
			if (!field) {
				throw new Error(`Field ${args[1]} not found`);
			}
			return field[func](...args.slice(2));
		});
	}
	// Field handler where 1st param is a field
	Messaging.addMessageListener('equals', (...args) => {
		let doc = docs[args[0]];
		if (!doc) {
			throw new Error(`Document ${args[0]} not found`);
		}
		let field = doc.fields[args[1]];
		if (!field) {
			throw new Error(`Field ${args[1]} not found`);
		}
		let otherField = doc.fields[args[2]];
		if (!otherField) {
			throw new Error(`Field ${args[2]} not found`);
		}
		return field.equals(otherField, ...args.slice(3));
	});
	// Static function overrides
	for (let func of ['preventAppNap', 'allowAppNap', 'install', 'isWordArm', 'isZoteroRosetta']) {
		Messaging.addMessageListener(func, (...args) => {
			return fn[func](...args);
		});
	}
	Messaging.addMessageListener('isWordInstalled', () => {
		var returnValue = new ctypes.bool();
		checkStatus(fn.isWordInstalled(returnValue.address()));
		return returnValue.value;
	});
	Messaging.addMessageListener('getMacOSVersion', (...args) => {
		return fn.getMacOSVersion(...args).readString();
	});
	Messaging.addMessageListener('getLastError', () => {
		return getLastError();
	});
	
}

function randomString(len, chars) {
	if (!chars) {
		chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
	}
	if (!len) {
		len = 8;
	}
	var randomstring = '';
	for (var i=0; i<len; i++) {
		var rnum = Math.floor(Math.random() * chars.length);
		randomstring += chars.substring(rnum,rnum+1);
	}
	return randomstring;
}

/**
 * Gets the last error that took place in C code.
 */
function getLastError() {
	var errPtr = fn.getError();
	if(errPtr.isNull()) {
		var err = "An unexpected error occurred.";
	} else {
		var err = errPtr.readString().replace("\u2019", "'", "g");
	}
	fn.clearError();
	return err;
}

/**
 * Checks the return status of a function to verify that no error occurred.
 * @param {Integer} status The return status code of a C function
 */
function checkStatus(status) {
	if(!status) return;

	if (status === STATUS_EXCEPTION) {
		throw new Error(getLastError());
	} else {
		if (status === STATUS_EXCEPTION_SB_DENIED) {
			(async () => {
				let message = await Zotero.getString('integration.error.macWordSBPermissionsMissing');
				let index = await Messaging.sendMessage('displayMoreInformationAlert', [
					await Zotero.getString('integration.error.macWordSBPermissionsMissing.title'),
					message
				]);
				if (index == 1) {
					Zotero.launchURL('https://www.zotero.org/support/kb/mac_word_permissions_missing')
				}
			})();
		}
		throw new Error("ExceptionAlreadyDisplayed");
	}
}

/**
 * Ensures that the document associated with this object has not been garbage collected
 */
function checkIfFreed(documentStatus) {
	if(!documentStatus.active) throw new Error("complete() method already called on document");
}


function getDocument(version, path) {
	var docPtr = new document_t.ptr();
	checkStatus(fn.getDocument(version, path, null, docPtr.address()));
	let doc = new Document(docPtr);
	return doc.id;
}

/**
 * See integrationTest.js
 */
var Document = function(cDoc) {
	this._document_t = cDoc;
	this._documentStatus = {active: true};
	this.id = randomString();
	this.fields = {};
	docs[this.id] = this;
};
Document.prototype = {
	displayAlert: function(dialogText, icon, buttons) {
		Zotero.debug("ZoteroMacWordIntegration: displayAlert", 4);
		var buttonPressed = new ctypes.unsigned_short();
		checkStatus(fn.displayAlert(this._document_t, dialogText, icon, buttons,
			buttonPressed.address()));
		return buttonPressed.value;
	},
	
	activate: function() {
		Zotero.debug("ZoteroMacWordIntegration: activate", 4);
		checkIfFreed(this._documentStatus);
		checkStatus(fn.activate(this._document_t));
	},
	
	canInsertField: function(fieldType) {
		Zotero.debug("ZoteroMacWordIntegration: canInsertField", 4);
		checkIfFreed(this._documentStatus);
		var returnValue = new ctypes.bool();
		checkStatus(fn.canInsertField(this._document_t, fieldType, returnValue.address()));
		return returnValue.value;
	},
	
	cursorInField: function(fieldType) {
		Zotero.debug("ZoteroMacWordIntegration: cursorInField", 4);
		checkIfFreed(this._documentStatus);
		var returnValue = new field_t();
		checkStatus(fn.cursorInField(this._document_t, fieldType, returnValue.address()));
		return (returnValue.value == 0 ? null : new Field(returnValue, this));
	},
	
	getDocumentData: function() {
		Zotero.debug("ZoteroMacWordIntegration: getDocumentData", 4);
		checkIfFreed(this._documentStatus);
		var returnValue = new ctypes.char.ptr();
		checkStatus(fn.getDocumentData(this._document_t, returnValue.address()));
		var data = returnValue.readString();
		fn.freeData(returnValue);
		return data;
	},
	
	setDocumentData: function(documentData) {
		Zotero.debug("ZoteroMacWordIntegration: setDocumentData", 4);
		checkIfFreed(this._documentStatus);
		checkStatus(fn.setDocumentData(this._document_t, documentData));
	},
	
	insertField: function(fieldType, noteType) {
		Zotero.debug("ZoteroMacWordIntegration: insertField", 4);
		checkIfFreed(this._documentStatus);
		var returnValue = new field_t();
		checkStatus(fn.insertField(this._document_t, fieldType, noteType, returnValue.address()));
		return new Field(returnValue, this);
	},
	
	getFields: function(fieldType, observer) {
		Zotero.debug("ZoteroMacWordIntegration: getFields", 4);
		checkIfFreed(this._documentStatus);
		var fieldListNode = new fieldListNode_t.ptr();
		checkStatus(fn.getFields(this._document_t, fieldType, fieldListNode.address()));
		var fnum = new FieldEnumerator(fieldListNode, this);
		var fields = [];
		while (fnum.hasMoreElements()) {
			fields.push(fnum.getNext());
		}
		return fields;
	},
	
	setBibliographyStyle: function(firstLineIndent, bodyIndent, lineSpacing, entrySpacing,
			tabStops) {
		Zotero.debug("ZoteroMacWordIntegration: setBibliographyStyle", 4);
		checkIfFreed(this._documentStatus);
		checkStatus(fn.setBibliographyStyle(this._document_t, firstLineIndent, bodyIndent, lineSpacing,
			entrySpacing, ctypes.long.array(tabStops.length)(tabStops), tabStops.length));
	},

	importDocument: function(fieldType) {
		Zotero.debug(`ZoteroWinMacIntegration: importDocument`, 4);
		checkIfFreed(this._documentStatus);
		var returnValue = new ctypes.bool();
		checkStatus(fn.importDocument(this._document_t, fieldType, returnValue.address()));
		return returnValue.value;
	},

	exportDocument: function(fieldType, importInstructions) {
		Zotero.debug(`ZoteroWinMacIntegration: exportDocument`, 4);
		checkIfFreed(this._documentStatus);
		checkStatus(fn.exportDocument(this._document_t, fieldType, importInstructions));
	},
	
	insertText: function(text) {
		Zotero.debug(`ZoteroMacWordIntegration: insertText`, 4);
		checkIfFreed(this._documentStatus);
		checkStatus(fn.insertText(this._document_t, text));
	},

	convertPlaceholdersToFields: function(placeholderIDs, noteType, fieldType) {
		Zotero.debug("ZoteroMacWordIntegration: convertPlaceholdersToFields", 4);
		checkIfFreed(this._documentStatus);
		var cPlaceholderIDs = placeholderIDs.map(placeholderID => ctypes.char.array()(placeholderID));
		var fieldListNode = new fieldListNode_t.ptr();
		checkStatus(
			fn.convertPlaceholdersToFields(
				this._document_t,
				ctypes.char.ptr.array()(cPlaceholderIDs),
				placeholderIDs.length,
				noteType,
				fieldType,
				fieldListNode.address()
			)
		);
		var fnum = new FieldEnumerator(fieldListNode, this);
		var fields = [];
		while (fnum.hasMoreElements()) {
			fields.push(fnum.getNext());
		}
		return fields;
	},
	
	convert: function(fields, toFieldType, toNoteTypes) {
		Zotero.debug("ZoteroMacWordIntegration: convert", 4);
		checkIfFreed(this._documentStatus);
		fields = fields.map(field => field._field_t);
		checkStatus(fn.convert(this._document_t, field_t.array()(fields),
			fields.length, ctypes.char.array()(toFieldType),
			ctypes.unsigned_short.array()(toNoteTypes)));
	},
	
	cleanup: function() {
		Zotero.debug("ZoteroMacWordIntegration: cleanup", 4);
		if (this._documentStatus.active) {
			checkStatus(fn.cleanup(this._document_t));
		}
		else {
			Zotero.debug("complete() already called on document; ignoring", 4);
		}
	},
	
	complete: function() {
		Zotero.debug("ZoteroMacWordIntegration: complete", 4);
		if (this._documentStatus.active) {
			checkStatus(fn.complete(this._document_t));
			this._documentStatus.active = false;
			fn.allowAppNap();
			// Remove this document from docs cache
			delete docs[this.id];
		}
		else {
			Zotero.debug("complete() already called on document; ignoring", 4);
		}
	}
};

/**
 * An enumerator implementation to handle passing off fields
 */
var FieldEnumerator = function(startNode, doc) {
	this._currentNode = startNode;
	this._doc = doc;
	this._documentStatus = doc._documentStatus;
};
FieldEnumerator.prototype = {
	hasMoreElements: function() {
		checkIfFreed(this._documentStatus);
		return !this._currentNode.isNull();
	},
	
	getNext: function() {
		checkIfFreed(this._documentStatus);
		var contents = this._currentNode.contents;
		var fieldPtr = contents.addressOfField("field").contents;
		this._currentNode = contents.addressOfField("next").contents;
		return new Field(fieldPtr, this._doc);
	},
};

/**
 * See integrationTest.js
 */
var Field = function(field_t, doc) {
	this._field_t = field_t;
	this._documentStatus = doc._documentStatus;
	this.id = randomString();
	doc.fields[this.id] = this;
};
Field.prototype = {
	delete: function() {
		Zotero.debug("ZoteroMacWordIntegration: delete", 4);
		checkIfFreed(this._documentStatus);
		checkStatus(fn.deleteField(this._field_t));
	},
	
	removeCode: function() {
		Zotero.debug("ZoteroMacWordIntegration: removeCode", 4);
		checkStatus(fn.removeCode(this._field_t));
	},
	
	select: function() {
		Zotero.debug("ZoteroMacWordIntegration: select", 4);
		checkIfFreed(this._documentStatus);
		checkStatus(fn.selectField(this._field_t));
	},
	
	setText: function(text, isRich) {
		Zotero.debug(`ZoteroMacWordIntegration: setText rtf:${isRich} ${text}`, 4);
		checkIfFreed(this._documentStatus);
		checkStatus(fn.setText(this._field_t, text, isRich));
	},
	
	getText: function() {
		checkIfFreed(this._documentStatus);
		var returnValue = new ctypes.char.ptr();
		checkStatus(fn.getText(this._field_t, returnValue.address()));
		var val = returnValue.readString();
		Zotero.debug(`ZoteroMacWordIntegration: getText ${val}`, 4);
		return val;
	},
	
	setCode: function(code) {
		Zotero.debug(`ZoteroMacWordIntegration: setCode ${code}`, 4);
		checkIfFreed(this._documentStatus);
		checkStatus(fn.setCode(this._field_t, code));
	},
	
	getCode: function() {
		checkIfFreed(this._documentStatus);
		var returnValue = new ctypes.char.ptr();
		checkStatus(fn.getCode(this._field_t, returnValue.address()));
		var val = returnValue.readString();
		Zotero.debug(`ZoteroMacWordIntegration: getCode ${val}`, 4);
		return val;
	},
	
	equals: function(field) {
		Zotero.debug("ZoteroMacWordIntegration: equals", 4);
		checkIfFreed(this._documentStatus);
		var returnValue = new ctypes.bool();
		checkStatus(fn.equals(this._field_t, field._field_t, returnValue.address()));
		return returnValue.value;
	},
	
	getNoteIndex: function() {
		Zotero.debug("ZoteroMacWordIntegration: getNoteIndex", 4);
		checkIfFreed(this._documentStatus);
		var returnValue = new ctypes.unsigned_long();
		checkStatus(fn.getNoteIndex(this._field_t, returnValue.address()));
		return parseInt(returnValue.value);
	},
	
	isAdjacentToNextField: function() {
		Zotero.debug("ZoteroMacWordIntegration: isAdjacentToNextField", 4);
		checkIfFreed(this._documentStatus);
		var returnValue = new ctypes.bool();
		checkStatus(fn.isAdjacentToNextField(this._field_t, returnValue.address()));
		return returnValue.value;
	}
}
