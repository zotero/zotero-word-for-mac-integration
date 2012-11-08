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

Components.utils.import("resource://gre/modules/XPCOMUtils.jsm");
Components.utils.import("resource://gre/modules/ctypes.jsm");

var Zotero = Components.classes["@zotero.org/Zotero;1"]
			.getService(Components.interfaces.nsISupports)
			.wrappedJSObject;
var field_t, document_t, fieldListNode_t, progressFunction_t, lib, libPath, f, fieldPtr;
var dataInUse = [];

/**
 * Loads libZoteroMacWordIntegration.dylib and initializes js-ctypes functions
 */
function init() {
	if(lib) return;
	var ios = Components.classes["@mozilla.org/network/io-service;1"]  
		.getService(Components.interfaces.nsIIOService);
	var resHandler = ios.getProtocolHandler("resource")
		.QueryInterface(Components.interfaces.nsIResProtocolHandler);
	var fileURI = resHandler.getSubstitution("zotero-macword-integration")
		.QueryInterface(Components.interfaces.nsIFileURL);
	libPath = fileURI.file;
	libPath.append("libzoteroMacWordIntegration.dylib");
	
	lib = ctypes.open(libPath.path);
	
	document_t = new ctypes.StructType("document_t");
	
	field_t = new ctypes.StructType("field_t", [
		{ "code":ctypes.char.ptr },
		{ "text":ctypes.char.ptr },
		{ "noteType":ctypes.unsigned_short },
		{ "entryIndex":ctypes.long },
		{ "bookmarkName":ctypes.char.ptr },
		{ "sbField":ctypes.voidptr_t },
		{ "sbBookmark":ctypes.voidptr_t }
		// There's more here, but we will never access it, and we do not create field_t objects
		// from JavaScript
	]);
	
	fieldListNode_t = new ctypes.StructType("fieldListNode_t");
	fieldListNode_t.define([
		{ "field":field_t.ptr },
		{ "next":fieldListNode_t.ptr }
	]);
	
	progressFunction_t = new ctypes.FunctionType(ctypes.default_abi, ctypes.void_t,
		[ctypes.int]).ptr;
	
	var statusCode = ctypes.unsigned_short;
	f = {
		// char* getError(void);
		"getError":lib.declare("getError", ctypes.default_abi, ctypes.char.ptr),
		
		// void clearError(void);
		"clearError":lib.declare("clearError", ctypes.default_abi, ctypes.void_t),
		
		// statusCode getDocument(bool isWord2004, const char* wordPath,
		//					   const char* documentName, Document** returnValue);
		"getDocument":lib.declare("getDocument", ctypes.default_abi, statusCode, ctypes.bool,
			ctypes.char.ptr, ctypes.char.ptr, document_t.ptr.ptr),
		
		// void freeDocument(Document *doc);
		"freeDocument":lib.declare("freeDocument", ctypes.default_abi, statusCode, document_t.ptr),
			
		// statusCode activate(Document *doc);
		"activate":lib.declare("activate", ctypes.default_abi, statusCode, document_t.ptr),
		
		// statusCode displayAlert(char const dialogText[], unsigned short icon,
		//						   unsigned short buttons, unsigned short* returnValue);
		"displayAlert":lib.declare("displayAlert", ctypes.default_abi, ctypes.unsigned_short,
			document_t.ptr, ctypes.char.ptr, ctypes.unsigned_short, ctypes.unsigned_short,
			ctypes.unsigned_short.ptr),
		
		// statusCode canInsertField(Document *doc, const char fieldType[], bool* returnValue);
		"canInsertField":lib.declare("canInsertField", ctypes.default_abi, statusCode,
			document_t.ptr, ctypes.char.ptr, ctypes.bool.ptr),
		
		// statusCode cursorInField(Document *doc, const char fieldType[], Field** returnValue);
		"cursorInField":lib.declare("cursorInField", ctypes.default_abi, statusCode, document_t.ptr,
			ctypes.char.ptr, field_t.ptr.ptr),
		
		// statusCode getDocumentData(Document *doc, char **returnValue);
		"getDocumentData":lib.declare("getDocumentData", ctypes.default_abi, statusCode,
			document_t.ptr, ctypes.char.ptr.ptr),
		
		// statusCode setDocumentData(Document *doc, const char documentData[]);
		"setDocumentData":lib.declare("setDocumentData", ctypes.default_abi, statusCode,
			document_t.ptr, ctypes.char.ptr),
		
		// statusCode insertField(Document *doc, const char fieldType[],
		//					      unsigned short noteType, Field **returnValue)
		"insertField":lib.declare("insertField", ctypes.default_abi, statusCode, document_t.ptr,
			ctypes.char.ptr, ctypes.unsigned_short, field_t.ptr.ptr),
		
		// statusCode getFields(document_t *doc, const char fieldType[],
		//					    fieldListNode_t** returnNode);
		"getFields":lib.declare("getFields", ctypes.default_abi, statusCode, document_t.ptr,
			ctypes.char.ptr, fieldListNode_t.ptr.ptr),
		
		// statusCode getFieldsAsync(document_t *doc, const char fieldType[],
		// 						     void (*onProgress)(int progress),
		// 						     fieldListNode_t** returnNode);
		"getFieldsAsync":lib.declare("getFieldsAsync", ctypes.default_abi, statusCode,
			document_t.ptr, ctypes.char.ptr, fieldListNode_t.ptr.ptr, progressFunction_t),
		
		// statusCode setBibliographyStyle(Document *doc, long firstLineIndent, 
		//								   long bodyIndent, unsigned long lineSpacing,
		//								   unsigned long entrySpacing, long tabStops[],
		//								   unsigned long tabStopCount);
		"setBibliographyStyle":lib.declare("setBibliographyStyle", ctypes.default_abi,
			statusCode, document_t.ptr, ctypes.long, ctypes.long, ctypes.unsigned_long,
			ctypes.unsigned_long, ctypes.long.array(), ctypes.unsigned_long),
		
		// statusCode convert(document_t *doc, field_t* fields[], unsigned long nFields,
		//				      const char toFieldType[], unsigned short noteType[]);
		"convert":lib.declare("convert", ctypes.default_abi, statusCode, document_t.ptr,
			field_t.ptr.ptr, ctypes.unsigned_long, ctypes.char.ptr, ctypes.unsigned_short.ptr),
		
		// statusCode cleanup(Document *doc);
		"cleanup":lib.declare("cleanup", ctypes.default_abi, statusCode, document_t.ptr),
		
		// statusCode deleteField(Field* field);
		"deleteField":lib.declare("deleteField", ctypes.default_abi, statusCode, field_t.ptr),
			
		// statusCode removeCode(Field* field);
		"removeCode":lib.declare("removeCode", ctypes.default_abi, statusCode, field_t.ptr),
			
		// statusCode selectField(Field* field);
		"selectField":lib.declare("selectField", ctypes.default_abi, statusCode, field_t.ptr),
			
		// statusCode setText(Field* field, const char string[], bool isRich);
		"setText":lib.declare("setText", ctypes.default_abi, statusCode, field_t.ptr,
			ctypes.char.ptr, ctypes.bool),
			
		// statusCode getText(Field* field, char** returnValue);
		"getText":lib.declare("getText", ctypes.default_abi, statusCode, field_t.ptr,
			ctypes.char.ptr.ptr),
			
		// statusCode setCode(Field *field, const char code[]);
		"setCode":lib.declare("setCode", ctypes.default_abi, statusCode, field_t.ptr,
			ctypes.char.ptr),
		
		// statusCode getNoteIndex(Field* field, unsigned long *returnValue);
		"getNoteIndex":lib.declare("getNoteIndex", ctypes.default_abi, statusCode,
			field_t.ptr, ctypes.unsigned_long.ptr),
		
		// statusCode install(const char templatePath[]);
		"install":lib.declare("install", ctypes.default_abi, statusCode, ctypes.char.ptr),
		
		// statusCode getScriptItemsDirectory(char** scriptFolder);
		"getScriptItemsDirectory":lib.declare("getScriptItemsDirectory", ctypes.default_abi,
			statusCode, ctypes.char.ptr.ptr),
		
		// statusCode writeScript(char* scriptPath, char* scriptContent);
		"writeScript":lib.declare("writeScript", ctypes.default_abi, statusCode, ctypes.char.ptr,
			ctypes.char.ptr),
		
		// statusCode freeData(void* ptr);
		"freeData":lib.declare("freeData", ctypes.default_abi, statusCode, ctypes.void_t.ptr)
	};
	
	fieldPtr = new ctypes.PointerType(field_t);
}

/**
 * Gets the last error that took place in C code.
 */
function getLastError() {
	var errPtr = f.getError();
	if(errPtr.isNull()) {
		var err = "An unexpected error occurred.";
	} else {
		var err = errPtr.readString().replace("\u2019", "'", "g");
	}
	f.clearError();
	return err;
}

/**
 * Checks the return status of a function to verify that no error occurred.
 * @param {Integer} status The return status code of a C function
 */
function checkStatus(status) {
	if(!status) return;
	
	if(status === 1) {
		throw(getLastError());
	} else {
		throw "ExceptionAlreadyDisplayed";
	}
}

/**
 * Ensures that the document associated with this object has not been garbage collected
 */
function checkIfFreed(documentStatus) {
	if(!documentStatus.active) throw "complete() method already called on document";
}

/**
 * Handles installation of Zotero Word for Mac Integration scripts and template file.
 */
var Installer = function() {
	init();
	this.wrappedJSObject = this;
};
Installer.prototype = {
	classDescription: "Zotero Word for Mac Integration Installer",
	classID:		Components.ID("{aa56c6c0-95f0-48c2-b223-b11b96b9c9e5}"),
	contractID:		"@zotero.org/Zotero/integration/installer?agent=MacWord;1",
	QueryInterface: XPCOMUtils.generateQI([Components.interfaces.nsISupports,
		Components.interfaces.nsIRunnable]),
	"service":		true,
	"run":function() {
		init();
		var template = libPath.parent.parent;
		template.append("install");
		template.append("Zotero.dot");
		checkStatus(f.install(template.path));
	},
	"getScriptItemsDirectory":function() {
		var returnValue = new ctypes.char.ptr();
		checkStatus(f.getScriptItemsDirectory(returnValue.address()));
		var outString = returnValue.readString();
		f.freeData(returnValue);
		return outString;
	},
	"writeScript":function(scriptPath, scriptContent) {
		checkStatus(f.writeScript(scriptPath, scriptContent));
	}
};

var Application2004 = function() {};
Application2004.prototype = {
	classDescription: "Zotero Word 2004 for Mac Integration Application",
	classID:		Components.ID("{b063dd87-5615-45c5-ac3d-4b0583034616}"),
	contractID:		"@zotero.org/Zotero/integration/application?agent=MacWord2004;1",
	QueryInterface: XPCOMUtils.generateQI([Components.interfaces.nsISupports,
		Components.interfaces.zoteroIntegrationApplication]),
	"service":		true,
	"getDocument":function(path) {
		init();
		var docPtr = new document_t.ptr();
		checkStatus(f.getDocument(true, path, null, docPtr.address()));
		return new Document(docPtr);
	},
	"getActiveDocument":function(path) {
		return this.getDocument(null);
	},
	"primaryFieldType":"Field",
	"secondaryFieldType":"Bookmark"
};

var Application2008 = function() {};
Application2008.prototype = {
	classDescription: "Zotero Word 2008+ for Mac Integration Application",
	classID:		Components.ID("{ea584d70-2797-4cd1-8015-1a5f5fb85af7}"),
	contractID:		"@zotero.org/Zotero/integration/application?agent=MacWord2008;1",
	QueryInterface: XPCOMUtils.generateQI([Components.interfaces.nsISupports,
		Components.interfaces.zoteroIntegrationApplication]),
	"service":		true,
	"getDocument":function(path) {
		init();
		var docPtr = new document_t.ptr();
		checkStatus(f.getDocument(false, path, null, docPtr.address()));
		return new Document(docPtr);
	},
	"getActiveDocument":function(path) {
		return this.getDocument(null);
	},
	"primaryFieldType":"Field",
	"secondaryFieldType":"Bookmark"
};

/**
 * See zoteroIntegration.idl
 */
var Document = function(cDoc) {
	this._document_t = cDoc;
	this._documentStatus = {"active":true};
	this.wrappedJSObject = this;
};
Document.prototype = {
	QueryInterface: XPCOMUtils.generateQI([Components.interfaces.nsISupports,
		Components.interfaces.zoteroIntegrationDocument]),
	
	"displayAlert":function(dialogText, icon, buttons) {
		Zotero.debug("ZoteroMacWordIntegration: displayAlert", 4);
		var buttonPressed = new ctypes.unsigned_short();
		checkStatus(f.displayAlert(this._document_t, dialogText, icon, buttons,
			buttonPressed.address()));
		return buttonPressed.value;
	},
	
	"activate":function() {
		Zotero.debug("ZoteroMacWordIntegration: activate", 4);
		checkIfFreed(this._documentStatus);
		checkStatus(f.activate(this._document_t));
	},
	
	"canInsertField":function(fieldType) {
		Zotero.debug("ZoteroMacWordIntegration: canInsertField", 4);
		checkIfFreed(this._documentStatus);
		var returnValue = new ctypes.bool();
		checkStatus(f.canInsertField(this._document_t, fieldType, returnValue.address()));
		return returnValue.value;
	},
	
	"cursorInField":function(fieldType) {
		Zotero.debug("ZoteroMacWordIntegration: cursorInField", 4);
		checkIfFreed(this._documentStatus);
		var returnValue = new field_t.ptr();
		checkStatus(f.cursorInField(this._document_t, fieldType, returnValue.address()));
		return (returnValue.isNull() ? null : new Field(returnValue, this._documentStatus));
	},
	
	"getDocumentData":function() {
		Zotero.debug("ZoteroMacWordIntegration: getDocumentData", 4);
		checkIfFreed(this._documentStatus);
		var returnValue = new ctypes.char.ptr();
		checkStatus(f.getDocumentData(this._document_t, returnValue.address()));
		var data = returnValue.readString();
		f.freeData(returnValue);
		return data;
	},
	
	"setDocumentData":function(documentData) {
		Zotero.debug("ZoteroMacWordIntegration: setDocumentData", 4);
		checkIfFreed(this._documentStatus);
		checkStatus(f.setDocumentData(this._document_t, documentData));
	},
	
	"insertField":function(fieldType, noteType) {
		Zotero.debug("ZoteroMacWordIntegration: insertField", 4);
		checkIfFreed(this._documentStatus);
		var returnValue = new field_t.ptr();
		checkStatus(f.insertField(this._document_t, fieldType, noteType, returnValue.address()));
		return new Field(returnValue, this._documentStatus);
	},
	
	"getFields":function(fieldType) {
		Zotero.debug("ZoteroMacWordIntegration: getFields", 4);
		checkIfFreed(this._documentStatus);
		var fieldListNode = new fieldListNode_t.ptr();
		checkStatus(f.getFields(this._document_t, fieldType, fieldListNode.address()));
		return new FieldEnumerator(fieldListNode, this._documentStatus);
	},
	
	"getFieldsAsync":function(fieldType, observer) {
		Zotero.debug("ZoteroMacWordIntegration: getFieldsAsync", 4);
		checkIfFreed(this._documentStatus);
		var documentStatus = this._documentStatus,
			callback = progressFunction_t(function(progress) {
			// Remove global reference that prevents GC
			dataInUse.splice(dataInUse.indexOf(callback), 2);
			
			if(progress == -1) {
				observer.observe(getLastError(), "fields-error", null);
			} else if(progress == 100) {
				observer.observe(new FieldEnumerator(fieldListNode, documentStatus),
					"fields-available", null);
			} else {
				observer.observe(progress, "fields-progress", null);
			}
		});
		var fieldListNode = new fieldListNode_t.ptr();
		
		// Prevent GC
		dataInUse = dataInUse.concat([callback, fieldListNode]);
		
		var me = this;
		Zotero.setTimeout(function() {
			checkStatus(f.getFieldsAsync(me._document_t, fieldType, fieldListNode.address(),
				callback));
		}, 0);
	},
	
	"setBibliographyStyle":function(firstLineIndent, bodyIndent, lineSpacing, entrySpacing,
			tabStops) {
		Zotero.debug("ZoteroMacWordIntegration: setBibliographyStyle", 4);
		checkIfFreed(this._documentStatus);
		checkStatus(f.setBibliographyStyle(this._document_t, firstLineIndent, bodyIndent, lineSpacing,
			entrySpacing, ctypes.long.array(tabStops.length)(tabStops), tabStops.length));
	},
	
	"convert":function(fieldEnumerator, toFieldType, toNoteTypes, nFields) {
		Zotero.debug("ZoteroMacWordIntegration: convert", 4);
		checkIfFreed(this._documentStatus);
		var fieldPointers = [];
		while(fieldEnumerator.hasMoreElements()) {
			fieldPointers.push(fieldEnumerator.getNext().wrappedJSObject._field_t);
		}
		checkStatus(f.convert(this._document_t, field_t.ptr.array()(fieldPointers),
			fieldPointers.length, ctypes.char.array()(toFieldType),
			ctypes.unsigned_short.array()(toNoteTypes)));
	},
	
	"cleanup":function() {
		Zotero.debug("ZoteroMacWordIntegration: cleanup", 4);
		if(this._documentStatus.active) {
			checkStatus(f.cleanup(this._document_t));
		} else {
			Zotero.debug("complete() already called on document; ignoring");
		}
	},
	
	"complete":function() {
		Zotero.debug("ZoteroMacWordIntegration: complete", 4);
		if(this._documentStatus.active) {
			f.freeDocument(this._document_t);
			this._documentStatus.active = false;
		} else {
			Zotero.debug("complete() already called on document; ignoring");
		}
	}
};

/**
 * An enumerator implementation to handle passing off fields
 */
var FieldEnumerator = function(startNode, documentStatus) {
	this._currentNode = startNode;
	this._documentStatus = documentStatus;
};
FieldEnumerator.prototype = {
	"hasMoreElements":function() {
		checkIfFreed(this._documentStatus);
		return !this._currentNode.isNull();
	},
	
	"getNext":function() {
		checkIfFreed(this._documentStatus);
		var contents = this._currentNode.contents;
		var fieldPtr = contents.addressOfField("field").contents;
		this._currentNode = contents.addressOfField("next").contents;
		return new Field(fieldPtr, this._documentStatus);
	},
	
	"QueryInterface": XPCOMUtils.generateQI([Components.interfaces.nsISupports,
		Components.interfaces.nsISimpleEnumerator])
};

/**
 * See zoteroIntegration.idl
 */
var Field = function(field_t, documentStatus) {
	this._field_t = field_t;
	this._isBookmark = field_t.contents.addressOfField("sbField").contents.isNull();
	this._documentStatus = documentStatus;
	this.wrappedJSObject = this;
};
Field.prototype = {
	QueryInterface: XPCOMUtils.generateQI([Components.interfaces.nsISupports,
		Components.interfaces.zoteroIntegrationField]),
	
	"delete":function() {
		Zotero.debug("ZoteroMacWordIntegration: delete", 4);
		checkIfFreed(this._documentStatus);
		checkStatus(f.deleteField(this._field_t));
	},
	
	"removeCode":function() {
		Zotero.debug("ZoteroMacWordIntegration: removeCode", 4);
		checkIfFreed(this._documentStatus);
		checkStatus(f.removeCode(this._field_t));
	},
	
	"select":function() {
		Zotero.debug("ZoteroMacWordIntegration: select", 4);
		checkIfFreed(this._documentStatus);
		checkStatus(f.selectField(this._field_t));
	},
	
	"setText":function(text, isRich) {
		Zotero.debug("ZoteroMacWordIntegration: setText", 4);
		checkIfFreed(this._documentStatus);
		checkStatus(f.setText(this._field_t, text, isRich));
	},
	
	"getText":function(text) {
		Zotero.debug("ZoteroMacWordIntegration: getText", 4);
		checkIfFreed(this._documentStatus);
		var returnValue = new ctypes.char.ptr();
		checkStatus(f.getText(this._field_t, returnValue.address()));
		return returnValue.readString();
	},
	
	"setCode":function(code) {
		Zotero.debug("ZoteroMacWordIntegration: setCode", 4);
		checkIfFreed(this._documentStatus);
		checkStatus(f.setCode(this._field_t, code));
	},
	
	"getCode":function(code) {
		Zotero.debug("ZoteroMacWordIntegration: getCode", 4);
		checkIfFreed(this._documentStatus);
		return this._field_t.contents.addressOfField("code").contents.readString();
	},
	
	"equals":function(field) {
		Zotero.debug("ZoteroMacWordIntegration: equals", 4);
		checkIfFreed(this._documentStatus);
		// Obviously, a field cannot be equal to a bookmark
		if(this._isBookmark !== field.wrappedJSObject._isBookmark) return false;
		
		if(this._isBookmark) {
			return this._field_t.contents.addressOfField("bookmarkName").contents.readString() ===
				field.wrappedJSObject._field_t.contents.addressOfField("bookmarkName").contents.readString();
		} else {
			var a = this._field_t.contents,
				b = field.wrappedJSObject._field_t.contents;
			// This is stupid.
			return a.addressOfField("noteType").contents.toString()
				=== b.addressOfField("noteType").contents.toString()
				&& a.addressOfField("entryIndex").contents.toString()
				=== b.addressOfField("entryIndex").contents.toString();
		}
	},
	
	"getNoteIndex":function(field) {
		Zotero.debug("ZoteroMacWordIntegration: getNoteIndex", 4);
		checkIfFreed(this._documentStatus);
		var returnValue = new ctypes.unsigned_long();
		checkStatus(f.getNoteIndex(this._field_t, returnValue.address()));
		return returnValue.value;
	}
}

const NSGetFactory = XPCOMUtils.generateNSGetFactory([
	Installer,
	Application2004,
	Application2008
]);