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
var field_t, document_t, fieldListNode_t, lib, libPath, f, fieldPtr;
function init() {
	if(lib) return;
	var ios = Components.classes["@mozilla.org/network/io-service;1"]  
		.getService(Components.interfaces.nsIIOService);
	var resHandler = ios.getProtocolHandler("resource")
		.QueryInterface(Components.interfaces.nsIResProtocolHandler);
	var fileURI = resHandler.getSubstitution("zotero-macword-integration")
		.QueryInterface(Components.interfaces.nsIFileURL);
	libPath = fileURI.file;
	libPath.append("libZoteroMacWordIntegration.dylib");
	
	lib = ctypes.open(libPath.path);
	
	document_t = new ctypes.StructType("document_t");
	
	field_t = new ctypes.StructType("field_t", [
		{ "code":ctypes.char.ptr },
		{ "text":ctypes.char.ptr },
		{ "sbApp":ctypes.voidptr_t },
		{ "sbField":ctypes.voidptr_t },
		{ "sbDoc":ctypes.voidptr_t },
		{ "sbCodeRange":ctypes.voidptr_t },
		{ "sbContentRange":ctypes.voidptr_t },
		{ "noteType":ctypes.unsigned_short },
		{ "entryIndex":ctypes.long },
		{ "textLocation":ctypes.long },
		{ "noteLocation":ctypes.long }
	]);
	
	fieldListNode_t = new ctypes.StructType("fieldListNode_t");
	fieldListNode_t.define([
		{ "field":field_t.ptr },
		{ "next":fieldListNode_t.ptr }
	]);
	
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
		
		// statusCode displayAlert(char const dialogText[], unsigned short icon,
		//						   unsigned short buttons, unsigned short* returnValue);
		"displayAlert":lib.declare("displayAlert", ctypes.default_abi, ctypes.unsigned_short,
			ctypes.char.ptr, ctypes.unsigned_short, ctypes.unsigned_short,
			ctypes.unsigned_short.ptr),
		
		// void freeDocument(Document *doc);
		"freeDocument":lib.declare("freeDocument", ctypes.default_abi, statusCode, document_t.ptr),
			
		// statusCode activate(Document *doc);
		"activate":lib.declare("activate", ctypes.default_abi, statusCode, document_t.ptr),
		
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
		
		// void freeFieldList(fieldListNode_t *fieldList);
		"freeFieldList":lib.declare("freeFieldList", ctypes.default_abi, statusCode,
			fieldListNode_t.ptr),
		
		// statusCode setBibliographyStyle(Document *doc, long firstLineIndent, 
		//								   long bodyIndent, unsigned long lineSpacing,
		//								   unsigned long entrySpacing, long tabStops[],
		//								   unsigned long tabStopCount);
		"setBibliographyStyle":lib.declare("setBibliographyStyle", ctypes.default_abi,
			statusCode, document_t.ptr, ctypes.long, ctypes.long, ctypes.unsigned_long,
			ctypes.unsigned_long, ctypes.long.array(), ctypes.unsigned_long),
		
		// statusCode cleanup(Document *doc);
		"cleanup":lib.declare("cleanup", ctypes.default_abi, statusCode, document_t.ptr),
		
		// void freeFields(field_t* fields[], unsigned long nFields);
		"freeFields":lib.declare("freeFields", ctypes.default_abi, statusCode, field_t.ptr.ptr,
			ctypes.unsigned_long),
		
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
		"install":lib.declare("install", ctypes.default_abi, statusCode, ctypes.char.ptr)
	};
	
	fieldPtr = new ctypes.PointerType(field_t);
}

function checkStatus(status) {
	if(!status) return;
	
	if(status === 1) {
		var errPtr = f.getError();
		if(errPtr.isNull()) {
			var err = "An unexpected error occurred.";
		} else {
			var err = errPtr.readString().replace("\u2019", "'", "g");
		}
		f.clearError();
		throw(err);
	} else {
		throw "ExceptionAlreadyDisplayed";
	}
}

var Installer = function() {
	init();
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
	"primaryFieldType":"Field",
	"secondaryFieldType":"Bookmark"
};

var Application = function() {};
Application.prototype = {
	classDescription: "Zotero Word for Mac Integration Application",
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
	"primaryFieldType":"Field",
	"secondaryFieldType":"Bookmark"
};

/**
 * See zoteroIntegration.idl
 */
var Document = function(cDoc) {
	this._cDoc = cDoc;
	this._fieldPointers = [];
	this._fieldListPointers = [];
	this.wrappedJSObject = this;
};
Document.prototype = {
	QueryInterface: XPCOMUtils.generateQI([Components.interfaces.nsISupports,
		Components.interfaces.zoteroIntegrationDocument]),
	
	"displayAlert":function(dialogText, icon, buttons) {
		var buttonPressed = new ctypes.unsigned_short();
		checkStatus(f.displayAlert(dialogText, icon, buttons, buttonPressed.address()));
		return buttonPressed.contents;
	},
	
	"activate":function() {
		checkStatus(f.activate(this._cDoc));
	},
	
	"canInsertField":function(fieldType) {
		var returnValue = new ctypes.bool();
		checkStatus(f.canInsertField(this._cDoc, fieldType, returnValue.address()));
		return returnValue.value;
	},
	
	"cursorInField":function(fieldType) {
		var returnValue = new field_t.ptr();
		checkStatus(f.cursorInField(this._cDoc, fieldType, returnValue.address()));
		
		if(returnValue.isNull()) {
			return null;
		} else {
			this._fieldPointers.push(returnValue);
			return new Field(returnValue);
		}
	},
	
	"getDocumentData":function() {
		var returnValue = new ctypes.char.ptr();
		checkStatus(f.getDocumentData(this._cDoc, returnValue.address()));
		return returnValue.readString();
	},
	
	"setDocumentData":function(documentData) {
		checkStatus(f.setDocumentData(this._cDoc, documentData));
	},
	
	"insertField":function(fieldType, noteType) {
		var returnValue = new field_t.ptr();
		checkStatus(f.insertField(this._cDoc, fieldType, noteType, returnValue.address()));
		this._fieldPointers.push(returnValue);
		return new Field(returnValue);
	},
	
	"getFields":function(fieldType) {
		var fieldListNode = new fieldListNode_t.ptr();
		checkStatus(f.getFields(this._cDoc, fieldType, fieldListNode.address()));
		
		var fieldPointers = []
		var currentNode = fieldListNode;
		while(!currentNode.isNull()) {
			var contents = currentNode.contents;
			fieldPointers.push(contents.addressOfField("field").contents);
			currentNode = contents.addressOfField("next").contents;
		}
		this._fieldPointers = this._fieldPointers.concat(fieldPointers);
		this._fieldListPointers.push(fieldListNode);
		
		return new FieldEnumerator(fieldPointers);
	},
	
	"getFieldsAsync":function(fieldType, observer) {
		observer.observe(this.getFields(fieldType), "fields-available", null);
	},
	
	"setBibliographyStyle":function(firstLineIndent, bodyIndent, lineSpacing, entrySpacing,
			tabStops) {
		checkStatus(f.setBibliographyStyle(this._cDoc, firstLineIndent, bodyIndent, lineSpacing,
			entrySpacing, ctypes.long.array(tabStops.length)(tabStops), tabStops.length));
	},
	
	"cleanup":function() {
		checkStatus(f.cleanup(this._cDoc));
	},
	
	"complete":function() {
		if(this._fieldPointers.length) {
			f.freeFields(field_t.ptr.array()(this._fieldPointers), this._fieldPointers.length);
		}
		for(var i=0; i<this._fieldListPointers; i++) f.freeFieldList(this._fieldListPointers[i]);
		f.freeDocument(this._cDoc);
	}
};

/**
 * An enumerator implementation to handle passing off fields
 */
var FieldEnumerator = function(fieldPointers) {
	this._fieldPointers = fieldPointers;
};
FieldEnumerator.prototype = {
	"hasMoreElements":function() {
		return !!this._fieldPointers.length;
	},
	
	"getNext":function() {
		return new Field(this._fieldPointers.shift());
	},
	
	"QueryInterface": XPCOMUtils.generateQI([Components.interfaces.nsISupports,
		Components.interfaces.nsISimpleEnumerator])
};

/**
 * See zoteroIntegration.idl
 */
var Field = function(field_t) {
	this._field_t = field_t;
	this.wrappedJSObject = this;
};
Field.prototype = {
	QueryInterface: XPCOMUtils.generateQI([Components.interfaces.nsISupports,
		Components.interfaces.zoteroIntegrationField]),
	
	"delete":function() {
		checkStatus(f.deleteField(this._field_t));
	},
	
	"removeCode":function() {
		checkStatus(f.removeCode(this._field_t));
	},
	
	"select":function() {
		checkStatus(f.selectField(this._field_t));
	},
	
	"setText":function(text, isRich) {
		checkStatus(f.setText(this._field_t, text, isRich));
	},
	
	"getText":function(text) {
		var returnValue = new ctypes.char.ptr();
		checkStatus(f.getText(this._field_t, returnValue.address()));
		return returnValue.readString();
	},
	
	"setCode":function(code) {
		checkStatus(f.setCode(this._field_t, code));
	},
	
	"getCode":function(code) {
		return this._field_t.contents.addressOfField("code").contents.readString();
	},
	
	"equals":function(field) {
		var a = this._field_t.contents,
			b = field.wrappedJSObject._field_t.contents;
		// This is stupid.
		return a.addressOfField("noteType").contents.toString()
			== b.addressOfField("noteType").contents.toString()
			&& a.addressOfField("entryIndex").contents.toString()
			== b.addressOfField("entryIndex").contents.toString();
	},
	
	"getNoteIndex":function(field) {
		var returnValue = new ctypes.unsigned_long();
		checkStatus(f.getNoteIndex(this._field_t, returnValue.address()));
		return returnValue.value;
	}
}

const NSGetFactory = XPCOMUtils.generateNSGetFactory([
	Installer,
	Application
]);