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

Components.utils.import("resource://gre/modules/ComponentUtils.jsm");
Components.utils.import("resource://gre/modules/Services.jsm");
Components.utils.import("resource://gre/modules/FileUtils.jsm");

var Zotero = Components.classes["@zotero.org/Zotero;1"]
			.getService(Components.interfaces.nsISupports)
			.wrappedJSObject;
<<<<<<< HEAD
=======
var field_t, document_t, fieldListNode_t, progressFunction_t, lib, libPath, f, x, fn, fieldPtr;
var dataInUse = [];
var flushWordVersion = false;
var m1OSOSVersionChecked = false;
>>>>>>> fa74f09 (Removes old installation code for 2011 and earlier versions. Closes #28)

Components.classes["@mozilla.org/moz/jssubscript-loader;1"]
	.getService(Components.interfaces.mozIJSSubScriptLoader)
	.loadSubScript('resource://zotero-macword-integration/messagingGeneric.js');

var fn, worker, pipesInitialized = false;
var m1OSOSVersionChecked = false;
var Messaging;

/**
 * Loads load the webWorker that will issue commands to word asynchronously
 * so as not to block the main thread and UI
 */
async function init() {
	if (worker) return;
	let libPath = FileUtils.getDir('ARes', []).parent.parent;
	libPath.append('integration');
	libPath.append('word-for-mac');
	libPath.append("libZoteroWordIntegration.dylib");
	
<<<<<<< HEAD
	worker = new ChromeWorker("resource://zotero-macword-integration/webWorkerLibInterface.js")
	
	Messaging = new MessagingGeneric({
		sendMessage: (...args) => worker.postMessage(args),
		addMessageListener: (listener) => worker.addEventListener('message', (event) => {
			listener(event.data);
		}),
		handlerFunctionOverrides: {
			'debug': true,
		},
		overrideTarget: Zotero
	});
	
	fn = {}
	for (let method of ["preventAppNap", "allowAppNap",
			"install", "isWordArm", "getMacOSVersion", "isZoteroRosetta"]) {
		fn[method] = (...args) => Messaging.sendMessage(method, args);
=======
	lib = ctypes.open(libPath.path);
	
	document_t = new ctypes.StructType("Document");
	
	field_t = ctypes.unsigned_long;
	
	fieldListNode_t = new ctypes.StructType("fieldListNode_t");
	fieldListNode_t.define([
		{ field: field_t },
		{ next: fieldListNode_t.ptr }
	]);
	
	progressFunction_t = new ctypes.FunctionType(ctypes.default_abi, ctypes.void_t,
		[ctypes.int]).ptr;
	
	var statusCode = ctypes.unsigned_short;
	f = {
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
		
		// statusCode getFieldsAsync(document_t *doc, const char fieldType[],
		// 						     void (*onProgress)(int progress),
		// 						     fieldListNode_t** returnNode);
		getFieldsAsync: lib.declare("getFieldsAsync", ctypes.default_abi, statusCode,
			document_t.ptr, ctypes.char.ptr, fieldListNode_t.ptr.ptr, progressFunction_t),
		
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
	
		// statusCode freeData(void* ptr);
		freeData: lib.declare("freeData", ctypes.default_abi, statusCode, ctypes.void_t.ptr),

		// bool isWordArm();
		isWordArm: lib.declare("isWordArm", ctypes.default_abi, ctypes.bool),
		
		// char *getWordVersion(const char wordPath[]);
		getWordVersion: lib.declare("getWordVersion", ctypes.default_abi, ctypes.char.ptr, ctypes.char.ptr),
		
		// char *getMacOSVersion();
		getMacOSVersion: lib.declare("getMacOSVersion", ctypes.default_abi, ctypes.char.ptr),

		// void flushBundleCache(const char wordPath[]);
		flushWordVersion: lib.declare("flushBundleCache", ctypes.default_abi, ctypes.void_t, ctypes.char.ptr),
		
		// int isZoteroRosetta();
		isZoteroRosetta: lib.declare("isZoteroRosetta", ctypes.default_abi, ctypes.int),
	};
	
	fn = f;

	fieldPtr = new ctypes.PointerType(field_t);
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
function checkStatus(status, pre2016=false) {
	if(!status) return;

	if (status === STATUS_EXCEPTION) {
		throw Components.Exception(getLastError());
	} else {
		if (status === STATUS_EXCEPTION_SB_DENIED) {
			let message = Zotero.getString('integration.error.macWordSBPermissionsMissing');
			if (pre2016) {
				message += '\n\n' + Zotero.getString('integration.error.macWordSBPermissionsMissing.pre2016');
			}
			let index = displayMoreInformationAlert(
				Zotero.getString('integration.error.macWordSBPermissionsMissing.title'),
				message
			);
			if (index == 1) {
				Zotero.launchURL('https://www.zotero.org/support/kb/mac_word_permissions_missing')
			}
		}
		throw new Error("ExceptionAlreadyDisplayed");
>>>>>>> fa74f09 (Removes old installation code for 2011 and earlier versions. Closes #28)
	}
	
	await Messaging.sendMessage('init', [libPath.path])
}

/**
 * Displays a dialog with a More Information button
 * @param title
 * @param message
 * @param ignoreCheckboxValue
 * @returns {*}
 */
function displayMoreInformationAlert(title, message, ignoreCheckboxValue = null) {
	let ps = Services.prompt;
	let buttonFlags = ps.BUTTON_POS_0 * ps.BUTTON_TITLE_OK
		+ ps.BUTTON_POS_1 * ps.BUTTON_TITLE_IS_STRING;
	return ps.confirmEx(
		null,
		title,
		message,
		buttonFlags,
		null,
		Zotero.getString('general.moreInformation'),
		null,
		ignoreCheckboxValue
			? Zotero.getString(
				'general.dontShowAgainFor',
				armIsSupportedWarningIgnoreDays,
				armIsSupportedWarningIgnoreDays
			)
			: null,
		ignoreCheckboxValue || {}
	);
}

// Checks if the user is using a M1 mac with macOS 11.3 or lower
// and displays a warning that they should upgrade to
// macOS 11.4 or later or Word will freeze when using with Zotero
// (due to 2048+ character bug in AE messaging).
async function checkM1OSAndShowWarning() {
	if (m1OSOSVersionChecked) return;
	m1OSOSVersionChecked = true;

	try {
		var isZoteroRosetta = (await fn.isZoteroRosetta()) == 1;
		if (!isZoteroRosetta) return;
		Zotero.debug('MacWord: M1 Mac detected. Checking macOS version.')

		var macOSVersion = await fn.getMacOSVersion();
		if (!macOSVersion.length) {
			Zotero.logError(`Failed to check macOS version: ${getLastError()}`);
			return;
		}

		Zotero.debug(`MacWord: macOS version: ${macOSVersion}`);
		macOSVersion = macOSVersion.split('.').map(num => parseInt(num));
		Zotero.debug(`MacWord: parsed macOS version: ${macOSVersion}`);
		if (macOSVersion[0] >= 12 || (macOSVersion[0] == 11 && macOSVersion[1] >= 4)) return;

		Zotero.debug(`MacWord: macOS version below 11.4, displaying an upgrade warning`);
		var title = Zotero.getString('integration.error.m1UpgradeOS.title');
		var message = Zotero.getString(`integration.error.m1UpgradeOS`, Zotero.appName);
		var url = 'https://www.zotero.org/support/kb/mac_word_apple_silicon_compatibility';
		var index = displayMoreInformationAlert(title, message);
		if (index == 1) {
			Zotero.launchURL(url)
		}
	} catch (e) {
		Zotero.debug('An unexpected error occurred while trying to check M1 mac OS version or display a warning dialog')
		Zotero.logError(e);
	}
	throw new Error("ExceptionAlreadyDisplayed");
}

async function initializePipes() {
	// If worker had been initialized then so have the pipes
	if (pipesInitialized) return;
	await Zotero.initializationPromise;
	Zotero.debug("ZoteroMacWordIntegration: Initializing integration pipes");
	init2016Pipe();
	pipesInitialized = true;
}

async function init2016Pipe() {
	var office2016Container = Components.classes["@mozilla.org/file/directory_service;1"].
		getService(Components.interfaces.nsIProperties).
		get("Home", Components.interfaces.nsIFile);
	office2016Container.append("Library");
	office2016Container.append("Containers");
	office2016Container.append("com.microsoft.Word");
	office2016Container.append("Data");
	
	if(!office2016Container.exists() || !office2016Container.isDirectory() || !office2016Container.isWritable()) return;

	var pipe = office2016Container.clone();
	pipe.append(".zoteroIntegrationPipe");

	if(pipe.exists()) {
		if(!Zotero.Integration.deletePipe(pipe)) return;
	}
	
	// try to initialize pipe
	Zotero.Integration.initPipe(pipe);
}

/**
 * Handles installation of Zotero Word for Mac Integration scripts and template file.
 */
var Installer = function() {
	initializePipes();
	var Integration = Components.utils.import("resource://zotero-macword-integration/integration.js").MacWordIntegration;
	Integration.init();
	this.wrappedJSObject = this;
};
Installer.prototype = {
	classDescription: "Zotero Word for Mac Integration Installer",
	classID:		Components.ID("{aa56c6c0-95f0-48c2-b223-b11b96b9c9e5}"),
	contractID:		"@zotero.org/Zotero/integration/installer?agent=MacWord;1",
	QueryInterface: ChromeUtils.generateQI([Components.interfaces.nsISupports,
		Components.interfaces.nsIRunnable]),
	service: 		true,
	run: async function() {
		let zoteroDot = FileUtils.getDir('ARes', []).parent.parent;
		zoteroDot.append('integration');
		zoteroDot.append('word-for-mac');
		var zoteroDotm = zoteroDot.clone();
		var zoteroScpt = zoteroDot.clone();
		zoteroDot.append("Zotero.dot");
		zoteroDotm.append("Zotero.dotm");
		zoteroScpt.append("Zotero.scpt");
<<<<<<< HEAD
		checkStatus(await fn.install(zoteroDot.path, zoteroDotm.path, zoteroScpt.path));
=======
		checkStatus(f.install(zoteroDot.path, zoteroDotm.path, zoteroScpt.path));
>>>>>>> fa74f09 (Removes old installation code for 2011 and earlier versions. Closes #28)
	}
};

var Application2016 = function() {
	this.wrappedJSObject = this;
};
Application2016.prototype = {
	classDescription: "Zotero Word 2016 for Mac Integration Application",
	classID:		Components.ID("{9c6e787b-27d7-4567-98d4-b57d0afa3d8c}"),
	contractID:		"@zotero.org/Zotero/integration/application?agent=MacWord2016;1",
	QueryInterface: ChromeUtils.generateQI([Components.interfaces.nsISupports]),
	service: 		true,
	getDocument: async function(path) {
		await init();
		let docId = Messaging.sendMessage('getDocument', [2016, path])
		await fn.preventAppNap();
		return new Document(docId);
	},
	getActiveDocument: async function() {
		return this.getDocument(null);
	},
	primaryFieldType: "Field",
	secondaryFieldType: "Bookmark",
	supportedNotes: ['footnotes', 'endnotes'],
	supportsImportExport: true,
	supportsTextInsertion: true,
	outputFormat: "rtf",
	processorName: "Word"
};


// Word 16.0 and higher
var Application16 = function() {
	this.wrappedJSObject = this;
};
Application16.prototype = {
	classDescription: "Zotero Word 16.xx for Mac Integration Application",
	classID:		Components.ID("{0a5ec6de-f9eb-11e7-8c3f-9a214cf093ae}"),
	contractID:		"@zotero.org/Zotero/integration/application?agent=MacWord16;1",
	QueryInterface: ChromeUtils.generateQI([Components.interfaces.nsISupports]),
	service: 		true,
	getDocument: async function(path) {
		await init();
		await checkM1OSAndShowWarning();
		let docId = await Messaging.sendMessage('getDocument', [16, path]);
		await fn.preventAppNap();
		return new Document(docId);
	},
	getActiveDocument: async function() {
		return this.getDocument(null);
	},
	primaryFieldType: "Field",
	secondaryFieldType: "Bookmark",
	supportedNotes: ['footnotes', 'endnotes'],
	supportsImportExport: true,
	supportsTextInsertion: true,
	outputFormat: "rtf",
	processorName: "Word"
};


/**
 * See integrationTest.js
 */
var Document = function(docId) {
	this.id = docId;
};
for (let method of ["activate", "displayAlert", "canInsertField", "getDocumentData",
		"setDocumentData", "insertField", "setBibliographyStyle", "exportDocument",
		"importDocument", "insertText", "convert", "cleanup", "complete"]) {
	Document.prototype[method] = function(...args) {
		 return Messaging.sendMessage(method, [this.id, ...args]);
	};
}
Document.prototype.cursorInField = async function(...args) {
	let fieldId = await Messaging.sendMessage('cursorInField', [this.id, ...args])
	return fieldId ? new Field(this, fieldId) : fieldId;
}
Document.prototype.insertField = async function(...args) {
	let fieldId = await Messaging.sendMessage('insertField', [this.id, ...args])
	return new Field(this, fieldId);
}
Document.prototype.getFields = async function(...args) {
	let fieldIds = await Messaging.sendMessage('getFields', [this.id, ...args])
	return fieldIds.map(id => new Field(this, id));
}
Document.prototype.convertPlaceholdersToFields = async function(...args) {
	let fieldIds = await Messaging.sendMessage('convertPlaceholdersToFields', [this.id, ...args])
	return fieldIds.map(id => new Field(this, id));
}
Document.prototype.convert = async function(fields, ...args) {
	fields = fields.map(f => f.id);
	return Messaging.sendMessage('convert', [this.id, fields, ...args])
}

var Field = function(doc, fieldId) {
	this.doc = doc;
	this.id = fieldId;
}
for (let method of ["delete", "removeCode", "selectField", "setText", "getText", "setCode", "getCode",
		"getNoteIndex", "isAdjacentToNextField"]) {
	Field.prototype[method] = function (...args) {
		return Messaging.sendMessage(method, [this.doc.id, this.id, ...args]);
	};
}
Field.prototype.equals = function (other) {
	return Messaging.sendMessage('equals', [this.doc.id, this.id, other.id]);
}

const NSGetFactory = ComponentUtils.generateNSGetFactory([
	Installer,
	Application2016,
	Application16
]);