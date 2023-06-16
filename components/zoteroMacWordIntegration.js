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
	// We run all AppleScript/SBBridge automation in a ChromeWorker (which is a type of SharedWorker)
	// in Zotero 7. To run those, Zotero needs permissions to automate Word. However, macOS does not
	// prompt for those permissions when trying to automate Word from the ChromeWorker.
	// MacOS tracks "responsibility" and is supposed to generally know, that Zotero is responsible for the
	// ChromeWorker thread, but for some reason it doesn't. See https://developer.apple.com/forums/thread/731504
	// What we do here is execute a no-op at the start of a Word transaction to make sure the users are prompted
	// to provide Zotero with the permission.
	try {
		await Zotero.Utilities.Internal.exec('/usr/bin/osascript', ['-e', 'tell application "Microsoft Word" to time to GMT'])
	} catch (e) {}
	
	let libPath = FileUtils.getDir('ARes', []).parent.parent;
	libPath.append('integration');
	libPath.append('word-for-mac');
	libPath.append("libZoteroWordIntegration.dylib");
	
	worker = new ChromeWorker("resource://zotero-macword-integration/webWorkerLibInterface.js")
	
	Messaging = new MessagingGeneric({
		sendMessage: (...args) => worker.postMessage(args),
		addMessageListener: (listener) => worker.addEventListener('message', (event) => {
			listener(event.data);
		}),
		handlerFunctionOverrides: {
			'debug': true,
			'getString': true,
			'launchURL': true
		},
		overrideTarget: Zotero
	});
	
	Messaging.addMessageListener('displayMoreInformationAlert', (...args) => displayMoreInformationAlert(...args))
	
	fn = {}
	for (let method of ["preventAppNap", "allowAppNap",
			"install", "isWordArm", "getMacOSVersion", "isZoteroRosetta"]) {
		fn[method] = (...args) => Messaging.sendMessage(method, args);
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
	if (pipesInitialized) return;
	await Zotero.initializationPromise;
	let macOSVersion = await Zotero.getOSVersion();
	if (Zotero.Utilities.semverCompare('13.999', macOSVersion.split(' ')[1]) <= 0) {
		Zotero.debug(`ZoteroMacWordIntegration: ${macOSVersion} detected -- not initializing pipes`);
		return
	}
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
		await fn.install(zoteroDot.path, zoteroDotm.path, zoteroScpt.path);
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