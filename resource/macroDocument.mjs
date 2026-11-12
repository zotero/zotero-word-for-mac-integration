/*
	***** BEGIN LICENSE BLOCK *****

	Copyright © 2026 Corporation for Digital Scholarship
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

/**
 * Documents and fields implemented with the macros in the ZoteroMacros module of the Zotero
 * template (build/template/Zotero.dotm/word/vbaProject.bin/ZoteroMacros.bas), which also
 * describes the protocol.
 *
 * Word handles one Apple Event per display frame (17 ms on a 60 Hz display), and most operations
 * take many events with Scripting Bridge, so each document and field operation that needs Word is
 * one macro call here instead. Field codes, text and adjacency are returned when fields are read
 * and cached; note indices are refreshed in one call after operations that can add or remove
 * notes.
 *
 * Documents of templates without these macros use the Scripting Bridge implementation in
 * webWorkerLibInterface.js. The macros don't support bookmarks, so the first bookmark operation of a
 * command hands the document over to it: the macro session ends, and the worker dispatches the rest
 * of the command to a Scripting Bridge document with the same id.
 *
 * Besides calling macros, this uses the native activate() and displayAlert(), since a macro can't
 * show long messages, and freeDocument().
 */

const PROTOCOL_VERSION = 1;
const MACRO_NAME = 'ZoteroTemplate.ZoteroMacros.ZoteroCall';

const STATUS_OK = 0;
const STATUS_ERROR = 1;

const CLIPBOARD_NONE = 0;
const CLIPBOARD_RTF = 1;
const CLIPBOARD_HTML = 2;

// The error that ZoteroMacros.bas raises itself (vbObjectError + 1000)
const MACRO_ERROR = -2147221504 + 1000;

// Longest macro input logged
const MAX_LOGGED_INPUT = 100;

function encodeNumber(number) {
	return `${number};`;
}

function encodeString(string) {
	return `${string.length}:${string}`;
}

class Reader {
	constructor(data, macroName) {
		this.data = data;
		this.position = 0;
		this.macroName = macroName;
	}

	number() {
		let end = this.data.indexOf(';', this.position);
		let token = end == -1 ? '' : this.data.slice(this.position, end);
		if (!/^-?\d+$/.test(token)) this.malformed();
		this.position = end + 1;
		return parseInt(token);
	}

	string() {
		let separator = this.data.indexOf(':', this.position);
		let token = separator == -1 ? '' : this.data.slice(this.position, separator);
		let length = /^\d+$/.test(token) ? parseInt(token) : NaN;
		if (isNaN(length) || separator + 1 + length > this.data.length) this.malformed();
		let string = this.data.slice(separator + 1, separator + 1 + length);
		this.position = separator + 1 + length;
		return string;
	}

	malformed() {
		throw new Error(`Malformed output from Zotero template macro ${this.macroName}`);
	}
}

/**
 * See integrationTest.js
 */
export class MacroDocument {
	/**
	 * Starts a macro session on a document from the native openDocument(), or returns null if the
	 * template doesn't have the macros of this protocol version, in which case the caller keeps
	 * ownership of cDoc and uses Scripting Bridge. Documents that store citations as bookmarks
	 * switch to Scripting Bridge at their first bookmark operation.
	 *
	 * @param {Object} env See the constructor
	 * @param {document_t.ptr} cDoc
	 * @return {MacroDocument|null}
	 */
	static open(env, cDoc) {
		let doc = new MacroDocument(env, cDoc);
		let reader = doc._call('ZoteroBegin', encodeNumber(PROTOCOL_VERSION), { probe: true });
		if (!reader) {
			env.Zotero.debug('ZoteroMacWordIntegration: Template macros not available. Using Scripting Bridge.', 2);
			return null;
		}
		let templateVersion = reader.number();
		if (templateVersion != PROTOCOL_VERSION) {
			env.Zotero.debug(`ZoteroMacWordIntegration: The Zotero template uses macro protocol version `
				+ `${templateVersion}, but the plugin uses ${PROTOCOL_VERSION}. Using Scripting Bridge.`, 2);
			return null;
		}
		env.Zotero.debug('ZoteroMacWordIntegration: Using template macros', 2);
		env.docs[doc.id] = doc;
		return doc;
	}

	/**
	 * @param {Object} env What this module uses from the worker, which owns the native library:
	 *   fn (the native functions), ctypes, checkStatus(), randomString(), docs (the documents by
	 *   id), Zotero (for debug()), and openSBDocument(cDoc, id), which starts a Scripting Bridge
	 *   session on a document and registers a Scripting Bridge document for it under the id
	 * @param {document_t.ptr} cDoc The native document that macros run on
	 */
	constructor(env, cDoc) {
		this._env = env;
		this._cDoc = cDoc;
		// False after complete(), so that later calls fail instead of using a freed document
		this._active = true;
		// Note indices cached by fields are valid while this matches the field's
		this._noteVersion = 0;
		this.id = env.randomString();
		this.fields = {};
	}

	/**
	 * Runs an operation of the macros, a function in ZoteroMacros.bas that ZoteroCall runs, and
	 * returns a Reader for its result.
	 *
	 * @param {Object} [options]
	 * @param {String} [options.rtf] RTF to put on the clipboard while the macro runs
	 * @param {String} [options.html] HTML to put on the clipboard while the macro runs
	 * @param {Boolean} [options.probe] Return null instead of failing if the template doesn't
	 *   have the macros
	 */
	_call(name, args = '', { rtf, html, probe } = {}) {
		let { fn, ctypes } = this._env;
		if (!this._active) throw new Error("complete() method already called on document");
		let input = encodeString(name) + args;
		let clipboardFormat = CLIPBOARD_NONE;
		let clipboardData = null;
		if (rtf !== undefined) {
			clipboardFormat = CLIPBOARD_RTF;
			clipboardData = rtf;
		}
		else if (html !== undefined) {
			clipboardFormat = CLIPBOARD_HTML;
			clipboardData = html;
		}

		let startTime = Date.now();
		let returnValue = new ctypes.char.ptr();
		try {
			this._env.checkStatus(fn.callMacro(this._cDoc, MACRO_NAME, input, clipboardData,
				clipboardFormat, !!probe, returnValue.address()));
		}
		catch (e) {
			if (e.message == 'ExceptionAlreadyDisplayed') throw e;
			throw new Error(`Zotero template macro ${name}: ${e.message}`);
		}
		if (returnValue.isNull()) return null;
		let output = returnValue.readString();
		fn.freeData(returnValue);
		let loggedArgs = args.length > MAX_LOGGED_INPUT ? args.slice(0, MAX_LOGGED_INPUT) + '…' : args;
		this._env.Zotero.debug(`ZoteroMacWordIntegration: ${name}(${loggedArgs}) took ${Date.now() - startTime} ms`, 5);

		let reader = new Reader(output, name);
		let status = reader.number();
		if (status == STATUS_ERROR) {
			let errorNumber = reader.number();
			let description = reader.string();
			if (errorNumber == MACRO_ERROR) {
				// Raised by the macros themselves, so the description says it all
				throw new Error(description);
			}
			throw new Error(`Word error ${errorNumber} in ${name}: ${description}`);
		}
		if (status != STATUS_OK) {
			throw new Error(`Zotero template macro ${name} returned an unknown status ${status}`);
		}
		return reader;
	}

	_free() {
		if (!this._cDoc) return;
		this._env.fn.freeDocument(this._cDoc);
		this._cDoc = null;
		this._active = false;
	}

	_readField(reader) {
		let field = new MacroField(this, reader.number());
		field._noteIndex = reader.number();
		field._noteVersion = this._noteVersion;
		field._adjacent = reader.number() != 0;
		field._code = reader.string();
		field._text = reader.string();
		return field;
	}

	_readFields(reader) {
		let fields = [];
		let count = reader.number();
		for (let i = 0; i < count; i++) {
			fields.push(this._readField(reader));
		}
		return fields;
	}

	// Called after operations that can add or remove notes, which changes note indices
	_notesChanged() {
		this._noteVersion++;
	}

	_refreshNoteIndices() {
		let reader = this._call('ZoteroGetNoteIndices');
		let noteIndices = new Map();
		let count = reader.number();
		for (let i = 0; i < count; i++) {
			let handle = reader.number();
			noteIndices.set(handle, reader.number());
		}
		for (let field of Object.values(this.fields)) {
			if (!(field instanceof MacroField) || !noteIndices.has(field._handle)) continue;
			field._noteIndex = noteIndices.get(field._handle);
			field._noteVersion = this._noteVersion;
		}
	}

	/**
	 * Ends the macro session and returns a Scripting Bridge document that continues the command on
	 * the same native document. The macros don't support bookmarks, so the first bookmark operation
	 * does this. The worker dispatches later calls to the Scripting Bridge document, which takes
	 * this document's id; fields read with macros can't be used after that.
	 */
	_handOver() {
		this._env.Zotero.debug('ZoteroMacWordIntegration: Continuing with Scripting Bridge to use bookmarks', 2);
		try {
			this._call('ZoteroComplete');
		}
		catch (e) {
			this._free();
			throw e;
		}
		let cDoc = this._cDoc;
		this._cDoc = null;
		this._active = false;
		try {
			return this._env.openSBDocument(cDoc, this.id);
		}
		catch (e) {
			this._env.fn.freeDocument(cDoc);
			throw e;
		}
	}

	displayAlert(dialogText, icon, buttons) {
		let { fn, ctypes } = this._env;
		let buttonPressed = new ctypes.unsigned_short();
		this._env.checkStatus(fn.displayAlert(this._cDoc, dialogText, icon, buttons,
			buttonPressed.address()));
		return buttonPressed.value;
	}

	activate() {
		this._env.checkStatus(this._env.fn.activate(this._cDoc));
	}

	canInsertField(fieldType) {
		if (fieldType == 'Bookmark') return this._handOver().canInsertField(fieldType);
		return this._call('ZoteroCanInsertField').number() != 0;
	}

	cursorInField(fieldType) {
		if (fieldType == 'Bookmark') return this._handOver().cursorInField(fieldType);
		let reader = this._call('ZoteroCursorInField');
		return reader.number() ? this._readField(reader) : null;
	}

	getDocumentData() {
		return this._call('ZoteroGetDocumentData').string();
	}

	setDocumentData(documentData) {
		this._call('ZoteroSetDocumentData', encodeString(documentData));
	}

	insertField(fieldType, noteType) {
		if (fieldType == 'Bookmark') return this._handOver().insertField(fieldType, noteType);
		let reader = this._call('ZoteroInsertField', encodeNumber(noteType));
		if (noteType) this._notesChanged();
		return this._readField(reader);
	}

	getFields(fieldType) {
		if (fieldType == 'Bookmark') return this._handOver().getFields(fieldType);
		return this._readFields(this._call('ZoteroGetFields'));
	}

	setBibliographyStyle(firstLineIndent, bodyIndent, lineSpacing, entrySpacing, tabStops) {
		let args = [firstLineIndent, bodyIndent, lineSpacing, entrySpacing, tabStops.length,
			...tabStops].map(n => encodeNumber(Math.trunc(n))).join('');
		this._call('ZoteroSetBibliographyStyle', args);
	}

	importDocument(fieldType) {
		if (fieldType == 'Bookmark') return this._handOver().importDocument(fieldType);
		return this._call('ZoteroImportDocument').number() != 0;
	}

	exportDocument(fieldType, importInstructions) {
		if (fieldType == 'Bookmark') return this._handOver().exportDocument(fieldType, importInstructions);
		this._call('ZoteroExportDocument', encodeString(importInstructions));
		this._notesChanged();
	}

	insertText(text) {
		this._call('ZoteroInsertText', '', { html: text });
		this._notesChanged();
	}

	convertPlaceholdersToFields(placeholderIDs, noteType, fieldType) {
		if (fieldType == 'Bookmark') {
			return this._handOver().convertPlaceholdersToFields(placeholderIDs, noteType, fieldType);
		}
		let args = encodeNumber(placeholderIDs.length) + placeholderIDs.map(encodeString).join('')
			+ encodeNumber(noteType);
		let fields = this._readFields(this._call('ZoteroConvertPlaceholdersToFields', args));
		if (noteType) this._notesChanged();
		return fields;
	}

	convert(fields, toFieldType, toNoteTypes) {
		if (toFieldType == 'Bookmark') {
			// Scripting Bridge can't use these fields, so it converts all fields of the document:
			// bibliographies in the text, and citations to the note type requested for them
			let citationIndex = fields.findIndex(field => !field._code.startsWith('BIBL'));
			let citationNoteType = citationIndex == -1 ? 0 : toNoteTypes[citationIndex];
			let sb = this._handOver();
			let sbFields = sb.getFields('Field');
			let sbNoteTypes = sbFields.map(field => (field.getCode().startsWith('BIBL') ? 0 : citationNoteType));
			return sb.convert(sbFields, toFieldType, sbNoteTypes);
		}
		let args = encodeNumber(fields.length)
			+ fields.map((field, i) => encodeNumber(field._handle) + encodeNumber(toNoteTypes[i])).join('');
		this._call('ZoteroConvert', args);
		this._notesChanged();
	}

	cleanup() {
		if (!this._active) {
			this._env.Zotero.debug("complete() already called on document; ignoring", 4);
			return;
		}
		this._call('ZoteroCleanup');
	}

	complete() {
		if (!this._active) {
			this._env.Zotero.debug("complete() already called on document; ignoring", 4);
			return;
		}
		try {
			this._call('ZoteroComplete');
		}
		finally {
			this._free();
			this._env.fn.allowAppNap();
			delete this._env.docs[this.id];
		}
	}
}

/**
 * See integrationTest.js
 */
class MacroField {
	constructor(doc, handle) {
		this._doc = doc;
		this._handle = handle;
		this._code = '';
		this._text = '';
		this._noteIndex = 0;
		this._noteVersion = -1;
		this._adjacent = false;
		this.id = doc._env.randomString();
		doc.fields[this.id] = this;
	}

	_call(name, args, options) {
		return this._doc._call(name, encodeNumber(this._handle) + args, options);
	}

	delete() {
		this._call('ZoteroDeleteField', '');
		this._doc._notesChanged();
	}

	removeCode() {
		this._call('ZoteroRemoveCode', '');
	}

	select() {
		this._call('ZoteroSelectField', '');
	}

	setText(text, isRich) {
		let isBibliography = this._code.startsWith('BIBL');
		let reader = this._call('ZoteroSetText', encodeNumber(isRich ? 1 : 0)
			+ encodeNumber(isBibliography ? 1 : 0) + encodeString(isRich ? '' : text),
		isRich ? { rtf: text } : {});
		this._text = reader.string();
	}

	getText() {
		return this._text;
	}

	setCode(code) {
		this._call('ZoteroSetCode', encodeString(code));
		this._code = code;
	}

	getCode() {
		return this._code;
	}

	equals(field) {
		return field instanceof MacroField && field._doc == this._doc && field._handle == this._handle;
	}

	getNoteIndex() {
		if (this._noteVersion != this._doc._noteVersion) {
			this._doc._refreshNoteIndices();
			if (this._noteVersion != this._doc._noteVersion) {
				throw new Error('Field has been deleted');
			}
		}
		return this._noteIndex;
	}

	isAdjacentToNextField() {
		return this._adjacent;
	}
}
