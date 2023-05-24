/*
	***** BEGIN LICENSE BLOCK *****
	
	Copyright © 2021 Corporation for Digital Scholarship
                     Vienna, Virginia, USA
					http://zotero.org
	
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

let MessagingGeneric = class {
	/**
	 * Set up messaging between two isolated JS contexts (frames, context scripts, background pages,
	 * web workers or anything else).
	 * @param options {Object}
	 * @param options.sendMessage {Function} [required]
	 * 		A function that invokes the messageListeners on the remote JS context with arguments
	 * 		
	 * @param options.addMessageListener {Function} [required]
	 * 		Add a handler for a sendMessage call from the remote JS context
	 * 		
	 * @param options.functionOverrides {Object}
	 * 		An object specifying which local function calls should be created/overriden and invoked
	 * 		as a function call at the remote JS context
	 * 		
	 * @param options.handlerFunctionOverrides {Object}
	 * 		An object specifying which function calls the remote will try to call in our space.
	 * 		Used to configure argument pre/post processing before and after running the function.
	 *
	 * @param options.overrideTarget {Object} [default=self]
	 * 		The root object for which function overrides are applied. Defaults to `self`, which is the
	 * 		root object of the context. Would generally be set to Zotero.
	 * 		
	 * @param options.supportsError {Boolean}
	 * 		Whether the sendMessage function supports sending errors without serialization.
	 * 		
	 * @param options.supportsResponse {Boolean}
	 * 		Whether the sendMessage function returns an async response and the message listeners
	 * 		expect an async response to be returned to remote JS context.
	 */
	 
	constructor(options={}) {
		if (!options.sendMessage || !options.addMessageListener) {
			throw new Error('Zotero.MessagingGeneric: mandatory constructor options missing');
		}
		if (options.sendMessage == 'frame') {
			options.sendMessage = (...args) => window.parent.postMessage(args, '*');
		}
		if (options.addMessageListener == 'frame') {
			options.addMessageListener = (fn) => window.addEventListener('message', (messageEvent) => {
				if (messageEvent.data && Array.isArray(messageEvent.data)) {
					fn(messageEvent.data);
				}
			});
		}
		this._options = options;
		this._sendMessage = options.sendMessage;
		this._addMessageListener = options.addMessageListener;
		this._functionOverrides = options.functionOverrides;
		this._handlerFunctionOverrides = options.handlerFunctionOverrides;
		this._messageListeners = {};
		this._responseListeners = {};
		this._overrideTarget = options.overrideTarget || self;

		// Handle function overrides
		for (const fnName in this._functionOverrides) {
			const fnPath = fnName.split('.');
			let fn = this._overrideTarget;
			for (let name of fnPath.slice(0, fnPath.length - 1)) {
				if (!(name in fn)) {
					fn[name] = {};
				}
				fn = fn[name];
			}
			fn[fnPath[fnPath.length - 1]] = async (...args) => {
				try {
					const messageConfig = this._functionOverrides[fnName];
					if (messageConfig.local && messageConfig.local.preSend) {
						args = await messageConfig.local.preSend(args);
					}
					
					let response = await this.sendMessage(fnName, args);
					if (messageConfig.local && messageConfig.local.postReceive) {
						response = await messageConfig.local.postReceive(response);
					}
					return response;
				}
				catch (e) {
					console.log(`Zotero: Failed to invoke a remote function ${fnName}`);
					console.log(e)
					throw e;
				}
			};
		}
		// And handler function overrides towards us
		for (let fnName in this._handlerFunctionOverrides) {
			this.addMessageListener(fnName, async (...args) => {
				const fnPath = fnName.split('.');
				const messageConfig = this._handlerFunctionOverrides[fnName];
				let result;
				try {
					let fn = this._overrideTarget;
					for (let name of fnPath) {
						fn = fn[name];
					}
					if (messageConfig.handler && messageConfig.handler.postReceive) {
						args = await messageConfig.handler.postReceive(args);
					}
					result = fn(...args);
					if (typeof result === 'object' && result.then) {
						result = await result;
					}
				}
				catch(e) {
					console.log(`Zotero: Failed to invoke a function call from a handler destination ${fnName}`);
					console.log(e);
					throw new Error(`Zotero: Failed to invoke a function call from a handler destination ${fnName}`);
				}
				if (messageConfig.handler && messageConfig.handler.preSend) {
					result = await messageConfig.handler.preSend(result);
				}
				return result;
			});
		}
		
		this._initMessageListener();
	}
	
	// Initialize message handler
	_initMessageListener() {
		this._addMessageListener(async (args) => {
			if (!Array.isArray(args) || args.length > 3) return;
			let [message, payload, messageId] = args;
			if (this._messageListeners[message]) {
				let result;
				try {
					result = this._messageListeners[message](...payload);
					if (typeof result === 'object' && result.then) result = await result;
				} catch (e) {
					if (this._options.supportsError) {
						result = e;
					}
					else {
						result = ['error', JSON.stringify(Object.assign({
							name: e.name,
							message: e.message,
							stack: e.stack
						}, e))];
					}
				}
				if (this._options.supportsResponse) {
					return result;
				}
				else {
					this._sendMessage(`response`, result, messageId);
				}
			}
			else if (this._responseListeners[messageId]) {
				let listener = this._responseListeners[messageId];
				delete this._responseListeners[messageId];
				listener(payload);
			}
		});
	}
	
	
	_randomString(len, chars) {
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

	
	async sendMessage(message, payload) {
		let response;
		if (this._options.supportsResponse) {
			 response = await this._sendMessage(message, payload)
		}
		else {
			response = await new Promise((resolve) => {
				const id = this._randomString();
				this._responseListeners[id] = resolve;
				this._sendMessage(message, payload, id);
			});
		}
		if (response instanceof Error) {
			throw response;
		}
		if (Array.isArray(response) && response[0] == 'error') {
			response[1] = JSON.parse(response[1]);
			let e = new Error(response[1].message);
			for (let key in response[1]) e[key] = response[1][key];
			throw e;
		}
		return response;
	}
	
	addMessageListener(message, listener) {
		this._messageListeners[message] = listener;
	}
};