/*
    ***** BEGIN LICENSE BLOCK *****
	
	Copyright (c) 2009  Zotero
	                    Center for History and New Media
						George Mason University, Fairfax, Virginia, USA
						http://zotero.org
	
	This program is free software: you can redistribute it and/or modify
	it under the terms of the GNU General Public License as published by
	the Free Software Foundation, either version 3 of the License, or
	(at your option) any later version.
	
	This program is distributed in the hope that it will be useful,
	but WITHOUT ANY WARRANTY; without even the implied warranty of
	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
	GNU General Public License for more details.
	
	You should have received a copy of the GNU General Public License
	along with this program.  If not, see <http://www.gnu.org/licenses/>.
    
    ***** END LICENSE BLOCK *****
*/

var ZoteroMacWordIntegration = new function() {
	this.EXTENSION_STRING = "Zotero MacWord Integration";
	this.EXTENSION_ID = "zoteroMacWordIntegration@zotero.org";
	this.EXTENSION_PREF_BRANCH = "extensions.zoteroMacWordIntegration.";
	this.EXTENSION_DIR = "zotero-macword-integration";
	this.APP = 'Microsoft Word';
	
	this.REQUIRED_ADDONS = [{
		name: "Zotero",
		url: "zotero.org",
		id: "zotero@chnm.gmu.edu",
		minVersion: "2.1a1.SVN"
	}, {
		name: "PythonExt",
		url: "zotero.org",
		id: "pythonext@mozdev.org",
		minVersion: "2.5"
	}];
	
	var zoteroPluginInstaller;
	
	this.verifyNotCorrupt = function(zpi) {
		zoteroPluginInstaller = zpi;
		
		var pythonExtComponents = zoteroPluginInstaller.getAddonPath("pythonext@mozdev.org");
		pythonExtComponents.append("components");
		if(!pythonExtComponents.directoryEntries.hasMoreElements()) {
			var prompt = Components.classes["@mozilla.org/embedcomp/prompt-service;1"]
				.getService(Components.interfaces.nsIPromptService)
				.confirm(null, 'Zotero MacWord Integration Error',
				'Zotero MacWord Integration requires PythonExt to run. While PythonExt is currently '+
				'installed, it appears to be corrupted or incompletely deleted. As such, the Firefox '+
				'Add-on manager may not be able to remove or reinstall it. Would you like Zotero '+
				'Integration to remove it for you and restart Firefox?\n\n'+
				'Upon restart, you may receive an error '+
				'stating that PythonExt is not installed, and you will need to reinstall PythonExt '+
				'before Zotero MacWord Integration will function properly.');
			if(prompt) {
				pythonExt.remove(true);
				restartFirefox();
			}
			throw "PythonExt is missing the extension helper; Zotero MacWord Integration will not function."
		}
		
		// see if component is registered
		if(!Components.classes["@zotero.org/Zotero/integration/application?agent=MacWord2008;1"]) {
			if(!zpi.failSilently) {
				if(Zotero.isFx4) {
					zpi.error("Zotero MacWord Integration is not properly loaded or registered. "+
						"Please ensure that you have the appropriate version of PythonExt installed "+
						"for your Firefox version.");
				} else {
					var prompt = Components.classes["@mozilla.org/embedcomp/prompt-service;1"]
						.getService(Components.interfaces.nsIPromptService)
						.confirm(null, 'Zotero MacWord Integration Error',
						'Zotero MacWord Integration installation succeeded, but a necessary '+
						'component does not appear to be loaded properly. \n\nZotero can attempt to correct '+
						'this for you. A Firefox restart will be required. Continue?');
					if(prompt) {
						clearComponentCache();
						restartFirefox();
					}
				}
			}
			throw "The Zotero MacWord Integration Python component is not registered."
		}
		
		if(Zotero.isFx4 && Zotero.oscpu === "Intel Mac OS X 10.5") {
			var err = "Zotero MacWord Integration is not compatible with Mac OS X 10.5 when run under "+
				(Zotero.isStandalone ? "Zotero Standalone" : "Firefox 4")+". Please upgrade to "+
				"Mac OS X 10.6.x, or downgrade to Firefox 3.6.x.";
			zpi.error(err);
			throw err;
		}
	}
	
	this.install = function(zpi) {
		zoteroPluginInstaller = zpi;
		
		try {
			var installScript = Components.classes["@zotero.org/Zotero/integration/installer?agent=MacWord;1"].
				createInstance(Components.interfaces.zoteroIntegrationInstaller);
			installScript.run(zpi.failSilently);
			zoteroPluginInstaller.success();
		} catch(e) {
			var message = "";
				
			var consoleService = Components.classes["@mozilla.org/consoleservice;1"]
				.getService(Components.interfaces.nsIConsoleService);
			
			var messages = {};
			consoleService.getMessageArray(messages, {});
			messages = messages.value;
			if(messages && messages.length) {
				var lastMessage = messages[messages.length-1];
				try {
					var error = lastMessage.QueryInterface(Components.interfaces.nsIScriptError);
				} catch(e2) {
					if(lastMessage.message && lastMessage.message.substr(0, 12) == "ERROR:xpcom:") {
						// print just the last line of the message, but re-throw the rest
						message = lastMessage.message.substr(0, lastMessage.message.length-1);
						message = "\n"+message.substr(message.lastIndexOf("\n"))
					}
				}
			}
			
			if(!message && typeof(e) == "object" && e.message) message = e.message;
			
			zoteroPluginInstaller.error();
			throw message;
		}
	}

	function restartFirefox() {	
		// The following code was borrowed from extensions.js
		// Notify all windows that an application quit has been requested.
		var os = Components.classes["@mozilla.org/observer-service;1"]
						   .getService(Components.interfaces.nsIObserverService);
		var cancelQuit = Components.classes["@mozilla.org/supports-PRBool;1"]
								   .createInstance(Components.interfaces.nsISupportsPRBool);
		os.notifyObservers(cancelQuit, "quit-application-requested", "restart");
		
		// Something aborted the quit process.
		if (cancelQuit.data)
		return;
		
		Components.classes["@mozilla.org/toolkit/app-startup;1"].getService(Components.interfaces.nsIAppStartup)
				  .quit(Components.interfaces.nsIAppStartup.eRestart | Components.interfaces.nsIAppStartup.eAttemptQuit);
	}

	function clearComponentCache() {
		// Delete relevant files
		var profileDirectory = Components.classes["@mozilla.org/file/directory_service;1"]
										 .getService(Components.interfaces.nsIProperties)
										 .get("ProfD", Components.interfaces.nsIFile);
		for each(var filename in ["compreg.dat", "extensions.cache", "xpti.dat", "extensions.rdf"]) {
			var file = profileDirectory.clone();
			file.append(filename);
			if(file.exists()) file.remove(false);
		}
	}
}