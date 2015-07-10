/*
    ***** BEGIN LICENSE BLOCK *****
	
	Copyright (c) 2009  Zotero
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

var EXPORTED_SYMBOLS = ["Installer"];
var Zotero = Components.classes["@zotero.org/Zotero;1"].getService(Components.interfaces.nsISupports).wrappedJSObject;
var ZoteroPluginInstaller = Components.utils.import("resource://zotero-macword-integration/installer_common.jsm").ZoteroPluginInstaller;
var Installer = function(failSilently, force) {
	return new ZoteroPluginInstaller(Plugin,
		failSilently !== undefined ? failSilently : Zotero.isStandalone,
		force);
}

var Plugin = new function() {
	this.EXTENSION_STRING = "Zotero Word for Mac Integration";
	this.EXTENSION_ID = "zoteroMacWordIntegration@zotero.org";
	this.EXTENSION_PREF_BRANCH = "extensions.zoteroMacWordIntegration.";
	this.EXTENSION_DIR = "zotero-macword-integration";
	this.APP = 'Microsoft Word';
	
	this.REQUIRED_ADDONS = [{
		name: "Zotero",
		url: "zotero.org",
		id: "zotero@chnm.gmu.edu",
		minVersion: "4.0.27.SOURCE",
		required: true
	}, {
		name: "Zotero LibreOffice Integration",
		url: "zotero.org/support/word_processor_plugin_installation",
		id: "zoteroOpenOfficeIntegration@zotero.org",
		minVersion: "3.5b2.SOURCE",
		required: false
	}];
	this.LAST_INSTALLED_FILE_UPDATE = "3.5.2";
	
	var zoteroPluginInstaller;
	
	this.verifyNotCorrupt = function(zpi) {}
	
	this.install = function(zpi) {
		zoteroPluginInstaller = zpi;
		
		Zotero.debug("Installing ZoteroMacWordIntegration");
		try {
			var installer = Components.classes["@zotero.org/Zotero/integration/installer?agent=MacWord;1"].
				createInstance(Components.interfaces.nsIRunnable);
			installer.run();
			zoteroPluginInstaller.success();
		} catch(e) {
			if(e.toString().indexOf("ExceptionAlreadyDisplayed") !== -1) {
				Components.classes["@mozilla.org/embedcomp/prompt-service;1"]
					.getService(Components.interfaces.nsIPromptService)
					.alert(null, this.EXTENSION_STRING,
					"You cancelled installation of Zotero Word for Mac Integration. To install later, visit the Cite pane in the Zotero preferences.");
				zoteroPluginInstaller.cancelled();
			} else {
				zoteroPluginInstaller.error("Installation could not be completed because an error occurred.\n\n"+e);
				throw e;
			}
		}
	}
}