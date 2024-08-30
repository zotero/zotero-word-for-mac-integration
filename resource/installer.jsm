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
const { Zotero } = ChromeUtils.importESModule("chrome://zotero/content/zotero.mjs");
var ZoteroPluginInstaller = Components.utils.import("resource://zotero/word-processor-plugin-installer.js").ZoteroPluginInstaller;
var Installer = function(failSilently=true, force) {
	return new ZoteroPluginInstaller(Plugin, failSilently, force);
}

var Plugin = new function() {
	this.EXTENSION_STRING = "Zotero Word for Mac Integration";
	this.EXTENSION_ID = "zoteroMacWordIntegration@zotero.org";
	this.EXTENSION_PREF_BRANCH = "extensions.zoteroMacWordIntegration.";
	this.EXTENSION_DIR = "zotero-macword-integration";
	this.APP = 'Microsoft Word';
	this.VERSION_FILE = 'resource://zotero-macword-integration/version.txt';

	// Bump to make Zotero update the template (Zotero.dotm) for existing installs. Do not remove "pre"
	this.LAST_INSTALLED_FILE_UPDATE = "7.0.5pre";
	
	var zoteroPluginInstaller;
	
	this.install = async function(zpi) {
		zoteroPluginInstaller = zpi;
		
		Zotero.debug("Installing ZoteroMacWordIntegration");
		try {
			const { Installer } = ChromeUtils.importESModule('chrome://zotero-macword-integration/content/zoteroMacWordIntegration.mjs');
			var installer = new Installer();
			await installer.run();
			zoteroPluginInstaller.success();
		} catch(e) {
			const message = e.toString();
			if (message.includes("ExceptionAlreadyDisplayed")) {
				Services.prompt.alert(null, this.EXTENSION_STRING,
					"You cancelled installation of Zotero Word for Mac Integration. To install later, visit the Cite pane in the Zotero preferences.");
				zoteroPluginInstaller.cancelled();
			}
			else if (!zpi.force && message.includes("Word does not appear to be installed on this computer.")) {
				// Do not display this as an error if not installing via the preferences window
				zoteroPluginInstaller.success();
			}
			else {
				zoteroPluginInstaller.error("Installation could not be completed because an error occurred.\n\n"+e);
				throw e;
			}
		}
	}
}
