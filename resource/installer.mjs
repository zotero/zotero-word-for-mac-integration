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

import { Zotero } from "chrome://zotero/content/zotero.mjs";

const { ZoteroPluginInstaller } = ChromeUtils.importESModule("resource://zotero/word-processor-plugin-installer.mjs");

export var Installer = function(failSilently=true, force) {
	return new ZoteroPluginInstaller(Plugin, failSilently, force);
};

var Plugin = new (function() {
	this.EXTENSION_STRING = "Zotero Word for Mac Integration";
	this.EXTENSION_ID = "zoteroMacWordIntegration@zotero.org";
	this.EXTENSION_PREF_BRANCH = "extensions.zoteroMacWordIntegration.";
	this.EXTENSION_DIR = "zotero-macword-integration";
	this.APP = 'Microsoft Word';
	this.VERSION_FILE = 'resource://zotero-macword-integration/version.txt';
	this.DISABLE_PROGRESS_WINDOW = true;

	// Bump to make Zotero update the template (Zotero.dotm) for existing installs. Do not remove "pre"
	this.LAST_INSTALLED_FILE_UPDATE = "7.0.6pre";
	
	var zoteroPluginInstaller;
	
	this.install = async function(zpi) {
		zoteroPluginInstaller = zpi;
		
		Zotero.debug("Installing ZoteroMacWordIntegration");
		try {
			const { Installer } = ChromeUtils.importESModule('chrome://zotero-macword-integration/content/zoteroMacWordIntegration.mjs');
			const installer = new Installer();
			const isWordInstalled = await installer.isWordInstalled();
			if (!isWordInstalled) return;
			const macOSVersion = (await Zotero.getOSVersion()).split(' ')[1];
			const dontAskAgainVersion = zpi.prefBranch.getCharPref('installationWarning.dontAskAgainVersion');
			const isSequoia = macOSVersion.split('.')[0] >= 15;
			const userDoesNotWantToBeAskedAgain = Zotero.Utilities.semverCompare(dontAskAgainVersion, this.LAST_INSTALLED_FILE_UPDATE) >= 0;
			if (!zpi.force && isSequoia) {
				if (userDoesNotWantToBeAskedAgain) return;
				const shouldProceed = await this.displayPermissionWarningBanner();
				if (!shouldProceed) return;
				// Since the user confirmed that they want to install the plugin
				// we should never fail silently here, especially since they might then
				// deny access to required file location
				zpi.failSilently = false;
			}
			await installer.run();
			zoteroPluginInstaller.success();
		} catch(e) {
			const message = e.toString();
			if (message.includes("ExceptionAlreadyDisplayed")) {
				Services.prompt.alert(null, this.EXTENSION_STRING,
					"You cancelled installation of Zotero Word for Mac Integration. To install later, visit the Cite pane in the Zotero preferences.");
				zoteroPluginInstaller.cancelled();
			}
			else {
				zoteroPluginInstaller.error("Installation could not be completed because an error occurred.\n\n"+e);
				throw e;
			}
		}
	}
	
	this.displayPermissionWarningBanner = async function() {
		zoteroPluginInstaller.debug('Displaying a permission warning banner');
		const remindInterval = 60 * 60 * 24; // Remind again in 24 hours
		const lastDisplayed = zoteroPluginInstaller.prefBranch.getIntPref('installationWarning.lastDisplayed');
		if (lastDisplayed > Math.round(Date.now() / 1000) - remindInterval) {
			return false;
		}
		let zp = Zotero.getActiveZoteroPane()
		let result = await zp.showMacWordPluginInstallWarning()
		zoteroPluginInstaller.debug('User closed banner with ' + JSON.stringify(result));
		if (result.install) return true;
		else if (result.dismiss) return false;
		else if (result.dontAskAgain) {
			zoteroPluginInstaller.prefBranch.setCharPref('installationWarning.dontAskAgainVersion', zoteroPluginInstaller._currentPluginVersion)
			return false;
		}
		// Dismissed with remind later.
		zoteroPluginInstaller.prefBranch.setIntPref(`installationWarning.lastDisplayed`, Math.round(Date.now() / 1000));
	}
})
