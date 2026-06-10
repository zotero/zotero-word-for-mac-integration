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
	this.LAST_INSTALLED_FILE_UPDATE = "9.0.0pre";
	
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
			const isSequoiaOrLater = macOSVersion.split('.')[0] >= 15;
			const isGoldenGateOrLater = macOSVersion.split('.')[0] >= 27;
			const userDoesNotWantToBeAskedAgain = Zotero.Utilities.semverCompare(dontAskAgainVersion, this.LAST_INSTALLED_FILE_UPDATE) >= 0;
			// On macOS 27 (Golden Gate) and later, writing into another app's group container
			// is denied by default and no longer triggers a permission prompt, but a folder
			// grant from a previous install persists, so check whether we can already write to
			// the startup folder and skip all prompting if so
			const hasStartupFolderAccess = isGoldenGateOrLater && await this.checkStartupFolderAccess();
			if (!zpi.force && isSequoiaOrLater && !hasStartupFolderAccess) {
				if (userDoesNotWantToBeAskedAgain) return;
				const shouldProceed = await this.displayPermissionWarningBanner(isGoldenGateOrLater);
				if (!shouldProceed) return;
				// Since the user confirmed that they want to install the plugin
				// we should never fail silently here, especially since they might then
				// deny access to required file location
				zpi.failSilently = false;
			}
			// Get com.apple.macl access to the Word startup folder via a folder-selection
			// dialog before running the installer
			if (isGoldenGateOrLater && !hasStartupFolderAccess
					&& !(await this.requestStartupFolderAccess())) {
				throw new Error("ExceptionAlreadyDisplayed");
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
	
	this.getWordStartupFolder = function() {
		return PathUtils.join(
			Services.dirsvc.get('Home', Ci.nsIFile).path,
			'Library', 'Group Containers', 'UBF8T346G9.Office',
			'User Content.localized', 'Startup.localized', 'Word'
		);
	}

	// Check whether we can already write to the Word startup folder (e.g., from a folder
	// grant during a previous install). Unlike on Sequoia through Tahoe, a denied write on
	// macOS 27 and later fails silently rather than triggering a permission prompt, so this
	// is safe to do without warning the user first.
	this.checkStartupFolderAccess = async function() {
		const testPath = PathUtils.join(this.getWordStartupFolder(), '.zotero-install-test');
		try {
			await IOUtils.writeUTF8(testPath, '');
			await IOUtils.remove(testPath);
			return true;
		}
		catch (e) {
			zoteroPluginInstaller.debug(`Word startup folder write check failed -- assuming no access (${e.message})`);
			return false;
		}
	}

	// Access to a folder that the user has picked in a folder-selection dialog is allowed even when
	// access to another app's group container is otherwise denied, so show a dialog pre-filled with
	// the Word startup folder to make the installer's writes there possible. The grant is recorded
	// as a com.apple.macl extended attribute on the folder. According to
	// https://eclecticlight.co/2026/04/21/the-macl-extended-attribute/, macl grants were originally
	// permanent (other than by deleting the folder or disabling SIP), but as of 26.4, app UUIDs
	// were being regenerated on restart, which invalidated the grants. Testing on Golden Gate
	// Developer Beta 1 shows continued access after a system restart, though, so grants appear to
	// again be permanent.
	this.requestStartupFolderAccess = async function() {
		zoteroPluginInstaller.debug('Requesting Word startup folder access via folder-selection dialog');
		const { FilePicker } = ChromeUtils.importESModule('chrome://zotero/content/modules/filePicker.mjs');
		const startupDir = this.getWordStartupFolder();
		while (true) {
			let fp = new FilePicker();
			fp.init(
				Zotero.getMainWindow(),
				// The title isn't displayed in panels on modern macOS, but it may still be read
				// via VoiceOver
				Zotero.getString('mac-word-plugin-install-folder-dialog-title'),
				fp.modeGetFolder
			);
			fp.okButtonLabel = Zotero.getString('mac-word-plugin-install-folder-dialog-button');
			fp.displayDirectory = startupDir;
			let returnValue = await fp.show();
			if (returnValue != fp.returnOK) {
				zoteroPluginInstaller.debug('User cancelled the folder-selection dialog');
				return false;
			}
			if (fp.file == startupDir) {
				return true;
			}
			zoteroPluginInstaller.debug(`Folder-selection dialog returned ${fp.file} instead of ${startupDir}`);
			const ps = Services.prompt;
			const buttonFlags = ps.BUTTON_POS_0 * ps.BUTTON_TITLE_OK
				+ ps.BUTTON_POS_1 * ps.BUTTON_TITLE_CANCEL;
			const index = ps.confirmEx(null, this.EXTENSION_STRING,
				Zotero.getString('mac-word-plugin-install-wrong-folder-selected'),
				buttonFlags, null, null, null, null, {});
			if (index == 1) {
				return false;
			}
		}
	}

	this.displayPermissionWarningBanner = async function(folderAccess) {
		zoteroPluginInstaller.debug('Displaying a permission warning banner');
		const remindInterval = 60 * 60 * 24; // Remind again in 24 hours
		const lastDisplayed = zoteroPluginInstaller.prefBranch.getIntPref('installationWarning.lastDisplayed');
		if (lastDisplayed > Math.round(Date.now() / 1000) - remindInterval) {
			return false;
		}
		let zp = Zotero.getActiveZoteroPane()
		let result = await zp.showMacWordPluginInstallWarning({ folderAccess })
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
