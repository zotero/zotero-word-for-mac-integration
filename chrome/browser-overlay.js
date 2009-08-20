/*
    ***** BEGIN LICENSE BLOCK *****
	
	Copyright (c) 2009  Center for History and New Media
						George Mason University, Fairfax, Virginia, USA
						http://chnm.gmu.edu
	
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

const ZOTEROMACWORDINTEGRATION_ID = "zoteroMacWordIntegration@zotero.org";
const ZOTEROMACWORDINTEGRATION_PREF = "extensions.zoteroMacWordIntegration.version";

function ZoteroMacWordIntegration_firstRun() {
	try {
		// run AppleScript to set up
		var installScript = Components.classes["@mozilla.org/extensions/manager;1"].
			getService(Components.interfaces.nsIExtensionManager).
			getInstallLocation(ZOTEROMACWORDINTEGRATION_ID).
			getItemLocation(ZOTEROMACWORDINTEGRATION_ID);
		installScript.append("install");
		installScript.append("install.scpt");
		
		var osascript = Components.classes["@mozilla.org/file/local;1"].
			createInstance(Components.interfaces.nsILocalFile);
		osascript.initWithPath("/usr/bin/osascript");
		var proc = Components.classes["@mozilla.org/process/util;1"].
				createInstance(Components.interfaces.nsIProcess);
		proc.init(osascript);
		proc.run(true, [installScript.path], 1);
	} catch(e) {
		var prompts = Components.classes["@mozilla.org/embedcomp/prompt-service;1"]
			.getService(Components.interfaces.nsIPromptService)
			.alert(null, 'Zotero Word Integration Error',
			'Zotero Word Integration could not complete installation because an error occurred. Please ensure that Word is closed, then restart Firefox.');
		throw e;
	}
}

var ext = Components.classes['@mozilla.org/extensions/manager;1']
   .getService(Components.interfaces.nsIExtensionManager).getItemForID(ZOTEROMACWORDINTEGRATION_ID);
var zoteroMacWordIntegration_prefService = Components.classes["@mozilla.org/preferences-service;1"].
	getService(Components.interfaces.nsIPrefBranch);
if(zoteroMacWordIntegration_prefService.getCharPref(ZOTEROMACWORDINTEGRATION_PREF) != ext.version) {
	ZoteroMacWordIntegration_firstRun();
	zoteroMacWordIntegration_prefService.setCharPref(ZOTEROMACWORDINTEGRATION_PREF, ext.version);
}