(*
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
*)

try
	set W2008_SCRIPT_NAMES to {"Add Bibliography\\cob.scpt", Â
		"Add Citation\\coa.scpt", Â
		"Edit Bibliography\\cod.scpt", Â
		"Edit Citation\\coe.scpt", Â
		"Refresh\\cor.scpt", Â
		"Remove Field Codes.scpt", Â
		"Set Document Preferences\\cop.scpt"}
	set W2008_SCRIPT_COMMANDS to {"addBibliography", Â
		"addCitation", Â
		"editBibliography", Â
		"editCitation", Â
		"refresh", Â
		"removeCodes", Â
		"setDocPrefs"}
	
	-- See if Office 2008 is installed
	set installed2008 to false
	try
		application id "com.Microsoft.Word"
		set installed2008 to true
	end try
	
	if installed2008 then
		-- Look for Script Menu Items
		set scriptMenuItemsFolder to false
		try
			set scriptMenuItemsFolder to alias ((path to home folder as text) & "Documents:Microsoft User Data:Word Script Menu Items")
		end try
		if scriptMenuItemsFolder is false then
			try
				set scriptMenuItemsFolder to alias ((path to preferences folder from user domain as text) & "Microsoft User Data:Word Script Menu Items")
			end try
		end if
		if scriptMenuItemsFolder is false then
			try
				tell application "System Events"
					activate
					set scriptMenuItemsFolder to choose folder with prompt "Select the Word Script Menu Items folder, usually located within the Microsoft User Data folder." default location (path to documents folder)
				end tell
			end try
		end if
		
		if scriptMenuItemsFolder is not false then
			-- Create Scripts
			set posixPathToScripts to quoted form of POSIX path of scriptMenuItemsFolder & "/Zotero/"
			do shell script "rm -rf " & posixPathToScripts & "; mkdir " & posixPathToScripts
			repeat with i from 1 to length of W2008_SCRIPT_NAMES
				do shell script "osacompile -d -e \"try\" -e \"alias ((path to home folder as text) & \\\".zoteroIntegrationPipe\\\")\" -e \"do shell script \\\"echo 'MacWord2008 " & (item i of W2008_SCRIPT_COMMANDS) & "' > ~/.zoteroIntegrationPipe\\\"\" -e \"on error\" -e \"display alert \\\"Word could not communicate with Zotero. Please ensure that Firefox is open and try again.\\\" as critical\" -e \"end try\" -o " & posixPathToScripts & "'" & (item i of W2008_SCRIPT_NAMES) & "'"
			end repeat
		end if
	end if
	
	-- See if Office 2004 is installed
	set installed2004 to false
	try
		set installed2004 to alias ((path to applications folder as text) & "Microsoft Office 2004:Microsoft Word")
	end try
	if installed2004 is false then
		try
			set installed2004 to (path to application id "MSWD")
			-- If not an application bundle, assume 2004
			if last character of (installed2004 as text) is ":" then
				set installed2004 to false
			end if
		end try
	end if
	
	if installed2004 is not false then
		-- Create startup folder if it doesn't exist
		set AppleScript's text item delimiters to ":"
		set pathToOffice to (text items 1 thru -2 of (installed2004 as text) as text)
		set startupDirectory to (quoted form of POSIX path of pathToOffice) & "/Office/Startup/Word"
		do shell script "mkdir -p " & startupDirectory
		do shell script "cp " & (quoted form of POSIX path of (text items 1 thru -2 of (path to me as text) as text)) & "/Zotero.dot " & startupDirectory
		set dot to alias (pathToOffice & ":Office:Startup:Word:Zotero.dot")
		tell application "Finder"
			set creator type of dot to "MSWD"
			set file type of dot to "W8TN"
		end tell
	end if
on error err
	tell application "System Events"
		activate
		display alert "Zotero MacWord Integration could not be installed because an error occured." message err as critical
	end tell
	return
end try

if installed2004 is false and installed2008 is false then
	tell application "System Events"
		activate
		display alert "Zotero MacWord Integration could not be installed because an error occurred." message "Neither Word 2004 nor Word 2008 appear to be installed on this computer. Please install Word, then try again." as critical
	end tell
else
	tell application "System Events"
		set isOpen to every process whose name is "Microsoft Word"
	end tell
	if isOpen is not {} then
		do shell script "killall -s 'Microsoft Word'"
		tell application "System Events"
			activate
			display alert "Zotero MacWord Integration has been successfully installed, but Word must be restarted before it can be used." message "Please restart Word before continuing." as informational
		end tell
	end if
end if