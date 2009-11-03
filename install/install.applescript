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
	set W2008_SCRIPT_NAMES to {"Add Bibliography\\cob.scpt", "Add Citation\\coa.scpt", "Edit Bibliography\\cod.scpt", "Edit Citation\\coe.scpt", "Refresh\\cor.scpt", "Remove Field Codes.scpt", "Set Document Preferences\\cop.scpt"}
	set W2008_SCRIPT_COMMANDS to {"addBibliography", "addCitation", "editBibliography", "editCitation", "refresh", "removeCodes", "setDocPrefs"}
	
	-- See if Office 2008 is installed
	set installed2008 to false
	try
		tell application "Finder" to file "Microsoft Word.app" of folder "Microsoft Office 2008" of (path to applications folder)
		set installed2008 to true
	end try
	if not installed2008 then
		try
			run script "application id \"com.Microsoft.Word\""
			set installed2008 to true
		end try
	end if
	
	if installed2008 then
		-- Look for Script Menu Items
		set scriptMenuItemsFolder to false
		try
			tell application "Finder" to set scriptMenuItemsFolder to (folder "Word Script Menu Items" of folder "Microsoft User Data" of (path to documents folder)) as alias
		end try
		if scriptMenuItemsFolder is false then
			try
				tell application "Finder" to set scriptMenuItemsFolder to (folder "Word Script Menu Items" of folder "Microsoft User Data" of (path to preferences folder from user domain)) as alias
			end try
		end if
		if scriptMenuItemsFolder is false then
			try
				tell application "System Events"
					activate
					set scriptMenuItemsFolder to choose folder with prompt "Select the Word Script Menu Items folder, usually located in Documents/Microsoft User Data" default location (path to documents folder)
				end tell
			end try
		end if
		
		if scriptMenuItemsFolder is not false then
			-- Create Scripts
			set posixPathToScripts to quoted form of POSIX path of scriptMenuItemsFolder & "/Zotero/"
			do shell script "rm -rf " & posixPathToScripts & "; mkdir " & posixPathToScripts
			repeat with i from 1 to length of W2008_SCRIPT_NAMES
				do shell script "osacompile -d -e \"try\" -e \"do shell script \\\"ls ~/.zoteroIntegrationPipe && echo 'MacWord2008 " & (item i of W2008_SCRIPT_COMMANDS) & "' > ~/.zoteroIntegrationPipe\\\"\" -e \"on error\" -e \"display alert \\\"Word could not communicate with Zotero. Please ensure that Firefox is open and try again.\\\" as critical\" -e \"end try\" -o " & posixPathToScripts & "'" & (item i of W2008_SCRIPT_NAMES) & "'"
			end repeat
		end if
	end if
	
	-- See if Office 2004 is installed
	set installed2004 to false
	try
		tell application "Finder" to set installed2004 to (file "Microsoft Word" of folder "Microsoft Office 2004" of (path to applications folder)) as alias
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
		tell application "Finder"
			set pathToOffice to container of installed2004
			
			-- Get Startup directory, creating it if necessary. On some systems, this is "Start"
			try
				set startupDirectory to (folder "Word" of folder "Startup" of folder "Office" of pathToOffice) as alias
			on error
				try
					set startupDirectory to (folder "Word" of folder "Start" of folder "Office" of pathToOffice) as alias
				on error
					do shell script "mkdir -p " & (quoted form of POSIX path of pathToOffice) & "/Office/Startup/Word"
					set startupDirectory to (folder "Word" of folder "Startup" of folder "Office" of pathToOffice) as alias
				end try
			end try
		end tell
		
		-- Copy template to startup directory
		set AppleScript's text item delimiters to "/"
		tell application "Finder"
			do shell script "cp " & quoted form of (text items 1 thru -2 of (POSIX path of (path to me)) as text) & "/Zotero.dot " & (quoted form of POSIX path of startupDirectory)
			set dot to file "Zotero.dot" of startupDirectory
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