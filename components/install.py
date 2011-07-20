#!/usr/bin/python
# coding=utf-8

# Hack to fix pylib loading errors
import sys, os.path
pylib = os.path.realpath(os.path.dirname(__file__)+"/../pylib")
sys.path.insert(0, pylib)

from xpcom import components
import appscript, osax, mactypes, string, aem, os, subprocess, plistlib, shutil, re

CHOOSE_SCRIPT_MENU_ITEMS_STRING = "Select the Word Script Menu Items folder, usually "+			\
	"located in Documents/Microsoft User Data"
ERROR_WORD_X_TITLE = "Zotero Word Integration requires Word 2004 or later."
ERROR_WORD_X_STRING = "Please upgrade Microsoft Word to use Zotero Word Integration."
ERROR_NO_WORD_TITLE = "Neither Word 2004 nor Word 2008 appear to be installed on this computer."
ERROR_NO_WORD_STRING = "Please install Microsoft Word, then try again."
ERROR_WORD_RUNNING_TITLE = "Zotero MacWord Integration has been successfully installed, but Word must be restarted before it can be used."
ERROR_WORD_RUNNING_STRING = "Please restart Word before continuing."
NO_PERMISSIONS_TITLE = "Authentication Required"
NO_PERMISSIONS_STRING = "Zotero MacWord Integration requires administrator permissions to continue installation. Please enter your password at the next prompt."

SCRIPT_TEMPLATE = string.Template(r"""try
	do shell script "PIPE=\"/Users/Shared/.zoteroIntegrationPipe_$LOGNAME\";  if [ ! -e \"$PIPE\" ]; then PIPE='~/.zoteroIntegrationPipe'; fi; if [ -e \"$PIPE\" ]; then echo 'MacWord2008 ${command} '" & quoted form of POSIX path of (path to current application) & " > \"$PIPE\"; else exit 1; fi;"
on error
	display alert "Word could not communicate with Zotero. Please ensure that Firefox is open and try again." as critical
end try""")

W2008_SCRIPT_NAMES = ["Add Bibliography\\cob.scpt", "Add Citation\\coa.scpt",
	"Edit Bibliography\\cod.scpt", "Edit Citation\\coe.scpt", "Refresh\\cor.scpt",
	"Remove Field Codes.scpt", "Set Document Preferences\\cop.scpt"]
W2008_SCRIPT_COMMANDS = ["addBibliography", "addCitation", "editBibliography",
	"editCitation", "refresh", "removeCodes", "setDocPrefs"]

class Installer:
	_com_interfaces_ = [components.interfaces.zoteroIntegrationInstaller]
	_reg_clsid_ = "{aa56c6c0-95f0-48c2-b223-b11b96b9c9e5}"
	_reg_contractid_ = "@zotero.org/Zotero/integration/installer?agent=MacWord;1"
	_reg_desc_ = "Zotero MacWord Integration Installer"
	
	def __showError(self, title, error):
		components.classes["@mozilla.org/embedcomp/prompt-service;1"]			\
			.getService(components.interfaces.nsIPromptService)					\
			.alert(None, title, error)
	
	def __makeWordTemplate(self, template):
		try:
			import MacOS
			MacOS.SetCreatorAndType(template, 'MSWD', 'W8TN')
		except ImportError:
			# for forward compatibility, when Apple fixes the AppleScript bug and Python removes
			# the MacOS module
			templateAS = appscript.app(u'Finder').files[mactypes.Alias(template).hfspath]
			templateAS.creator_type.set('MSWD')
			templateAS.file_type.set('W8TN')
	
	def run(self, failSilently):
		osa = osax.OSAX(osaxname="StandardAdditions")
		applicationsFolder = osa.path_to(appscript.k.applications_folder)
		documentsFolder = osa.path_to(appscript.k.documents_folder)
		preferencesFolder = osa.path_to(appscript.k.preferences_folder)
		
		## See if we can find Office 2008
		# First look in the obvious place
		installed2008 = False
		wordInAppsDir = applicationsFolder.path+"/Microsoft Office 2008/Microsoft Word.app"
		installed2008 = os.path.isdir(wordInAppsDir)
		# Next query LS
		if not installed2008:
			wordPath = None
			try:
				wordPath = aem.findapp.byid('com.Microsoft.Word')
			except aem.findapp.ApplicationNotFoundError:
				pass
			
			# Check to make sure this is really Word >= 2008
			if wordPath:
				infoPlist = wordPath+"/Contents/Info.plist"
				if os.path.exists(infoPlist):
					infoPlistData = plistlib.readPlist(infoPlist)
					installed2008 = int(infoPlistData["CFBundleVersion"][0:2]) >= 12
		
		if installed2008:
			potentialMicrosoftUserDataNames = ["Microsoft User Data", "Microsoft-Benutzerdaten",
				u"Données Utilisateurs Microsoft", "Datos del Usuario de Microsoft",
				"Datos de Usuario de Microsoft", "Dati utente Microsoft"]
			foundMicrosoftUserData = False
			
			# Look for script menu items folder in various places
			for microsoftUserDataName in potentialMicrosoftUserDataNames:
				scriptMenuItemsFolder = documentsFolder.path+"/"+microsoftUserDataName+"/Word Script Menu Items"
				if os.path.isdir(scriptMenuItemsFolder):
					foundMicrosoftUserData = True
					break
				else:
					scriptMenuItemsFolder = preferencesFolder.path+"/"+microsoftUserDataName+"/Word Script Menu Items"
					if os.path.isdir(scriptMenuItemsFolder):
						foundMicrosoftUserData = True
						break
			
			# If we couldn't find a script menu items folder, ask the user to locate it
			if not foundMicrosoftUserData:
				try:
					scriptMenuItemsFolder = osa.choose_folder(
						with_prompt=CHOOSE_SCRIPT_MENU_ITEMS_STRING,
						default_location=documentsFolder).path
				except appscript.reference.CommandError:
					scriptMenuItemsFolder = None
		
			if scriptMenuItemsFolder:
				## Install the scripts
				scriptPath = scriptMenuItemsFolder+"/Zotero/"
				
				if os.path.exists(scriptPath):
					# Remove old scripts
					shutil.rmtree(scriptPath)
				# Create directory
				os.mkdir(scriptPath)
				
				# Generate scripts
				for i in range(len(W2008_SCRIPT_NAMES)):
					p = subprocess.Popen(['/usr/bin/osacompile', '-t', 'osas', '-c', 'ToyS', '-o', scriptPath+W2008_SCRIPT_NAMES[i]],
						stdin=subprocess.PIPE)
					p.stdin.write(SCRIPT_TEMPLATE.safe_substitute(command=W2008_SCRIPT_COMMANDS[i]))
					p.stdin.close()
					
		## See if we can find Office 2004
		# Fix template permissions
		template = os.path.realpath(os.path.dirname(__file__)+"/../install/Zotero.dot")
		self.__makeWordTemplate(template)
		
		# First look in the obvious place
		oldWord = False
		installed2004 = False
		wordPath = applicationsFolder.path+"/Microsoft Office 2004/Microsoft Word"
		installed2004 = os.path.exists(wordPath)
		if not installed2004:
			wordPath = False
			try:
				wordPath = aem.findapp.bycreator('MSWD')
			except aem.findapp.ApplicationNotFoundError:
				pass
			
			# Check to make sure this is really Word 2004
			if wordPath:
				appVersion = appscript.app(u'Finder').files[mactypes.Alias(wordPath).hfspath].version.get()
				if appVersion[0:2] == "11":
					installed2004 = True
				else:
					oldWord = True
		
		if installed2004:
			## Install the template
			# Get path to startup folder
			officeDir = os.path.dirname(wordPath)
			startupDir = officeDir+"/Office/Startup/Word"
			if not os.path.exists(startupDir):
				if os.path.exists(officeDir+"/Office/Start/Word"):
					startupDir = officeDir+"/Office/Start/Word"
				else:
					os.makedirs(startupDir)
			
			# Copy the template there
			newTemplate = startupDir+"/Zotero.dot"
			try:
				shutil.copy(template, newTemplate)				
				self.__makeWordTemplate(newTemplate)
			except IOError:
				components.classes["@mozilla.org/embedcomp/prompt-service;1"]			\
					.getService(components.interfaces.nsIPromptService)					\
					.alert(None, NO_PERMISSIONS_TITLE, NO_PERMISSIONS_STRING)
				osa.do_shell_script("ditto '"+template.replace("'", "'\\''")+"' '"+newTemplate.replace("'", "'\\''")+"'",
					administrator_privileges=True)
		
		if not installed2004 and not installed2008:
			if not failSilently:
				if oldWord:
					self.__showError(ERROR_WORD_X_TITLE, ERROR_WORD_X_STRING)
				else:
					self.__showError(ERROR_NO_WORD_TITLE, ERROR_NO_WORD_STRING)
		else:
			## Check whether Word is running
			p = subprocess.Popen(['/bin/ps', '-xo', 'command'], stdout=subprocess.PIPE)
			output = p.stdout.read()
			p.stdout.close()
			m = re.search(r'Microsoft Word', output)
			if m:
				self.__showError(ERROR_WORD_RUNNING_TITLE, ERROR_WORD_RUNNING_STRING)